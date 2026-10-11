"""管理各工作表的原始快照、工作副本、预览有效性和撤销历史。"""

from dataclasses import dataclass, field
from pathlib import Path

import pandas as pd

from app.core.cleaner import CleaningPreview, CleaningRules
from app.core.batch_cleaner import BatchCleaningPreview, BatchColumnChange, config_to_rules
from app.core.type_utils import values_equal
from app.core.annotation_builder import AnnotationColumn, AnnotationPreview, build_annotation_values, generate_annotation_name
from app.core.file_loader import ExcelPreview
from app.core.sheet_analyzer import SheetAnalysis
from app.models.cleaning_action import AnnotationAction, BatchCleaningAction, CleaningAction, HistoryAction, RestoreAction
from app.models.export_data import ExportSnapshot, SheetExportSnapshot


class DataStateError(Exception):
    """操作前提不满足或清洗预览已经失效。"""


@dataclass(frozen=True)
class SheetScanCache:
    revision: int
    result: SheetAnalysis


@dataclass(frozen=True)
class SheetScanSnapshot:
    generation: int
    sheet_name: str
    revision: int
    dataframe: pd.DataFrame


@dataclass
class SheetData:
    original_df: pd.DataFrame
    working_df: pd.DataFrame
    revision: int = 0
    operation_count: int = 0
    modified_columns: set[int] = field(default_factory=set)
    column_rules: dict[int, tuple[CleaningRules, ...]] = field(default_factory=dict)
    scan_cache: SheetScanCache | None = None
    annotations: tuple[AnnotationColumn, ...] = ()

    @property
    def is_modified(self) -> bool:
        return bool(self.modified_columns or self.annotations)


@dataclass(frozen=True)
class PreparedPreview:
    generation: int
    sheet_name: str
    column_index: int
    revision: int
    result: CleaningPreview
    before_values: pd.Series
    after_values: pd.Series


@dataclass(frozen=True)
class PreparedAnnotationPreview:
    generation: int
    sheet_name: str
    revision: int
    result: AnnotationPreview

    @property
    def column_index(self) -> int:
        return self.result.annotation.source_column_index


@dataclass(frozen=True)
class PreparedBatchPreview:
    generation: int
    sheet_name: str
    revision: int
    result: BatchCleaningPreview


class WorkbookDataManager:
    """只在内存中修改工作副本，原始快照不参与写入操作。"""

    def __init__(self) -> None:
        self.file_path: Path | None = None
        self.source_digest: str | None = None
        self.sheet_names: tuple[str, ...] = ()
        self.current_sheet: str | None = None
        self.sheet_data: dict[str, SheetData] = {}
        self.undo_stack: list[HistoryAction] = []
        self._generation = 0

    @staticmethod
    def _new_sheet(dataframe: pd.DataFrame) -> SheetData:
        original_df = dataframe.copy(deep=True)
        return SheetData(original_df, original_df.copy(deep=True))

    def open_workbook(self, preview: ExcelPreview) -> None:
        """成功打开文件后开始新会话；显式重开同一文件也清空旧历史。"""
        state = self._new_sheet(preview.dataframe)
        self.file_path = preview.file_path.resolve()
        self.source_digest = preview.source_digest
        self.sheet_names = preview.sheet_names
        self.current_sheet = preview.sheet_name
        self.sheet_data = {preview.sheet_name: state}
        self.undo_stack = []
        self._generation += 1

    def add_sheet(self, preview: ExcelPreview) -> None:
        if self.file_path != preview.file_path.resolve():
            raise DataStateError("工作簿已变化，请重新打开文件。")
        if preview.sheet_name not in self.sheet_names:
            raise DataStateError("当前工作簿中没有此工作表。")
        if (self.source_digest is not None and preview.source_digest is not None
                and self.source_digest != preview.source_digest):
            raise DataStateError("源文件在打开后发生变化，请重新打开文件后再操作。")
        if preview.sheet_name not in self.sheet_data:
            self.sheet_data[preview.sheet_name] = self._new_sheet(preview.dataframe)
        self.current_sheet = preview.sheet_name

    def select_sheet(self, sheet_name: str) -> ExcelPreview:
        if sheet_name not in self.sheet_data or self.file_path is None:
            raise DataStateError("此工作表尚未载入。")
        self.current_sheet = sheet_name
        return ExcelPreview(
            self.file_path, self.sheet_names, sheet_name,
            self.sheet_data[sheet_name].original_df,
            self.source_digest,
        )

    @property
    def export_state_token(self) -> tuple:
        """确认框返回后检查会话和各 Sheet 版本，防止导出过期快照。"""
        return self._generation, tuple((name, state.revision) for name, state in self.sheet_data.items())

    def get_export_snapshot(self) -> ExportSnapshot:
        if self.file_path is None:
            raise DataStateError("请先打开 Excel 文件。")
        sheets = tuple(
            SheetExportSnapshot(name, state.original_df.copy(deep=True), state.working_df.copy(deep=True),
                                dict(state.column_rules), tuple(state.annotations))
            for name, state in self.sheet_data.items()
        )
        return ExportSnapshot(self.file_path, self.source_digest, sheets, self.export_state_token)

    @property
    def current_data(self) -> SheetData:
        if self.current_sheet is None:
            raise DataStateError("请先打开 Excel 文件。")
        return self.sheet_data[self.current_sheet]

    @property
    def original_df(self) -> pd.DataFrame:
        return self.current_data.original_df

    @property
    def working_df(self) -> pd.DataFrame:
        return self.current_data.working_df

    @property
    def is_modified(self) -> bool:
        return any(state.is_modified for state in self.sheet_data.values())

    def prepare_sheet_scan(self) -> SheetScanSnapshot:
        """复制轻量容器，不在 UI 线程深拷贝整表；检测只读，修改总是替换工作副本。"""
        state = self.current_data
        return SheetScanSnapshot(
            self._generation, self.current_sheet, state.revision,
            state.working_df.copy(deep=False),
        )

    def is_sheet_scan_current(self, snapshot: SheetScanSnapshot) -> bool:
        return (
            snapshot.generation == self._generation
            and snapshot.sheet_name == self.current_sheet
            and snapshot.sheet_name in self.sheet_data
            and snapshot.revision == self.sheet_data[snapshot.sheet_name].revision
        )

    def save_sheet_scan(self, snapshot: SheetScanSnapshot, result: SheetAnalysis) -> None:
        if not self.is_sheet_scan_current(snapshot):
            raise DataStateError("工作表数据已变化，请重新扫描。")
        if (result.row_count != len(snapshot.dataframe)
                or result.column_count != snapshot.dataframe.shape[1]):
            raise DataStateError("扫描结果与当前工作表不匹配，请重新扫描。")
        self.current_data.scan_cache = SheetScanCache(snapshot.revision, result)

    def prepare_annotation_preview(self, result: AnnotationPreview) -> PreparedAnnotationPreview:
        state = self.current_data
        spec = result.annotation
        if not 0 <= spec.source_column_index < state.original_df.shape[1]:
            raise DataStateError("请选择原始字段，不能在新增辅助列上继续生成辅助列。")
        source = state.working_df.iloc[:, spec.source_column_index]
        if (not source.equals(result.source_values) or result.source_name != str(source.name)
                or not source.index.equals(result.values.index) or len(source) != len(result.values)
                or spec.column_name != generate_annotation_name(str(source.name), state.working_df.columns, spec.kind)):
            raise DataStateError("辅助列预览与当前数据不匹配，请重新生成。")
        expected = build_annotation_values(source, spec.kind, date_confirmed=spec.date_confirmed)
        expected.name = spec.column_name
        if not expected.equals(result.values):
            raise DataStateError("辅助列结果与当前字段不一致，请重新生成预览。")
        return PreparedAnnotationPreview(self._generation, self.current_sheet, state.revision, result)

    def apply_annotation(self, preview: PreparedAnnotationPreview) -> AnnotationAction:
        if not self.is_preview_current(preview):
            raise DataStateError("辅助列预览已失效，请重新生成。")
        # 确认返回后重新检查源值、列名和标签，防止添加过期或被改动的预览。
        self.prepare_annotation_preview(preview.result)
        state = self.current_data
        result = preview.result
        if state.working_df.shape[1] >= 16384:
            raise DataStateError("工作表已达到 Excel 最大列数，不能新增辅助列。")
        action = AnnotationAction(
            self.current_sheet, state.working_df.shape[1],
            result.annotation.column_name, result.annotation.source_column_index,
            state.operation_count,
        )
        updated = state.working_df.copy(deep=True)
        updated.insert(len(updated.columns), action.column_name, result.values.copy(deep=True))
        annotations = state.annotations + (result.annotation,)
        history = self.undo_stack + [action]
        # 分配成功后一起提交；原始快照及原有字段不赋值。
        state.working_df = updated
        state.annotations = annotations
        state.operation_count += 1
        state.revision += 1
        self.undo_stack = history
        return action

    @staticmethod
    def _refresh_annotations(
        dataframe: pd.DataFrame, annotations: tuple[AnnotationColumn, ...],
        source_column: int | None = None,
    ) -> None:
        """仅修改尚未提交的候选副本，使标签随源列清洗/撤销同步变化。"""
        first_annotation = dataframe.shape[1] - len(annotations)
        for position, spec in enumerate(annotations):
            if source_column is None or spec.source_column_index == source_column:
                values = build_annotation_values(dataframe.iloc[:, spec.source_column_index], spec.kind,
                                                 date_confirmed=spec.date_confirmed)
                dataframe.isetitem(first_annotation + position, values)

    def prepare_preview(self, column_index: int, result: CleaningPreview) -> PreparedPreview:
        state = self.current_data
        if not 0 <= column_index < state.working_df.shape[1]:
            raise DataStateError("请先选择有效的数据列。")
        if column_index >= state.original_df.shape[1]:
            raise DataStateError("新增辅助列不能直接清洗，请选择原始字段。")
        before = state.working_df.iloc[:, column_index].copy(deep=True)
        after = result.preview_series.copy(deep=True)
        if (result.column_name != str(before.name) or result.row_count != len(before)
                or len(result.records) != len(before) or not after.index.equals(before.index)
                or (after.dtype != before.dtype and not (result.rules.numeric_conversion or result.rules.date_normalize))
                or result.changed_count != sum(record.changed for record in result.records)):
            raise DataStateError("预览与当前列不匹配，请重新生成预览。")
        source = pd.Series(
            [record.original_value for record in result.records],
            index=before.index, name=before.name, dtype=before.dtype,
        )
        displayed = pd.Series(
            [record.cleaned_value for record in result.records],
            index=after.index, name=after.name, dtype=after.dtype,
        )
        if not before.equals(source) or not after.equals(displayed):
            raise DataStateError("数据与预览内容不一致，请重新生成预览。")
        return PreparedPreview(
            self._generation, self.current_sheet, column_index,
            state.revision, result, before, after,
        )

    def is_preview_current(self, preview: PreparedPreview | PreparedAnnotationPreview | PreparedBatchPreview) -> bool:
        return (
            preview.generation == self._generation
            and preview.sheet_name == self.current_sheet
            and preview.sheet_name in self.sheet_data
            and preview.revision == self.sheet_data[preview.sheet_name].revision
        )

    def apply_preview(self, preview: PreparedPreview) -> CleaningAction | None:
        if not self.is_preview_current(preview):
            raise DataStateError("清洗预览已失效，请重新生成后再应用。")
        state = self.current_data
        if not state.working_df.iloc[:, preview.column_index].equals(preview.before_values):
            raise DataStateError("当前列已变化，请重新生成预览。")
        affected = tuple(index for index, record in enumerate(preview.result.records) if record.changed)
        if not affected:
            return None
        action = CleaningAction(
            preview.sheet_name, preview.column_index, preview.result.column_name,
            preview.before_values.copy(deep=True), preview.after_values.copy(deep=True),
            affected, preview.result.rules,
            previous_rules=state.column_rules.get(preview.column_index, ()),
        )
        updated = self._replace_column(state.working_df, action.column_index, action.after_values)
        self._refresh_annotations(updated, state.annotations, action.column_index)
        modified_columns = self._modified_columns_after(state, action.column_index, updated)
        history = self.undo_stack + [action]
        column_rules = dict(state.column_rules)
        column_rules[action.column_index] = action.previous_rules + (action.rules,)
        # 大块快照和历史分配完成后才提交，失败时旧状态仍可使用。
        state.working_df = updated
        state.operation_count += 1
        state.revision += 1
        state.modified_columns = modified_columns
        state.column_rules = column_rules
        self.undo_stack = history
        return action

    def prepare_batch_preview(
        self, snapshot: SheetScanSnapshot, result: BatchCleaningPreview,
    ) -> PreparedBatchPreview:
        if not self.is_sheet_scan_current(snapshot):
            raise DataStateError("批量预览已失效，请重新生成。")
        state = self.current_data
        if result.sheet_name != self.current_sheet or result.row_count != len(state.working_df):
            raise DataStateError("批量预览与当前工作表不匹配。")
        configs = {config.column_index: config for config in result.configs}
        if len(configs) != len(result.configs) or not configs:
            raise DataStateError("批量字段配置不完整，请重新生成。")
        for index, config in configs.items():
            if (not config.enabled or config.is_generated or not 0 <= index < state.original_df.shape[1]
                    or str(state.working_df.columns[index]) != config.column_name):
                raise DataStateError("批量字段配置已变化，请重新生成。")
            config_to_rules(config)
        seen, rows = set(), set()
        for change in result.column_changes:
            index = change.column_index
            if index in seen or index not in configs:
                raise DataStateError("批量预览包含重复或未知字段。")
            seen.add(index)
            before, after = change.before_values, change.after_values
            current = state.working_df.iloc[:, index]
            if (not current.equals(before) or len(after) != len(before) or not before.index.equals(after.index)
                    or before.name != after.name or change.column_name != str(current.name)
                    or change.rules != config_to_rules(configs[index])):
                raise DataStateError("批量预览与当前字段内容不一致，请重新生成。")
            affected = tuple(position for position, (old, new) in enumerate(zip(before, after))
                             if not values_equal(old, new))
            if affected != change.affected_indices or not affected:
                raise DataStateError("批量预览变化统计不一致，请重新生成。")
            rows.update(position + 2 for position in affected)
        if tuple(sorted(rows)) != result.affected_rows:
            raise DataStateError("批量预览行数统计不一致，请重新生成。")
        return PreparedBatchPreview(self._generation, self.current_sheet, state.revision, result)

    def apply_batch_preview(self, preview: PreparedBatchPreview) -> BatchCleaningAction | None:
        """所有列和辅助标注先写入一个候选副本，成功后整体提交一次历史。"""
        if not self.is_preview_current(preview):
            raise DataStateError("批量预览已失效，请重新生成后再应用。")
        self.prepare_batch_preview(self.prepare_sheet_scan(), preview.result)
        state = self.current_data
        if not preview.result.column_changes:
            return None
        changes = tuple(BatchColumnChange(
            change.column_index, change.column_name, change.before_values.copy(deep=True),
            change.after_values.copy(deep=True), change.affected_indices, change.rules,
        ) for change in preview.result.column_changes)
        action = BatchCleaningAction(
            preview.sheet_name, changes, state.operation_count,
            frozenset(state.modified_columns), dict(state.column_rules),
        )
        updated = state.working_df.copy(deep=True)
        modified_columns = set(state.modified_columns)
        column_rules = dict(state.column_rules)
        for change in changes:
            updated.isetitem(change.column_index, change.after_values.copy(deep=True))
            if updated.iloc[:, change.column_index].equals(state.original_df.iloc[:, change.column_index]):
                modified_columns.discard(change.column_index)
            else:
                modified_columns.add(change.column_index)
            column_rules[change.column_index] = column_rules.get(change.column_index, ()) + (change.rules,)
        self._refresh_annotations(updated, state.annotations)
        history = self.undo_stack + [action]
        state.working_df = updated
        state.modified_columns = modified_columns
        state.column_rules = column_rules
        state.operation_count += 1
        state.revision += 1
        self.undo_stack = history
        return action

    @staticmethod
    def _replace_column(dataframe: pd.DataFrame, column_index: int, values: pd.Series) -> pd.DataFrame:
        # 先构造完整候选副本，赋值失败时保留原有工作数据及撤销历史。
        updated = dataframe.copy(deep=True)
        updated.isetitem(column_index, values.copy(deep=True))
        return updated

    @staticmethod
    def _modified_columns_after(state: SheetData, column_index: int, updated: pd.DataFrame) -> set[int]:
        columns = set(state.modified_columns)
        if updated.iloc[:, column_index].equals(state.original_df.iloc[:, column_index]):
            columns.discard(column_index)
        else:
            columns.add(column_index)
        return columns

    def undo(self) -> HistoryAction | None:
        if not self.undo_stack:
            return None
        action = self.undo_stack[-1]
        state = self.sheet_data[action.sheet_name]
        annotations = state.annotations
        if isinstance(action, AnnotationAction):
            if (action.column_index != state.working_df.shape[1] - 1
                    or not annotations or annotations[-1].column_name != action.column_name):
                raise DataStateError("辅助列结构已变化，无法安全撤销。")
            updated = state.working_df.iloc[:, :-1].copy(deep=True)
            annotations = annotations[:-1]
            count = action.previous_operation_count
            modified_columns = set(state.modified_columns)
            column_rules = dict(state.column_rules)
        elif isinstance(action, BatchCleaningAction):
            updated = state.working_df.copy(deep=True)
            for change in action.column_changes:
                updated.isetitem(change.column_index, change.before_values.copy(deep=True))
            self._refresh_annotations(updated, annotations)
            count = action.previous_operation_count
            modified_columns = set(action.previous_modified_columns)
            column_rules = dict(action.previous_column_rules)
        elif isinstance(action, CleaningAction):
            updated = self._replace_column(state.working_df, action.column_index, action.before_values)
            self._refresh_annotations(updated, annotations, action.column_index)
            count = state.operation_count - 1
            modified_columns = self._modified_columns_after(state, action.column_index, updated)
            column_rules = dict(state.column_rules)
            if action.previous_rules:
                column_rules[action.column_index] = action.previous_rules
            else:
                column_rules.pop(action.column_index, None)
        else:
            updated = action.before_values.copy(deep=True)
            count = action.previous_operation_count
            modified_columns = set(action.previous_modified_columns)
            column_rules = dict(action.previous_column_rules)
            annotations = action.previous_annotations
        history = self.undo_stack[:-1]
        state.working_df = updated
        state.annotations = annotations
        state.operation_count = count
        state.modified_columns = modified_columns
        state.column_rules = column_rules
        state.revision += 1
        self.undo_stack = history
        self.current_sheet = action.sheet_name
        return action

    def restore_original(self) -> RestoreAction | None:
        state = self.current_data
        if not state.is_modified:
            return None
        action = RestoreAction(
            self.current_sheet, state.working_df.copy(deep=True),
            state.operation_count, frozenset(state.modified_columns),
            previous_column_rules=dict(state.column_rules),
            previous_annotations=state.annotations,
        )
        restored = state.original_df.copy(deep=True)
        history = self.undo_stack + [action]
        state.working_df = restored
        state.annotations = ()
        state.operation_count = 0
        state.modified_columns.clear()
        state.column_rules = {}
        state.revision += 1
        self.undo_stack = history
        return action
