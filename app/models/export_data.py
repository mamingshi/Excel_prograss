"""导出任务的独立快照、差异记录和结果，不包含界面或文件写入逻辑。"""

from dataclasses import dataclass, field
from enum import StrEnum
from pathlib import Path

import pandas as pd
from app.core.cleaner import CleaningRules
from app.core.annotation_builder import AnnotationColumn


class ExportMode(StrEnum):
    UPDATE_COLUMNS = "update_columns"
    ADD_FORMULA_COLUMNS = "add_formula_columns"

    @property
    def label(self) -> str:
        return "原列更新" if self is ExportMode.UPDATE_COLUMNS else "新增对照列（公式）"


@dataclass(frozen=True)
class SheetExportSnapshot:
    sheet_name: str
    original_df: pd.DataFrame
    working_df: pd.DataFrame
    column_rules: dict[int, tuple[CleaningRules, ...]] = field(default_factory=dict)
    annotations: tuple[AnnotationColumn, ...] = ()


@dataclass(frozen=True)
class ExportSnapshot:
    source_path: Path
    source_digest: str | None
    sheets: tuple[SheetExportSnapshot, ...]
    state_token: tuple


@dataclass(frozen=True)
class CellChange:
    sheet_name: str
    excel_row: int
    excel_column: int
    column_name: str
    old_value: object
    new_value: object


@dataclass(frozen=True)
class ExportPlan:
    snapshot: ExportSnapshot
    changes: tuple[CellChange, ...]
    mode: ExportMode = ExportMode.UPDATE_COLUMNS

    @property
    def sheet_count(self) -> int:
        return len({change.sheet_name for change in self.changes}
                   | {sheet.sheet_name for sheet in self.snapshot.sheets if sheet.annotations})

    @property
    def annotation_column_count(self) -> int:
        return sum(len(sheet.annotations) for sheet in self.snapshot.sheets)

    @property
    def cell_count(self) -> int:
        return len(self.changes)

    @property
    def column_count(self) -> int:
        return len({(change.sheet_name, change.excel_column) for change in self.changes})


@dataclass(frozen=True)
class ExportResult:
    target_path: Path
    sheet_count: int
    cell_count: int
    mode: ExportMode = ExportMode.UPDATE_COLUMNS
    added_column_count: int = 0
    annotation_column_count: int = 0
