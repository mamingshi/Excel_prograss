"""多列只读预览、推荐及风险检查；清洗算法复用单列入口。"""

from collections.abc import Callable, Sequence
from dataclasses import dataclass, replace

import pandas as pd

from app.core.cleaner import (
    ALL_RULE_ORDER, CleaningRules, apply_cleaning_rules, generate_cleaning_preview, rule_label,
)
from app.core.sheet_analyzer import ScanCancelled, SheetAnalysis
from app.core.type_utils import has_leading_zero
from app.models.cleaning_rule_config import ColumnCleaningConfig


BATCH_RULE_ORDER = ALL_RULE_ORDER
PREVIEW_LIMIT = 1000


class BatchRuleError(ValueError):
    """配置冲突或存在不能自动处理的风险。"""


@dataclass(frozen=True)
class BatchPreviewRecord:
    sheet_name: str
    column_name: str
    column_index: int
    excel_row: int
    original_value: object
    cleaned_value: object
    rules: str
    changed: bool
    notes: tuple[str, ...] = ()


@dataclass(frozen=True)
class BatchColumnChange:
    column_index: int
    column_name: str
    before_values: pd.Series
    after_values: pd.Series
    affected_indices: tuple[int, ...]
    rules: CleaningRules


@dataclass(frozen=True)
class BatchCleaningPreview:
    sheet_name: str
    row_count: int
    configs: tuple[ColumnCleaningConfig, ...]
    column_changes: tuple[BatchColumnChange, ...]
    affected_rows: tuple[int, ...]
    changed_records: tuple[BatchPreviewRecord, ...]
    all_records: tuple[BatchPreviewRecord, ...]
    protected_count: int
    preview_limit: int = PREVIEW_LIMIT

    @property
    def changed_cell_count(self) -> int:
        return sum(len(column.affected_indices) for column in self.column_changes)

    @property
    def changed_row_count(self) -> int:
        return len(self.affected_rows)

    @property
    def column_counts(self) -> dict[int, int]:
        changed = {column.column_index: len(column.affected_indices) for column in self.column_changes}
        return {config.column_index: changed.get(config.column_index, 0) for config in self.configs}


def recommend_batch_configs(
    columns: Sequence, analysis: SheetAnalysis | None, *, original_column_count: int,
) -> tuple[ColumnCleaningConfig, ...]:
    """复用整表统计生成每列独立建议；重复与空值只列入人工判断提示。"""
    qualities = {item.column_index: item for item in analysis.columns} if analysis else {}
    configs = []
    for index, name in enumerate(columns):
        quality = qualities.get(index)
        generated = index >= original_column_count
        rules, warnings, problems = [], [], []
        if quality is not None:
            for attribute, label in (
                ("outer_spaces", "前后空格"), ("nbsp", "CHAR(160)"), ("hidden_chars", "Tab / 换行"),
                ("empty", "空值"), ("duplicates", "重复值"), ("numeric_mixture", "文本数字混合"),
                ("leading_zero", "前导0风险"), ("date_anomalies", "日期异常"), ("fullwidth", "全角字符"),
            ):
                count = getattr(quality, attribute)
                if count:
                    problems.append(f"{label} {count} 条")
            if quality.outer_spaces:
                rules.append("outer_spaces")
            if quality.nbsp:
                rules.extend(("nbsp", "outer_spaces"))
            if quality.hidden_chars:
                rules.extend(("tab", "newline", "carriage_return"))
            if quality.fullwidth:
                rules.append("fullwidth")
            if quality.numeric_mixture and not quality.leading_zero:
                rules.append("numeric_conversion")
            if quality.date_anomalies:
                rules.append("date_normalize")
                warnings.append("日期标准化需确认字段含义；无法识别、歧义及含时间的值保留。")
            if quality.leading_zero:
                warnings.append("存在前导0风险，不建议自动转数字；本版禁止该列转数字。")
            if quality.duplicates:
                warnings.append(f"重复值 {quality.duplicates} 条：建议生成辅助列，由人工判断。")
            if quality.empty:
                warnings.append(f"空值 {quality.empty} 条：仅提示，不填充或删除。")
        if generated:
            rules = []
            warnings.append("软件生成的辅助列只作标注，不能直接批量清洗。")
        ordered = tuple(rule for rule in BATCH_RULE_ORDER if rule in rules)
        configs.append(ColumnCleaningConfig(
            index, str(name), bool(ordered), ordered, tuple(warnings),
            "；".join(problems) or ("未发现当前支持的问题" if analysis else "尚未扫描"),
            is_generated=generated,
        ))
    return tuple(configs)


def config_to_rules(config: ColumnCleaningConfig) -> CleaningRules:
    unknown = set(config.rules) - set(BATCH_RULE_ORDER)
    if unknown:
        raise BatchRuleError(f"{config.column_name}：不支持的清洗规则：{', '.join(sorted(unknown))}")
    if not config.rules:
        raise BatchRuleError(f"{config.column_name}：请选择至少一项清洗规则，或取消处理该列。")
    if "numeric_conversion" in config.rules and "date_normalize" in config.rules:
        raise BatchRuleError(f"{config.column_name}：文本数字转换与日期标准化不能同时启用。")
    try:
        return CleaningRules(**{rule: True for rule in config.rules}, date_confirmed=config.date_confirmed)
    except ValueError as error:
        raise BatchRuleError(f"{config.column_name}：{error}") from error


def rules_description(config: ColumnCleaningConfig) -> str:
    return "；".join(rule_label(rule) for rule in BATCH_RULE_ORDER if rule in config.rules) or "未选择规则"


def generate_batch_preview(
    dataframe: pd.DataFrame, configs: Sequence[ColumnCleaningConfig], sheet_name: str, *,
    original_column_count: int | None = None, preview_limit: int = PREVIEW_LIMIT,
    progress: Callable[[int, int, str], None] | None = None,
    is_cancelled: Callable[[], bool] | None = None,
) -> BatchCleaningPreview:
    """仅复制目标 Series；界面样本限量，完整变化仍保留用于应用与撤销。"""
    if preview_limit < 1:
        raise ValueError("预览上限必须大于0。")
    active = tuple(config for config in configs if config.enabled)
    if not active:
        raise BatchRuleError("请至少启用一个字段，并为它选择清洗规则。")
    width = dataframe.shape[1] if original_column_count is None else original_column_count
    seen, prepared = set(), []
    for config in active:
        if config.column_index in seen:
            raise BatchRuleError("同一个字段不能重复配置，请重新生成方案。")
        seen.add(config.column_index)
        if (config.is_generated or not 0 <= config.column_index < width
                or config.column_index >= dataframe.shape[1]
                or str(dataframe.columns[config.column_index]) != config.column_name):
            raise BatchRuleError(f"{config.column_name}：字段已变化或属于辅助列，请重新配置。")
        prepared.append((config, config_to_rules(config)))
    changed_records, all_records, changes, affected_rows = [], [], [], set()
    protected_count = 0
    for completed, (config, rules) in enumerate(prepared, 1):
        if is_cancelled and is_cancelled():
            raise ScanCancelled()
        column = dataframe.iloc[:, config.column_index]
        if rules.numeric_conversion:
            # 检查原始值和前置字符规则后的值，避免全角或空白掩盖前导0。
            preceding = replace(rules, numeric_conversion=False, date_normalize=False)
            if any(has_leading_zero(value) or has_leading_zero(apply_cleaning_rules(value, preceding))
                   for value in column):
                raise BatchRuleError(
                    f"{config.column_name}列存在前导0数据，转数字后可能丢失前导0。"
                    "本版禁止继续，请取消该列的“文本数字转数字”规则。"
                )
        preview = generate_cleaning_preview(column, rules)
        affected = tuple(index for index, record in enumerate(preview.records) if record.changed)
        if affected:
            changes.append(BatchColumnChange(
                config.column_index, config.column_name, column.copy(deep=True),
                preview.preview_series, affected, rules,
            ))
            affected_rows.update(index + 2 for index in affected)
        rule_text = rules_description(config)
        for record in preview.records:
            protected_count += bool(record.notes)
            if len(all_records) >= preview_limit and (not record.changed or len(changed_records) >= preview_limit):
                continue
            sample = BatchPreviewRecord(
                sheet_name, config.column_name, config.column_index, record.excel_row,
                record.original_value, record.cleaned_value, rule_text, record.changed, record.notes,
            )
            if len(all_records) < preview_limit:
                all_records.append(sample)
            if record.changed and len(changed_records) < preview_limit:
                changed_records.append(sample)
        if progress:
            progress(completed, len(prepared), config.column_name)
    if is_cancelled and is_cancelled():
        raise ScanCancelled()
    return BatchCleaningPreview(
        sheet_name, len(dataframe), active, tuple(changes), tuple(sorted(affected_rows)),
        tuple(changed_records), tuple(all_records), protected_count, preview_limit,
    )
