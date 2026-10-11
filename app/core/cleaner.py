"""字符与安全类型清洗及内存预览；不修改输入，不操作文件或 Qt。"""

from dataclasses import dataclass
from typing import Any, Literal

import pandas as pd

from app.core.type_cleaner import fullwidth_to_halfwidth, numeric_conversion, date_conversion


HiddenStrategy = Literal["space", "delete"]
# 清洗和参考公式共用顺序定义，首尾普通空格始终最后处理。
RULE_ORDER = ("nbsp", "tab", "newline", "carriage_return", "outer_spaces")
TYPE_RULE_ORDER = ("fullwidth", "numeric_conversion", "date_normalize")
# 单列与批量共用：先处理隐藏字符，再转换全角，最后去首尾空格和转换类型。
ALL_RULE_ORDER = RULE_ORDER[:-1] + ("fullwidth", "outer_spaces", "numeric_conversion", "date_normalize")
RULE_LABELS = {
    "nbsp": "CHAR(160) → 普通空格",
    "tab": "Tab → 普通空格",
    "newline": "换行（\\n）→ 普通空格",
    "carriage_return": "回车（\\r）→ 普通空格",
    "outer_spaces": "去除前后普通空格",
    "fullwidth": "全角转半角",
    "numeric_conversion": "文本数字转数字",
    "date_normalize": "统一日期格式（YYYY-MM-DD）",
}


@dataclass(frozen=True)
class CleaningRules:
    nbsp: bool = False
    tab: bool = False
    newline: bool = False
    carriage_return: bool = False
    outer_spaces: bool = False
    hidden_strategy: HiddenStrategy = "space"
    fullwidth: bool = False
    numeric_conversion: bool = False
    date_normalize: bool = False
    date_confirmed: bool = False

    def __post_init__(self) -> None:
        _validate_strategy(self.hidden_strategy)
        if self.date_normalize and not self.date_confirmed:
            raise ValueError("请先明确确认当前字段是日期列，再生成日期标准化预览。")

    @property
    def has_type_rules(self) -> bool:
        return self.fullwidth or self.numeric_conversion or self.date_normalize


@dataclass(frozen=True)
class CleaningRecord:
    excel_row: int
    original_value: Any
    cleaned_value: Any
    operations: tuple[str, ...]
    changed: bool
    notes: tuple[str, ...] = ()


@dataclass(frozen=True)
class CleaningPreview:
    column_name: str
    row_count: int
    changed_count: int
    unchanged_count: int
    processed_counts: dict[str, int]
    records: tuple[CleaningRecord, ...]
    preview_series: pd.Series
    rules: CleaningRules


def _validate_strategy(strategy: str) -> None:
    if strategy not in ("space", "delete"):
        raise ValueError("隐藏字符处理策略仅支持 space（替换为空格）或 delete（直接删除）。")


def clean_outer_spaces(value: Any) -> Any:
    """仅删除首尾 U+0020；保留中间普通空格和其他空白字符。"""
    return value.strip(" ") if isinstance(value, str) else value


def clean_nbsp(value: Any) -> Any:
    """将 U+00A0 替换为普通空格，保留词语之间的分隔。"""
    return value.replace("\u00a0", " ") if isinstance(value, str) else value


def _clean_character(value: Any, character: str, strategy: HiddenStrategy) -> Any:
    _validate_strategy(strategy)
    if not isinstance(value, str):
        return value
    return value.replace(character, " " if strategy == "space" else "")


def clean_tab(value: Any, strategy: HiddenStrategy = "space") -> Any:
    return _clean_character(value, "\t", strategy)


def clean_newline(value: Any, strategy: HiddenStrategy = "space") -> Any:
    return _clean_character(value, "\n", strategy)


def clean_carriage_return(value: Any, strategy: HiddenStrategy = "space") -> Any:
    return _clean_character(value, "\r", strategy)


def clean_hidden_whitespace(value: Any, strategy: HiddenStrategy = "space") -> Any:
    """按固定顺序组合处理 Tab、换行和回车。"""
    return clean_carriage_return(clean_newline(clean_tab(value, strategy), strategy), strategy)


def rule_label(rule: str, strategy: HiddenStrategy = "space") -> str:
    if strategy == "delete" and rule in ("tab", "newline", "carriage_return"):
        return {"tab": "删除 Tab", "newline": "删除换行（\\n）", "carriage_return": "删除回车（\\r）"}[rule]
    return RULE_LABELS[rule]


_RULE_FUNCTIONS = {
    "nbsp": clean_nbsp,
    "tab": clean_tab,
    "newline": clean_newline,
    "carriage_return": clean_carriage_return,
    "outer_spaces": clean_outer_spaces,
    "fullwidth": fullwidth_to_halfwidth,
}


def _apply_with_operations(value: Any, rules: CleaningRules) -> tuple[Any, tuple[str, ...], tuple[str, ...]]:
    if not isinstance(value, str) and not rules.date_normalize:
        return value, (), ()
    applied, notes = [], []
    for rule in ALL_RULE_ORDER:
        if not getattr(rules, rule):
            continue
        note = None
        if rule == "numeric_conversion":
            result, note = numeric_conversion(value)
        elif rule == "date_normalize":
            result, note = date_conversion(value)
        else:
            function = _RULE_FUNCTIONS[rule]
            result = (function(value, rules.hidden_strategy)
                      if rule in ("tab", "newline", "carriage_return") else function(value))
        if note is not None:
            notes.append(note)
        # 类型转换也算变化；缺失值函数均返回原对象，不用 NaN != NaN 判断。
        if result is not value and (type(result) is not type(value) or result != value):
            applied.append(rule)
        value = result
    return value, tuple(applied), tuple(notes)


def apply_cleaning_rules(value: Any, rules: CleaningRules) -> Any:
    """仅执行勾选规则；非字符串仅在明确确认的日期标准化中参与转换。"""
    return _apply_with_operations(value, rules)[0]


def generate_cleaning_preview(
    column: pd.Series, rules: CleaningRules, first_excel_row: int = 2
) -> CleaningPreview:
    """创建独立 Series 和逐行对比，记录真正发生变化的操作。"""
    if first_excel_row < 1:
        raise ValueError("Excel 起始行号必须大于或等于 1。")
    preview_series = (column.astype(object).copy(deep=True)
                      if rules.numeric_conversion or rules.date_normalize else column.copy(deep=True))
    records = []
    changed_count = 0
    processed_counts = {"前后空格": 0, "CHAR(160)": 0, "Tab / 换行": 0}
    for rule in TYPE_RULE_ORDER:
        if getattr(rules, rule):
            processed_counts[RULE_LABELS[rule]] = 0
    for position, original in enumerate(column):
        cleaned, applied, notes = _apply_with_operations(original, rules)
        changed = bool(applied)
        if changed:
            preview_series.iloc[position] = cleaned
            changed_count += 1
        if "outer_spaces" in applied:
            processed_counts["前后空格"] += 1
        if "nbsp" in applied:
            processed_counts["CHAR(160)"] += 1
        if any(rule in applied for rule in ("tab", "newline", "carriage_return")):
            processed_counts["Tab / 换行"] += 1
        for rule in TYPE_RULE_ORDER:
            if rule in applied:
                processed_counts[RULE_LABELS[rule]] += 1
        records.append(CleaningRecord(
            excel_row=position + first_excel_row,
            original_value=original,
            cleaned_value=cleaned,
            operations=tuple(rule_label(rule, rules.hidden_strategy) for rule in applied),
            changed=changed, notes=notes,
        ))
    return CleaningPreview(
        column_name=str(column.name) if column.name is not None else "未命名列",
        row_count=len(column),
        changed_count=changed_count,
        unchanged_count=len(column) - changed_count,
        processed_counts=processed_counts,
        records=tuple(records),
        preview_series=preview_series,
        rules=rules,
    )
