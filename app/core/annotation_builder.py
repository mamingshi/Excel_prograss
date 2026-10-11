"""异常标注：只生成重复检查标签、兼容公式和预览，不修改源数据。"""

from dataclasses import dataclass
from collections.abc import Iterable
from datetime import date, datetime, time, timedelta
from numbers import Number

from openpyxl.utils import get_column_letter
from openpyxl.utils.datetime import CALENDAR_MAC_1904, to_excel
import pandas as pd

from app.core.detector import detect_duplicates, is_empty_value
from app.core.type_utils import (
    numeric_type_label, has_leading_zero, parse_date_value, summarize_dates, date_check_label,
)


ANNOTATION_SUFFIXES = {"duplicate": "重复检查", "numeric_type": "类型检查",
                       "leading_zero": "前导0风险", "date": "日期检查"}


@dataclass(frozen=True)
class AnnotationColumn:
    column_name: str
    source_column_index: int
    kind: str = "duplicate"
    date_confirmed: bool = False

    def __post_init__(self) -> None:
        if self.kind not in ANNOTATION_SUFFIXES:
            raise ValueError("不支持的辅助列类型。")


@dataclass(frozen=True)
class AnnotationPreview:
    annotation: AnnotationColumn
    source_name: str
    source_values: pd.Series
    values: pd.Series

    @property
    def row_count(self) -> int:
        return len(self.source_values)

    @property
    def duplicate_count(self) -> int:
        return int((self.values == "重复").sum())

    @property
    def empty_count(self) -> int:
        return int((self.values == "空值").sum())

    @property
    def problem_count(self) -> int:
        labels = {"duplicate": {"重复"}, "numeric_type": {"文本数字", "文本数字（前导0）"},
                  "leading_zero": {"前导0"}, "date": {"格式不统一", "无法识别", "日期歧义", "含时间（保留）"}}
        return int(self.values.isin(labels[self.annotation.kind]).sum())

    def formula(self, row_position: int) -> str:
        if not 0 <= row_position < self.row_count:
            raise ValueError("预览行位置超出范围。")
        if self.annotation.kind != "duplicate":
            return "检测结果（导出为值）"
        return build_duplicate_formula(
            self.annotation.source_column_index, row_position + 2, self.row_count + 1,
        )


def generate_annotation_name(source_name: str, existing_columns: Iterable, kind: str = "duplicate") -> str:
    used = {str(name).casefold() for name in existing_columns}
    base = f"{source_name}_{ANNOTATION_SUFFIXES[kind]}"
    name, number = base, 2
    while name.casefold() in used:
        name = f"{base}_{number}"
        number += 1
    return name


def build_duplicate_formula(source_column_index: int, excel_row: int, last_excel_row: int) -> str:
    """使用旧版 Excel 的 SUMPRODUCT/EXACT，区分大小写及数字/文本，不依赖动态数组。"""
    if (not 0 <= source_column_index < 16384
            or not 2 <= excel_row <= last_excel_row <= 1048576):
        raise ValueError("辅助列公式的 Excel 行列位置无效。")
    letter = get_column_letter(source_column_index + 1)
    reference = f"{letter}{excel_row}"
    area = "$" + letter + "$2:$" + letter + "$" + str(last_excel_row)
    # 只比较数据区，排除表头；EXACT 将 * / ? / ~ 当普通文字，不作为通配符。
    # 数值分支排除文本和真空单元格，布尔值与数值按现有检测器的相等关系比较。
    # 比较为真时倒数是 1，为假或遇到源错误值时是错误；ISNUMBER 将后两者排除。
    # 不在数组内使用 IF/IFERROR，避免旧版 Excel 的隐式交集，不需要 Ctrl+Shift+Enter。
    text_count = f"SUMPRODUCT(--ISNUMBER(1/EXACT({area},{reference})),--ISTEXT({area}))"
    # 数值和布尔值统一转成数值比较；类型掩码排除文本和真空单元格。
    # 例如 True 与 1 相等，但文本 "1" 不参与数值比较。
    number_count = (
        f"SUMPRODUCT(--ISNUMBER(1/({area}*1={reference}*1)),"
        f"ISNUMBER({area})+ISLOGICAL({area}))"
    )
    count = f"IF(ISTEXT({reference}),{text_count},{number_count})"
    # IFERROR 仅包裹最终结果，统一处理当前单元格错误，不在数组内进行条件分支。
    return f'=IFERROR(IF({reference}="","空值",IF({count}>1,"重复","正常")),"空值")'


def _validate_excel_types(column: pd.Series) -> None:
    # Excel 日期也存成数值；混合列中两种 Python 值对应同一 Excel 存储值时，
    # 通用公式无法区分，明确拒绝而不输出与检测不同的标注。
    dates = [value for value in column
             if isinstance(value, (date, datetime, time, timedelta)) and not is_empty_value(value)]
    if not dates:
        return
    numbers = {value for value in column if isinstance(value, Number) and not is_empty_value(value)}
    if numbers and any(
        to_excel(value) in numbers or to_excel(value, CALENDAR_MAC_1904) in numbers
        for value in dates
    ):
        raise ValueError("当前字段包含在 Excel 中存储值相同的日期和数值，无法保证精确标注，请先人工核对该字段。")


def build_duplicate_values(column: pd.Series) -> pd.Series:
    """沿用 V0.2 原始值分组口径，重复数量包含首次出现，空值单独标记。"""
    _validate_excel_types(column)
    duplicates = {record.excel_row - 2 for record in detect_duplicates(column)}
    return pd.Series(
        ["空值" if is_empty_value(value) else "重复" if position in duplicates else "正常"
         for position, value in enumerate(column)],
        index=column.index, dtype=object,
    )


def generate_duplicate_preview(
    dataframe: pd.DataFrame, source_column_index: int,
) -> AnnotationPreview:
    if not 0 <= source_column_index < dataframe.shape[1]:
        raise ValueError("请先选择有效字段。")
    if dataframe.shape[1] >= 16384:
        raise ValueError("工作表已达到 Excel 最大列数，不能新增辅助列。")
    source = dataframe.iloc[:, source_column_index].copy(deep=True)
    name = generate_annotation_name(str(source.name), dataframe.columns)
    values = build_duplicate_values(source)
    values.name = name
    return AnnotationPreview(AnnotationColumn(name, source_column_index), str(source.name), source, values)


def build_annotation_values(column: pd.Series, kind: str = "duplicate", *, date_confirmed: bool = False) -> pd.Series:
    if kind == "duplicate":
        return build_duplicate_values(column)
    if kind == "numeric_type":
        values = [numeric_type_label(value) for value in column]
    elif kind == "leading_zero":
        values = ["空值" if is_empty_value(value) else "前导0" if has_leading_zero(value) else "正常"
                  for value in column]
    elif kind == "date":
        results = [None if is_empty_value(value) else parse_date_value(value) for value in column]
        profile = summarize_dates(results, str(column.name))
        # 用户曾确认的日期辅助列继续按该含义刷新；未经确认时不误报普通文本。
        if not date_confirmed and not profile.suspected:
            values = ["空值" if result is None else "非日期列" for result in results]
        else:
            values = ["空值" if result is None else date_check_label(result) for result in results]
    else:
        raise ValueError("不支持的辅助列类型。")
    return pd.Series(values, index=column.index, dtype=object)


def generate_annotation_preview(
    dataframe: pd.DataFrame, source_column_index: int, kind: str = "duplicate", *,
    date_confirmed: bool = False,
) -> AnnotationPreview:
    if kind not in ANNOTATION_SUFFIXES:
        raise ValueError("不支持的辅助列类型。")
    if kind == "duplicate":
        return generate_duplicate_preview(dataframe, source_column_index)
    if not 0 <= source_column_index < dataframe.shape[1]:
        raise ValueError("请先选择有效字段。")
    if dataframe.shape[1] >= 16384:
        raise ValueError("工作表已达到 Excel 最大列数，不能新增辅助列。")
    source = dataframe.iloc[:, source_column_index].copy(deep=True)
    if kind == "date":
        dates = [None if is_empty_value(value) else parse_date_value(value) for value in source]
        if not date_confirmed and not summarize_dates(dates, str(source.name)).suspected:
            raise ValueError("该列日期证据不足，请先确认当前字段是日期列。")
    name = generate_annotation_name(str(source.name), dataframe.columns, kind)
    values = build_annotation_values(source, kind, date_confirmed=date_confirmed)
    values.name = name
    return AnnotationPreview(AnnotationColumn(name, source_column_index, kind, date_confirmed),
                             str(source.name), source, values)
