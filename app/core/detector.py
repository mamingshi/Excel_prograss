"""列质量检测：单列明细与整表统计共享同一套规则，始终只读。"""

from dataclasses import dataclass
from typing import Any, Callable

import pandas as pd

from app.core.type_utils import (
    DateProfile, DateValue, is_empty_value, is_numeric_value, is_numeric_text,
    has_leading_zero, contains_fullwidth, parse_date_value, summarize_dates, date_problem,
)


PROBLEM_TYPES = ("前后空格", "CHAR(160)", "Tab / 换行", "空值", "重复值",
                 "文本数字混合", "前导0风险", "日期异常", "全角字符")


@dataclass(frozen=True)
class IssueRecord:
    excel_row: int
    original_value: Any
    problem_type: str
    description: str


@dataclass(frozen=True)
class DuplicateGroup:
    value: Any
    excel_rows: tuple[int, ...]


@dataclass(frozen=True)
class ColumnAnalysis:
    column_name: str
    row_count: int
    counts: dict[str, int]
    issues: tuple[IssueRecord, ...]
    duplicate_groups: tuple[DuplicateGroup, ...]
    affected_rows: int = 0
    numeric_text_count: int = 0
    date_profile: DateProfile | None = None


def _cells(column: pd.Series, first_excel_row: int):
    # 按数据位置计算行号，不依赖可能非连续、重复或为文本的 Series index。
    return enumerate(column, start=first_excel_row)


def _outer_spaces_description(value: str) -> str | None:
    locations = []
    if value.startswith(" "):
        locations.append("开头")
    if value.endswith(" "):
        locations.append("末尾")
    return "和".join(locations) + "包含普通空格（U+0020）" if locations else None


def _nbsp_description(value: str) -> str | None:
    if "\u00a0" not in value:
        return None
    locations = []
    if value.startswith("\u00a0"):
        locations.append("开头")
    if "\u00a0" in value[1:-1]:
        locations.append("中间")
    if value.endswith("\u00a0"):
        locations.append("末尾")
    return "、".join(locations) + f"包含不间断空格（U+00A0），共 {value.count(chr(160))} 个"


def _hidden_chars_description(value: str) -> str | None:
    characters = (("\t", "Tab（\\t）"), ("\n", "换行（\\n）"), ("\r", "回车（\\r）"))
    found = [f"{name} × {value.count(character)}" for character, name in characters if character in value]
    return "包含 " + "、".join(found) if found else None


def _leading_zero_description(value: str) -> str | None:
    return ("文本数字存在前导0，转换为数字可能改变编码含义；禁止自动转换"
            if has_leading_zero(value) else None)


def _fullwidth_description(value: str) -> str | None:
    return ("包含全角 ASCII 字符或全角空格（U+FF01～U+FF5E / U+3000）"
            if contains_fullwidth(value) else None)


TEXT_RULES = (
    ("前后空格", _outer_spaces_description),
    ("CHAR(160)", _nbsp_description),
    ("Tab / 换行", _hidden_chars_description),
    ("前导0风险", _leading_zero_description),
    ("全角字符", _fullwidth_description),
)


def _detect_text_column(
    column: pd.Series, first_excel_row: int,
    problem_type: str, describe: Callable[[str], str | None],
) -> list[IssueRecord]:
    issues = []
    for excel_row, value in _cells(column, first_excel_row):
        if isinstance(value, str) and (description := describe(value)) is not None:
            issues.append(IssueRecord(excel_row, value, problem_type, description))
    return issues


def detect_outer_spaces(column: pd.Series, first_excel_row: int = 2) -> list[IssueRecord]:
    return _detect_text_column(column, first_excel_row, "前后空格", _outer_spaces_description)


def detect_nbsp(column: pd.Series, first_excel_row: int = 2) -> list[IssueRecord]:
    return _detect_text_column(column, first_excel_row, "CHAR(160)", _nbsp_description)


def detect_hidden_chars(column: pd.Series, first_excel_row: int = 2) -> list[IssueRecord]:
    return _detect_text_column(column, first_excel_row, "Tab / 换行", _hidden_chars_description)


def _empty_record(excel_row: int, value: Any) -> IssueRecord:
    description = "纯空字符串" if isinstance(value, str) else "空单元格或缺失值（None / NaN / NA）"
    return IssueRecord(excel_row, value, "空值", description)


def detect_empty_values(column: pd.Series, first_excel_row: int = 2) -> list[IssueRecord]:
    return [_empty_record(row, value) for row, value in _cells(column, first_excel_row) if is_empty_value(value)]


def _add_group(groups: dict, value: Any, excel_row: int) -> None:
    # 唯一值只保存一个行号；确实重复后才分配行号列表。
    previous = groups.get(value)
    if previous is None:
        groups[value] = excel_row
    elif isinstance(previous, int):
        groups[value] = [previous, excel_row]
    else:
        previous.append(excel_row)


def _duplicate_records(rows: list[int], values: list, first_excel_row: int) -> list[IssueRecord]:
    row_text = "、".join(str(row) for row in rows[:20])
    if len(rows) > 20:
        row_text += f"……（共 {len(rows)} 行）"
    description = f"相同原始值共 {len(rows)} 条（含首次出现）；Excel 行号：{row_text}"
    return [IssueRecord(row, values[row - first_excel_row], "重复值", description) for row in rows]


def detect_duplicates(column: pd.Series, first_excel_row: int = 2) -> list[IssueRecord]:
    # 原始值直接分组，不去空格、不互转数字和文本；排除空值，包含首次出现。
    values = list(column)
    groups: dict[Any, int | list[int]] = {}
    for excel_row, value in enumerate(values, start=first_excel_row):
        if not is_empty_value(value):
            _add_group(groups, value, excel_row)
    issues = [
        issue for rows in groups.values() if isinstance(rows, list)
        for issue in _duplicate_records(rows, values, first_excel_row)
    ]
    return sorted(issues, key=lambda issue: issue.excel_row)


def analyze_column(
    column: pd.Series, first_excel_row: int = 2, *, include_details: bool = True,
) -> ColumnAnalysis:
    """单次收集基础问题及类型特征；列上下文确定后汇总混合类型和日期问题。"""
    if first_excel_row < 1:
        raise ValueError("Excel 起始行号必须大于或等于 1。")
    counts = dict.fromkeys(PROBLEM_TYPES, 0)
    issues = []
    groups: dict[Any, int | list[int]] = {}
    values = [] if include_details else None
    # 每个位置只有一个标记，多种问题不会重复计入受影响行数。
    affected = bytearray(len(column))
    numeric_rows = []
    numeric_count = 0
    dates: list[DateValue | None] = []
    for position, value in enumerate(column):
        excel_row = position + first_excel_row
        if values is not None:
            values.append(value)
        if is_empty_value(value):
            dates.append(None)
            counts["空值"] += 1
            affected[position] = 1
            if include_details:
                issues.append(_empty_record(excel_row, value))
            continue
        dates.append(parse_date_value(value))
        numeric_count += int(is_numeric_value(value))
        if is_numeric_text(value):
            numeric_rows.append(position)
        if isinstance(value, str):
            for problem_type, describe in TEXT_RULES:
                if (description := describe(value)) is not None:
                    counts[problem_type] += 1
                    affected[position] = 1
                    if include_details:
                        issues.append(IssueRecord(excel_row, value, problem_type, description))
        _add_group(groups, value, excel_row)

    if numeric_count and numeric_rows:
        counts["文本数字混合"] = len(numeric_rows)
        for position in numeric_rows:
            affected[position] = 1
            if include_details:
                value = column.iloc[position]
                description = "本列同时存在数值类型和文本数字"
                if has_leading_zero(value):
                    description += "；文本数字存在前导0风险，不能自动转数字"
                issues.append(IssueRecord(position + first_excel_row, value, "文本数字混合", description))
    date_profile = summarize_dates(dates, str(column.name))
    if date_profile.suspected:
        for position, result in enumerate(dates):
            if result is not None and (description := date_problem(result)) is not None:
                counts["日期异常"] += 1
                affected[position] = 1
                if include_details:
                    issues.append(IssueRecord(position + first_excel_row, column.iloc[position], "日期异常", description))

    duplicate_groups = []
    for value, rows in groups.items():
        if not isinstance(rows, list):
            continue
        counts["重复值"] += len(rows)
        for row in rows:
            affected[row - first_excel_row] = 1
        if include_details:
            issues.extend(_duplicate_records(rows, values, first_excel_row))
            duplicate_groups.append(DuplicateGroup(value, tuple(rows)))
    return ColumnAnalysis(
        column_name=str(column.name) if column.name is not None else "未命名列",
        row_count=len(column),
        counts=counts,
        issues=tuple(sorted(issues, key=lambda issue: issue.excel_row)),
        duplicate_groups=tuple(duplicate_groups),
        affected_rows=sum(affected),
        numeric_text_count=len(numeric_rows), date_profile=date_profile,
    )


def detect_numeric_type_mixture(column: pd.Series, first_excel_row: int = 2) -> list[IssueRecord]:
    """只标记混合列中的文本数字；普通文本不当作数字。"""
    return [record for record in analyze_column(column, first_excel_row).issues
            if record.problem_type == "文本数字混合"]


def detect_leading_zero_risk(column: pd.Series, first_excel_row: int = 2) -> list[IssueRecord]:
    return _detect_text_column(column, first_excel_row, "前导0风险",
                               _leading_zero_description)


def detect_fullwidth_chars(column: pd.Series, first_excel_row: int = 2) -> list[IssueRecord]:
    return _detect_text_column(column, first_excel_row, "全角字符",
                               _fullwidth_description)


def detect_date_anomalies(
    column: pd.Series, first_excel_row: int = 2, *, confirmed: bool = False,
) -> list[IssueRecord]:
    results = [None if is_empty_value(value) else parse_date_value(value) for value in column]
    profile = summarize_dates(results, str(column.name))
    if not confirmed and not profile.suspected:
        return []
    return [IssueRecord(position + first_excel_row, column.iloc[position], "日期异常", description)
            for position, result in enumerate(results)
            if result is not None and (description := date_problem(result)) is not None]
