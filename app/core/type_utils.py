"""类型特征的纯函数：检测、辅助标注和转换共用解析口径，不修改数据。"""

from collections import Counter
from dataclasses import dataclass
from datetime import date, datetime
from decimal import Decimal
from math import isfinite
from numbers import Real
import re

import pandas as pd


NUMERIC_TEXT = re.compile(r"[+-]?(?:[0-9]+(?:\.[0-9]+)?|\.[0-9]+)\Z")
YEAR_FIRST = re.compile(r"([0-9]{4})([-/])([0-9]{1,2})\2([0-9]{1,2})\Z")
CHINESE_DATE = re.compile(r"([0-9]{4})年([0-9]{1,2})月([0-9]{1,2})日\Z")
YEAR_LAST = re.compile(r"([0-9]{1,2})/([0-9]{1,2})/([0-9]{2}|[0-9]{4})\Z")
ISO_TIME = re.compile(r"[0-9]{4}-[0-9]{2}-[0-9]{2}[ T][0-9]{2}:[0-9]{2}(?::[0-9]{2}(?:\.[0-9]{1,6})?)?\Z")
DATE_THRESHOLD = 0.60


def is_empty_value(value) -> bool:
    if isinstance(value, str):
        return value == ""
    return value is None or bool(pd.isna(value))


def is_numeric_value(value) -> bool:
    return (isinstance(value, (Real, Decimal)) and not isinstance(value, bool)
            and not is_empty_value(value) and isfinite(value))


def is_numeric_text(value) -> bool:
    # 不猜测金额、千分位、百分比、科学计数法或含空白的字符串。
    return isinstance(value, str) and NUMERIC_TEXT.fullmatch(value) is not None


def has_leading_zero(value) -> bool:
    if not is_numeric_text(value):
        return False
    integer = value.lstrip("+-").split(".", 1)[0]
    return len(integer) > 1 and integer.startswith("0")


def contains_fullwidth(value) -> bool:
    return isinstance(value, str) and any(
        "\uff01" <= character <= "\uff5e" or character == "\u3000" for character in value
    )


def numeric_type_label(value) -> str:
    if is_empty_value(value):
        return "空值"
    if is_numeric_value(value):
        return "数字"
    if is_numeric_text(value):
        return "文本数字（前导0）" if has_leading_zero(value) else "文本数字"
    if isinstance(value, (date, datetime)):
        return "日期"
    if isinstance(value, bool):
        return "布尔值"
    return "文本" if isinstance(value, str) else "其他"


@dataclass(frozen=True)
class DateValue:
    parsed: datetime | None = None
    style: str = "无法识别"
    ambiguous: bool = False

    @property
    def has_time(self) -> bool:
        return self.parsed is not None and (
            self.parsed.hour != 0 or self.parsed.minute != 0 or self.parsed.second != 0
            or self.parsed.microsecond != 0 or self.parsed.tzinfo is not None
        )


UNKNOWN_DATE = DateValue()
AMBIGUOUS_DATE = DateValue(style="日期歧义", ambiguous=True)


def _make_date(year: int, month: int, day: int, style: str) -> DateValue:
    try:
        return DateValue(datetime(year, month, day), style)
    except ValueError:
        return UNKNOWN_DATE


def parse_date_value(value) -> DateValue:
    """只按明确格式解析；不把裸数值/8位编码猜成 Excel 日期或 Unix 时间。"""
    if is_empty_value(value):
        return UNKNOWN_DATE
    if isinstance(value, datetime):
        return DateValue(value, "原生日期")
    if isinstance(value, date):
        return DateValue(datetime(value.year, value.month, value.day), "原生日期")
    if not isinstance(value, str):
        return UNKNOWN_DATE
    if match := YEAR_FIRST.fullmatch(value):
        year, _, month, day = match.groups()
        style = "标准日期" if match[2] == "-" and len(month) == len(day) == 2 else "非标准日期"
        return _make_date(int(year), int(month), int(day), style)
    if match := CHINESE_DATE.fullmatch(value):
        return _make_date(*(int(part) for part in match.groups()), "中文日期")
    if match := YEAR_LAST.fullmatch(value):
        first, second, year = (int(part) for part in match.groups())
        if len(match[3]) == 2:
            return AMBIGUOUS_DATE
        if first < 1 or second < 1 or first > 31 or second > 31:
            return UNKNOWN_DATE
        if first <= 12 and second <= 12 and first != second:
            return AMBIGUOUS_DATE
        if first > 12 and second <= 12:
            return _make_date(year, second, first, "日月年日期")
        if first <= 12:
            return _make_date(year, first, second, "月日年日期")
        return UNKNOWN_DATE
    if ISO_TIME.fullmatch(value):
        try:
            return DateValue(datetime.fromisoformat(value), "文本日期时间")
        except ValueError:
            return UNKNOWN_DATE
    return UNKNOWN_DATE


@dataclass(frozen=True)
class DateProfile:
    nonempty_count: int
    recognized_count: int
    suspected: bool
    formats: tuple[str, ...]
    name_hint: bool = False

    @property
    def ratio(self) -> float:
        return self.recognized_count / self.nonempty_count if self.nonempty_count else 0.0


def summarize_dates(results: list[DateValue | None], column_name: str) -> DateProfile:
    nonempty = sum(result is not None for result in results)
    formats = Counter(result.style for result in results if result is not None and result.parsed is not None)
    recognized = sum(formats.values())
    hint = any(word in column_name.casefold() for word in ("日期", "时间", "date", "time"))
    return DateProfile(nonempty, recognized, bool(nonempty and recognized / nonempty >= DATE_THRESHOLD),
                       tuple(formats), hint)


def date_problem(result: DateValue) -> str | None:
    if result.ambiguous:
        return "日期顺序存在歧义，保留原值并人工确认"
    if result.parsed is None:
        return "无法识别日期，保留原值"
    if result.has_time:
        return "包含时间或时区，统一日期格式会丢失信息，保留原值"
    if result.style not in ("标准日期", "原生日期"):
        return "格式不统一；目标格式为 YYYY-MM-DD"
    return None


def date_check_label(result: DateValue) -> str:
    description = date_problem(result)
    if description is None:
        return "正常"
    if result.ambiguous:
        return "日期歧义"
    if result.parsed is None:
        return "无法识别"
    return "含时间（保留）" if result.has_time else "格式不统一"
