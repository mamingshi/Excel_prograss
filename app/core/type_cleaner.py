"""V0.8 安全类型转换：只返回结果，前导0、精度风险和模糊日期保持原样。"""

from decimal import Decimal
import math

from app.core.type_utils import has_leading_zero, is_numeric_text, parse_date_value, is_empty_value


_FULLWIDTH_MAP = {code: code - 0xFEE0 for code in range(0xFF01, 0xFF5F)}
_FULLWIDTH_MAP[0x3000] = ord(" ")


def fullwidth_to_halfwidth(value):
    """仅转换全角 ASCII 和全角空格，不使用会额外改变其他 Unicode 字符的 NFKC。"""
    return value.translate(_FULLWIDTH_MAP) if isinstance(value, str) else value


def numeric_conversion(value) -> tuple[object, str | None]:
    if not is_numeric_text(value):
        return value, None
    if has_leading_zero(value):
        return value, "前导0风险：转换为数字会丢失编码前缀，保留原值"
    decimal = Decimal(value)
    # Excel 仅可靠保存 15 位有效数字，拒绝长编码及不能精确往返的浮点文本。
    if len(decimal.as_tuple().digits) > 15:
        return value, "精度风险：超过15位有效数字，保留原值"
    if "." not in value:
        return int(value), None
    converted = float(decimal)
    if not math.isfinite(converted) or Decimal(str(converted)) != decimal:
        return value, "精度风险：转换后数值不一致，保留原值"
    return converted, None


def convert_text_number(value):
    return numeric_conversion(value)[0]


def date_conversion(value) -> tuple[object, str | None]:
    if is_empty_value(value):
        return value, None
    result = parse_date_value(value)
    if result.ambiguous:
        return value, "日期歧义：年月日顺序不明确，保留原值"
    if result.parsed is None:
        return value, "无法识别日期，保留原值"
    if result.has_time:
        return value, "包含时间或时区，保留原值以避免截断信息"
    parsed = result.parsed
    return f"{parsed.year:04d}-{parsed.month:02d}-{parsed.day:02d}", None


def standardize_date(value):
    return date_conversion(value)[0]
