"""V0.8 安全转换、保护说明、类型变化和 Undo 前的独立预览。"""

from datetime import datetime

import pandas as pd
import pytest

from app.core.cleaner import CleaningRules, apply_cleaning_rules, generate_cleaning_preview
from app.core.type_cleaner import fullwidth_to_halfwidth, convert_text_number, standardize_date


@pytest.mark.parametrize("source,expected", [
    ("Ｐ００１", "P001"), ("ＡＢＣ１２３", "ABC123"), ("　ＡＢＣ！", " ABC!"),
    ("P001", "P001"), ("中文①ﬀ\u200b", "中文①ﬀ\u200b"), (1002, 1002), (None, None),
])
def test_fullwidth_conversion_is_limited_to_requested_range(source, expected):
    assert fullwidth_to_halfwidth(source) == expected


@pytest.mark.parametrize("source,expected", [
    ("100", 100), ("12.5", 12.5), ("-5", -5), ("0", 0), ("0.5", 0.5),
    ("00123", "00123"), ("00.5", "00.5"), ("1,000", "1,000"), ("￥100", "￥100"),
    ("10%", "10%"), ("1e3", "1e3"), ("ABC", "ABC"), (" 100", " 100"),
    ("1234567890123456", "1234567890123456"), ("0.1234567890123456", "0.1234567890123456"),
    (1002, 1002), (0, 0), (None, None),
])
def test_numeric_conversion_protects_codes_and_precision(source, expected):
    actual = convert_text_number(source)
    assert actual == expected and type(actual) is type(expected)


@pytest.mark.parametrize("source,expected", [
    ("2026/10/10", "2026-10-10"), ("2026年10月10日", "2026-10-10"),
    ("2026-10-10", "2026-10-10"), ("10/10/2026", "2026-10-10"),
    ("13/10/2026", "2026-10-13"), ("10/13/2026", "2026-10-13"),
    (datetime(2026, 10, 10), "2026-10-10"), ("abc", "abc"),
    ("01/02/03", "01/02/03"), ("01/02/2026", "01/02/2026"),
    ("2026-02-30", "2026-02-30"), ("2026-10-10 15:30:00", "2026-10-10 15:30:00"),
    (datetime(2026, 10, 10, 15, 30), datetime(2026, 10, 10, 15, 30)), (None, None),
])
def test_date_conversion_uses_unambiguous_formats_and_retains_unknowns(source, expected):
    assert standardize_date(source) == expected


def test_date_column_confirmation_is_required():
    with pytest.raises(ValueError, match="确认"):
        CleaningRules(date_normalize=True)


def test_combined_rules_keep_leading_zero_and_show_protection_reason():
    column = pd.Series(["　００１２３　", "１００２", 1002, None, float("nan")],
                       dtype=object, name="编码", index=[9] * 5)
    before = column.copy(deep=True)
    preview = generate_cleaning_preview(column, CleaningRules(
        fullwidth=True, outer_spaces=True, numeric_conversion=True,
    ))
    assert preview.preview_series.iloc[:3].tolist() == ["00123", 1002, 1002]
    assert preview.changed_count == 2
    assert "前导0" in preview.records[0].notes[0]
    assert preview.records[1].operations == ("全角转半角", "文本数字转数字")
    assert preview.records[0].excel_row == 2
    assert preview.records[2].operations == ()
    pd.testing.assert_series_equal(column, before)


def test_date_preview_retains_ambiguity_and_invalid_rows_with_explanation():
    column = pd.Series(["2026/10/10", "abc", "01/02/03", None], name="日期", dtype=object)
    result = generate_cleaning_preview(column, CleaningRules(date_normalize=True, date_confirmed=True))
    assert result.changed_count == 1
    assert result.records[1].notes == ("无法识别日期，保留原值",)
    assert "歧义" in result.records[2].notes[0]
    assert result.preview_series.tolist() == ["2026-10-10", "abc", "01/02/03", None]


@pytest.mark.parametrize("value", [None, float("nan"), pd.NA, pd.NaT, 1002, True])
def test_missing_and_nontext_values_are_safe(value):
    assert fullwidth_to_halfwidth(value) is value
    assert convert_text_number(value) is value
    result = apply_cleaning_rules(value, CleaningRules(fullwidth=True, numeric_conversion=True))
    assert result is value


def test_string_dtype_can_be_safely_converted_without_changing_input():
    column = pd.Series(["1002", "00123", pd.NA], dtype="string", name="编码")
    before = column.copy(deep=True)
    result = generate_cleaning_preview(column, CleaningRules(numeric_conversion=True))
    assert result.preview_series.dtype == object
    assert isinstance(result.preview_series.iloc[0], int)
    assert result.preview_series.iloc[1] == "00123"
    pd.testing.assert_series_equal(column, before)
