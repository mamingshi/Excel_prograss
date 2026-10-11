"""V0.3 清洗规则、固定顺序、类型保持和输入数据不变。"""

from datetime import datetime

import pandas as pd
import pytest

from app.core.cleaner import (
    CleaningRules,
    RULE_ORDER,
    apply_cleaning_rules,
    clean_carriage_return,
    clean_hidden_whitespace,
    clean_nbsp,
    clean_newline,
    clean_outer_spaces,
    clean_tab,
    generate_cleaning_preview,
)


ALL_RULES = CleaningRules(nbsp=True, tab=True, newline=True, carriage_return=True, outer_spaces=True)


@pytest.mark.parametrize("original,expected", [
    (" P001 ", "P001"),
    ("P001\u00a0", "P001"),
    ("ABC\t123", "ABC 123"),
    ("ABC\n123", "ABC 123"),
    ("ABC\r123", "ABC 123"),
    ("ABC 123", "ABC 123"),
    (" ABC  123 ", "ABC  123"),
    (" \tP001\u00a0\n ", "P001"),
    ("ABC\u00a0123", "ABC 123"),
    ("A\r\nB", "A  B"),
    ("", ""),
])
def test_requested_cleaning_examples(original, expected):
    assert apply_cleaning_rules(original, ALL_RULES) == expected


def test_outer_spaces_does_not_touch_middle_spaces_or_other_whitespace():
    assert clean_outer_spaces(" ABC 123 ") == "ABC 123"
    assert clean_outer_spaces("\t ABC  123 \u00a0") == "\t ABC  123 \u00a0"
    assert clean_outer_spaces("   ") == ""


def test_nbsp_is_replaced_not_deleted_or_trimmed():
    assert clean_nbsp("P001\u00a0") == "P001 "
    assert clean_nbsp("ABC\u00a0\u00a0123") == "ABC  123"


@pytest.mark.parametrize("function,character", [
    (clean_tab, "\t"), (clean_newline, "\n"), (clean_carriage_return, "\r"),
])
def test_hidden_characters_have_two_strategies(function, character):
    assert function(f"ABC{character}123") == "ABC 123"
    assert function(f"ABC{character}123", strategy="delete") == "ABC123"


def test_hidden_whitespace_combines_only_tab_newline_and_carriage_return():
    assert clean_hidden_whitespace("A\tB\nC\rD\u00a0") == "A B C D\u00a0"
    assert clean_hidden_whitespace("A\tB\nC\rD", "delete") == "ABCD"


@pytest.mark.parametrize("value", [1002, 0, 12.5, None, float("nan"), pd.NA, True, datetime(2026, 10, 5)])
def test_non_strings_keep_their_value_and_type(value):
    functions = (clean_outer_spaces, clean_nbsp, clean_tab, clean_newline,
                 clean_carriage_return, clean_hidden_whitespace)
    for function in functions:
        assert function(value) is value
    assert apply_cleaning_rules(value, ALL_RULES) is value


def test_unselected_rules_are_not_applied():
    value = " \tP001\u00a0\n "
    assert apply_cleaning_rules(value, CleaningRules()) == value
    assert apply_cleaning_rules(value, CleaningRules(outer_spaces=True)) == "\tP001\u00a0\n"
    assert apply_cleaning_rules("P001\u00a0", CleaningRules(nbsp=True)) == "P001 "


def test_fixed_order_and_operation_log():
    column = pd.Series([" \tP001\u00a0\n\r "], dtype=object, name="商品编码")
    preview = generate_cleaning_preview(column, ALL_RULES)
    assert RULE_ORDER == ("nbsp", "tab", "newline", "carriage_return", "outer_spaces")
    assert preview.records[0].cleaned_value == "P001"
    assert preview.records[0].operations == (
        "CHAR(160) → 普通空格", "Tab → 普通空格", "换行（\\n）→ 普通空格",
        "回车（\\r）→ 普通空格", "去除前后普通空格",
    )
    assert preview.processed_counts == {"前后空格": 1, "CHAR(160)": 1, "Tab / 换行": 1}
    assert preview.changed_count == 1  # 多个处理类别不重复计算变化行。


def test_delete_strategy_is_explicit_and_nbsp_still_becomes_space():
    rules = CleaningRules(nbsp=True, tab=True, newline=True, carriage_return=True, hidden_strategy="delete")
    preview = generate_cleaning_preview(pd.Series(["ABC\t123\u00a0"], dtype=object), rules)
    assert preview.records[0].cleaned_value == "ABC123 "
    assert preview.records[0].operations == ("CHAR(160) → 普通空格", "删除 Tab")


def test_only_changed_operations_are_recorded():
    preview = generate_cleaning_preview(pd.Series(["ABC 123", " P001 "], dtype=object), ALL_RULES)
    assert not preview.records[0].changed
    assert preview.records[0].operations == ()
    assert preview.records[1].operations == ("去除前后普通空格",)
    assert (preview.row_count, preview.changed_count, preview.unchanged_count) == (2, 1, 1)


def test_dataframe_is_unchanged_and_preview_storage_is_independent():
    dataframe = pd.DataFrame({"商品编码": pd.Series(
        [" P001 ", "ABC\t123", 1002, 0, None, float("nan"), "", "DUP", "DUP"], dtype=object
    )})
    original = dataframe.copy(deep=True)
    preview = generate_cleaning_preview(dataframe.iloc[:, 0], ALL_RULES)
    assert preview.changed_count == 2
    assert preview.unchanged_count == 7
    assert len(preview.records) == len(dataframe)
    assert preview.preview_series.iloc[2] == 1002
    assert type(preview.preview_series.iloc[2]) is int
    assert preview.preview_series.iloc[3] == 0
    assert preview.preview_series.iloc[4] is None
    assert pd.isna(preview.preview_series.iloc[5])
    assert preview.preview_series.iloc[6] == ""
    assert preview.preview_series.iloc[7:].tolist() == ["DUP", "DUP"]
    assert all(not record.changed for record in preview.records[2:])
    pd.testing.assert_frame_equal(dataframe, original)
    preview.preview_series.iloc[0] = "只修改独立预览副本"
    pd.testing.assert_frame_equal(dataframe, original)


@pytest.mark.parametrize("dtype,values", [
    ("Int64", [1002, 0, pd.NA]),
    ("string", [" P001 ", pd.NA, "ABC 123"]),
    ("float64", [12.5, float("nan"), 0.0]),
])
def test_series_dtype_and_missing_values_are_preserved(dtype, values):
    column = pd.Series(values, dtype=dtype, name="字段")
    original = column.copy(deep=True)
    preview = generate_cleaning_preview(column, ALL_RULES)
    assert preview.preview_series.dtype == column.dtype
    assert preview.preview_series.isna().equals(column.isna())
    pd.testing.assert_series_equal(column, original)


def test_excel_rows_follow_position_while_series_index_is_preserved():
    column = pd.Series([" P001 ", "P002\u00a0"], index=["行", "行"], name="编码", dtype=object)
    preview = generate_cleaning_preview(column, ALL_RULES, first_excel_row=20)
    assert [record.excel_row for record in preview.records] == [20, 21]
    assert preview.preview_series.index.equals(column.index)
    assert preview.preview_series.name == "编码"


def test_second_application_does_not_introduce_further_changes():
    first = generate_cleaning_preview(pd.Series([" \tP001\u00a0\n ", "ABC  123"], dtype=object), ALL_RULES)
    second = generate_cleaning_preview(first.preview_series, ALL_RULES)
    assert second.changed_count == 0
    pd.testing.assert_series_equal(first.preview_series, second.preview_series)


def test_future_unicode_cleaning_is_not_implemented():
    value = "\u3000ABC\u200b123\u202f"
    assert apply_cleaning_rules(value, ALL_RULES) == value


def test_empty_series_has_valid_zero_statistics():
    result = generate_cleaning_preview(pd.Series([], dtype=object), ALL_RULES)
    assert (result.row_count, result.changed_count, result.unchanged_count) == (0, 0, 0)
    assert result.records == ()


def test_invalid_configuration_is_rejected():
    with pytest.raises(ValueError, match="处理策略"):
        CleaningRules(hidden_strategy="unknown")
    with pytest.raises(ValueError, match="起始行号"):
        generate_cleaning_preview(pd.Series(["P001"]), ALL_RULES, first_excel_row=0)
