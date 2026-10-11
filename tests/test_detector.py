"""五类检测规则、统计口径、Excel 行号和只读保证。"""

import pandas as pd
import pytest

from app.core.detector import (
    PROBLEM_TYPES,
    analyze_column,
    detect_duplicates,
    detect_empty_values,
    detect_hidden_chars,
    detect_nbsp,
    detect_outer_spaces,
)


def series(values, **kwargs):
    return pd.Series(values, dtype=object, name="商品编码", **kwargs)


def test_outer_spaces_only_checks_normal_spaces_at_edges():
    issues = detect_outer_spaces(series([" P001", "P001 ", " P001 ", "ABC 123", "P001", 0]))
    assert [issue.excel_row for issue in issues] == [2, 3, 4]
    assert "开头和末尾" in issues[2].description
    assert issues[0].original_value == " P001"


def test_nbsp_counts_cells_and_reports_locations():
    issues = detect_nbsp(series(["P001\u00a0", "\u00a0P002", "P\u00a0\u00a0003", "\u00a0", " ABC"]))
    assert [issue.excel_row for issue in issues] == [2, 3, 4, 5]
    assert "末尾包含不间断空格" in issues[0].description
    assert "开头" in issues[1].description
    assert "中间" in issues[2].description
    assert "共 2 个" in issues[2].description


@pytest.mark.parametrize("character", ["\t", "\n", "\r", "\r\n", "\t\n\r"])
def test_hidden_chars_are_counted_once_per_cell(character):
    issues = detect_hidden_chars(series([f"P{character}001", "ABC 123"]))
    assert len(issues) == 1
    assert issues[0].excel_row == 2
    assert issues[0].original_value == f"P{character}001"


def test_hidden_description_identifies_each_character():
    issue = detect_hidden_chars(series(["A\t\tB\nC\r"]))[0]
    assert "Tab（\\t） × 2" in issue.description
    assert "换行（\\n） × 1" in issue.description
    assert "回车（\\r） × 1" in issue.description


def test_empty_values_include_nan_none_na_and_empty_string_but_not_zero():
    issues = detect_empty_values(series([float("nan"), None, pd.NA, "", 0, "0", " ", "NaN"]))
    assert [issue.excel_row for issue in issues] == [2, 3, 4, 5]
    assert issues[-1].description == "纯空字符串"


def test_nullable_pandas_values_are_supported():
    column = pd.Series([0, pd.NA, 12], dtype="Int64", name="数量")
    result = analyze_column(column)
    assert result.counts["空值"] == 1
    assert result.issues[0].excel_row == 3


def test_duplicates_exclude_empty_values_and_include_first_occurrence():
    column = series(["A", None, "", "A", float("nan"), pd.NA, "B", "B", "B", 0, "0", 1, "1"])
    issues = detect_duplicates(column)
    assert [issue.excel_row for issue in issues] == [2, 5, 8, 9, 10]
    result = analyze_column(column)
    assert result.counts["重复值"] == 5
    assert [(group.value, group.excel_rows) for group in result.duplicate_groups] == [
        ("A", (2, 5)),
        ("B", (8, 9, 10)),
    ]
    assert "含首次出现" in issues[0].description
    assert "2、5" in issues[0].description


def test_duplicates_do_not_trim_or_coerce_text_and_numbers():
    column = series(["P001", " P001", "P001 ", 1002, "1002", 0, "0"])
    assert detect_duplicates(column) == []


def test_numeric_equal_values_are_duplicates_without_string_conversion():
    issues = detect_duplicates(series([1, "1", 1.0, "1 "]))
    assert [issue.excel_row for issue in issues] == [2, 4]
    assert type(issues[0].original_value) is int
    assert type(issues[1].original_value) is float


def test_large_duplicate_group_retains_all_rows_without_huge_descriptions():
    result = analyze_column(series(["P001"] * 1000))
    assert result.counts["重复值"] == 1000
    assert result.duplicate_groups[0].excel_rows == tuple(range(2, 1002))
    assert len(result.issues[0].description) < 300


def test_one_cell_can_have_multiple_issue_types():
    value = " \u00a0P001\t\n\r "
    result = analyze_column(series([value, value]))
    assert result.counts == {"前后空格": 2, "CHAR(160)": 2, "Tab / 换行": 2, "空值": 0, "重复值": 2,
                             "文本数字混合": 0, "前导0风险": 0, "日期异常": 0, "全角字符": 0}
    assert len(result.issues) == 8
    assert {issue.problem_type for issue in result.issues if issue.excel_row == 2} == {
        "前后空格", "CHAR(160)", "Tab / 换行", "重复值"
    }
    assert all(issue.original_value == value for issue in result.issues)


def test_normal_values_only_report_the_new_numeric_type_mixture():
    result = analyze_column(series(["P001", "ABC 123", 1002, 0, "0", "NA", "NaN", "None"]))
    expected = dict.fromkeys(PROBLEM_TYPES, 0)
    expected["文本数字混合"] = 1
    assert result.counts == expected
    assert [(issue.excel_row, issue.original_value, issue.problem_type) for issue in result.issues] == [(6, "0", "文本数字混合")]
    assert result.duplicate_groups == ()


def test_fullwidth_space_is_supported_and_other_unicode_remains_out_of_scope():
    result = analyze_column(series(["\u3000P001", "P\u200b001", "P001\u2007", "A\vB"]))
    assert [(issue.excel_row, issue.problem_type) for issue in result.issues] == [(2, "全角字符")]


def test_excel_rows_follow_positions_not_series_index():
    result = analyze_column(series([" P001", None, "P001\u00a0"], index=[100, "行", 100]))
    assert [issue.excel_row for issue in result.issues] == [2, 3, 4]
    shifted = analyze_column(series(["A", "A"], index=[999, -1]), first_excel_row=20)
    assert shifted.duplicate_groups[0].excel_rows == (20, 21)


def test_analysis_does_not_change_dataframe_or_series():
    dataframe = pd.DataFrame({"商品编码": series([" P001", "P002\u00a0", "A\tB", None, "", "DUP", "DUP", 0])})
    before = dataframe.copy(deep=True)
    column = dataframe.iloc[:, 0]
    before_column = column.copy(deep=True)
    result = analyze_column(column)
    assert result.row_count == len(column)
    assert result.column_name == "商品编码"
    pd.testing.assert_series_equal(column, before_column)
    pd.testing.assert_frame_equal(dataframe, before)


def test_empty_column_has_zero_counts():
    result = analyze_column(series([]))
    assert result.row_count == 0
    assert result.counts == dict.fromkeys(PROBLEM_TYPES, 0)
    assert result.issues == ()


def test_invalid_excel_start_row_is_rejected():
    with pytest.raises(ValueError, match="起始行号"):
        analyze_column(series(["A"]), first_excel_row=0)
