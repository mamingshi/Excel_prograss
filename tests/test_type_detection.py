"""V0.8 四类检测、日期证据门槛和唯一受影响行计数。"""

from datetime import datetime

from openpyxl import Workbook
import pandas as pd
import pytest

from app.core.detector import (
    analyze_column, detect_numeric_type_mixture, detect_leading_zero_risk,
    detect_date_anomalies, detect_fullwidth_chars,
)
from app.core.file_loader import ExcelFileLoader
from app.core.sheet_analyzer import analyze_sheet
from app.core.type_utils import is_numeric_text, parse_date_value


@pytest.mark.parametrize("value,expected", [
    ("1002", True), ("12.5", True), ("-5", True), ("00123", True), ("0", True),
    ("ABC", False), ("1,000", False), ("￥100", False), ("10%", False),
    ("1e3", False), (" 100", False), ("１２３", False), (1002, False), (None, False),
])
def test_pure_numeric_text_only(value, expected):
    assert is_numeric_text(value) is expected


def test_mixture_marks_text_numbers_and_keeps_business_values_readonly():
    column = pd.Series([1001, "1002", 1003, "1004", "ABC", "00123", 0, "0", None],
                       dtype=object, name="金额", index=[99] * 9)
    before = column.copy(deep=True)
    issues = detect_numeric_type_mixture(column)
    assert [issue.excel_row for issue in issues] == [3, 5, 7, 9]
    assert "前导0" in issues[2].description
    analysis = analyze_column(column)
    assert analysis.counts["文本数字混合"] == 4
    assert analysis.counts["前导0风险"] == 1
    assert analysis.affected_rows == 5  # 四个文本数字和一个空值；前导0不重复计数
    pd.testing.assert_series_equal(column, before)
    assert detect_numeric_type_mixture(pd.Series(["1001", "1002", "ABC"])) == []
    assert detect_numeric_type_mixture(pd.Series([True, "1"])) == []


@pytest.mark.parametrize("value,risk", [
    ("00123", True), ("0008", True), ("012345", True), ("000", True),
    ("-0012", True), ("00.5", True), ("0", False), ("0.5", False),
    ("P001", False), ("012ABC", False), (0, False), (None, False),
])
def test_leading_zero_risk(value, risk):
    assert bool(detect_leading_zero_risk(pd.Series([value], dtype=object))) is risk


def test_fullwidth_scope_does_not_include_unrelated_unicode():
    values = ["Ｐ００１", "ＡＢＣ１２３", "　", "！", "P001", "中文①", "a\u200bb", None, 1002]
    assert [issue.excel_row for issue in detect_fullwidth_chars(pd.Series(values, dtype=object))] == [2, 3, 4, 5]


def test_dates_have_explicit_formats_and_real_calendar_validation():
    column = pd.Series(["2026-10-10", "2026/10/10", "2026年10月10日", "10/10/2026", "abc"],
                       dtype=object, name="订单日期")
    issues = detect_date_anomalies(column)
    assert [issue.excel_row for issue in issues] == [3, 4, 5, 6]
    assert "无法识别" in issues[-1].description
    result = analyze_column(column)
    assert result.date_profile.suspected
    assert result.date_profile.ratio == 0.8
    assert result.counts["日期异常"] == 4
    assert not parse_date_value("2026-02-30").parsed
    assert parse_date_value("2024-02-29").parsed == datetime(2024, 2, 29)


def test_date_threshold_does_not_turn_normal_text_or_numbers_into_dates():
    for column in (
        pd.Series(["ABC", "DEF", "GHI"], name="订单日期"),
        pd.Series(["2026-10-10", "ABC", "DEF"], name="日期"),
        pd.Series([20261010, 45123, 1002], name="时间"),
        pd.Series(["20261010", "1002", "00123"], name="date_code"),
    ):
        assert detect_date_anomalies(column) == []
        assert not analyze_column(column).date_profile.suspected
    threshold = pd.Series(["2026-10-10", "2026-10-11", "2026-10-12", "ABC", "DEF", None], name="字段")
    result = analyze_column(threshold)
    assert result.date_profile.ratio == 0.6
    assert result.counts["日期异常"] == 2
    assert len(detect_date_anomalies(pd.Series(["ABC"]), confirmed=True)) == 1


def test_ambiguous_dates_and_time_are_flagged_without_guessing():
    column = pd.Series(["2026-10-10"] * 4 + ["01/02/03", "01/02/2026",
                                           datetime(2026, 10, 10, 15, 30)], dtype=object)
    issues = detect_date_anomalies(column)
    assert len(issues) == 3
    assert all("歧义" in issue.description for issue in issues[:2])
    assert "时间" in issues[2].description


def test_formatted_excel_date_is_loaded_as_datetime_but_bare_serial_is_not_guessed(tmp_path):
    book = Workbook()
    sheet = book.active
    sheet.append(["日期", "序号"])
    sheet.append([datetime(2026, 10, 10), 46305])
    sheet["A2"].number_format = "yyyy/mm/dd"
    path = tmp_path / "Excel日期.xlsx"
    book.save(path)
    book.close()
    data = ExcelFileLoader().load(path).dataframe
    assert analyze_column(data.iloc[:, 0]).date_profile.suspected
    assert not analyze_column(data.iloc[:, 1]).date_profile.suspected


def test_whole_sheet_and_detail_use_identical_rules_and_unique_affected_rows():
    column = pd.Series([1001, "00123", "Ｐ００１", "2026/10/10", None], name="编码", dtype=object)
    full, light = analyze_column(column), analyze_column(column, include_details=False)
    assert full.counts == light.counts
    assert full.affected_rows == light.affected_rows == len({issue.excel_row for issue in full.issues})
    result = analyze_sheet(column.to_frame()).columns[0]
    assert result.numeric_mixture == 1
    assert result.leading_zero == 1
    assert result.fullwidth == 1
    assert result.date_anomalies == 0
    assert result.affected_rows == 3
    assert result.quality_score == 40
    assert light.issues == ()


def test_new_detector_entry_points_share_exact_issue_records_with_main_analysis():
    from app.core.detector import (
        detect_outer_spaces, detect_nbsp, detect_hidden_chars, detect_empty_values, detect_duplicates,
    )
    column = pd.Series(["2026-10-10", "2026/10/10", "2026年10月10日", "2026-10-11",
                        "2026-10-12", "00123", 123, " Ｐ００１\u00a0\t"], dtype=object)
    expected = []
    for function in (detect_outer_spaces, detect_nbsp, detect_hidden_chars, detect_empty_values,
                     detect_duplicates, detect_numeric_type_mixture, detect_leading_zero_risk,
                     detect_date_anomalies, detect_fullwidth_chars):
        expected.extend(function(column, 20))
    result = analyze_column(column, 20)
    assert sorted(expected, key=lambda record: (record.excel_row, record.problem_type)) == sorted(
        result.issues, key=lambda record: (record.excel_row, record.problem_type)
    )
