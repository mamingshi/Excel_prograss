"""整表扫描规则、质量分、唯一行计数、取消及只读保证。"""

from datetime import datetime
from unittest.mock import patch

import pandas as pd
import pytest

from app.core.detector import PROBLEM_TYPES, analyze_column
from app.core.sheet_analyzer import ScanCancelled, analyze_sheet, quality_grade, quality_score


def test_multicolumn_scan_reuses_detector_and_reports_progress():
    frame = pd.DataFrame({
        "商品编码": [" P001", "P002", "P003"],
        "客户名称": ["甲", "乙", "丙"],
        "备注": ["正常", "需\n确认", "末尾\u00a0"],
    })
    events = []
    with patch("app.core.sheet_analyzer.analyze_column", wraps=analyze_column) as detector:
        result = analyze_sheet(frame, progress=lambda *event: events.append(event))
    assert detector.call_count == 3
    assert all(call.kwargs == {"include_details": False} for call in detector.call_args_list)
    code, customer, remark = result.columns
    assert (code.outer_spaces, code.affected_rows, code.quality_score) == (1, 1, 67)
    assert customer.quality_score == 100
    assert customer.summary == "未发现当前支持的问题"
    assert (remark.nbsp, remark.hidden_chars, remark.affected_rows) == (1, 1, 2)
    assert result.problem_column_count == 2
    assert result.normal_column_count == 1
    assert events == [(1, 3, "商品编码"), (2, 3, "客户名称"), (3, 3, "备注")]


@pytest.mark.parametrize("values,affected,score,status", [
    (["P001", "P002", "P003"], 0, 100, "优秀"),
    ([1001, 1002, 1003], 0, 100, "优秀"),
    (["P001", "ABC 123", 1002, 0, "0"], 1, 80, "建议清洗"),
    ([datetime(2026, 1, 1), datetime(2026, 1, 2)], 0, 100, "优秀"),
    ([True, False], 0, 100, "优秀"),
])
def test_normal_and_nonstring_columns_have_type_aware_statistics(values, affected, score, status):
    result = analyze_sheet(pd.DataFrame({"字段": pd.Series(values, dtype=object)})).columns[0]
    assert result.quality_score == score
    assert result.affected_rows == affected
    assert result.status == status


def test_all_empty_column_excludes_missing_from_duplicates():
    frame = pd.DataFrame({"全空": pd.Series([None, pd.NA, float("nan"), "", pd.NaT], dtype=object)})
    column = analyze_sheet(frame).columns[0]
    assert column.empty == 5
    assert column.duplicates == 0
    assert column.affected_rows == 5
    assert column.quality_score == 0
    assert column.status == "问题较多"


def test_duplicate_records_include_first_and_preserve_value_types():
    column = analyze_sheet(pd.DataFrame({
        "字段": pd.Series(["A", "A", None, "", float("nan"), 0, "0", "A ", " B\u00a0\t\n "], dtype=object),
    })).columns[0]
    assert column.duplicates == 2
    assert column.outer_spaces == 2
    assert column.nbsp == 1
    assert column.hidden_chars == 1
    assert column.empty == 3
    assert column.numeric_mixture == 1
    assert column.affected_rows == 8
    assert column.quality_score == 11


def test_multiple_problems_on_same_cell_count_only_one_affected_row():
    value = " \u00a0P001\t\n\r "
    result = analyze_sheet(pd.DataFrame({"字段": [value, value, "P002", "P003"]})).columns[0]
    assert (result.outer_spaces, result.nbsp, result.hidden_chars, result.duplicates) == (2, 2, 2, 2)
    assert result.affected_rows == 2
    assert result.quality_score == 50


@pytest.mark.parametrize(("affected", "total", "expected"), [
    (0, 0, 100), (0, 100, 100), (1, 100, 99), (18, 100, 82),
    (100, 100, 0), (200, 100, 0), (-1, 100, 100),
])
def test_score_is_bounded(affected, total, expected):
    assert quality_score(affected, total) == expected
    assert 0 <= quality_score(affected, total) <= 100


@pytest.mark.parametrize(("score", "expected"), [
    (100, "优秀"), (95, "优秀"), (94, "良好"), (85, "良好"),
    (84, "建议清洗"), (70, "建议清洗"), (69, "问题较多"), (0, "问题较多"),
])
def test_quality_grade_boundaries(score, expected):
    assert quality_grade(score) == expected


def test_empty_table_and_header_only_sheet():
    assert analyze_sheet(pd.DataFrame()).columns == ()
    result = analyze_sheet(pd.DataFrame(columns=["编码", "名称"]))
    assert result.row_count == 0
    assert result.column_count == 2
    assert result.normal_column_count == 2
    assert all(column.quality_score == 100 for column in result.columns)


def test_duplicate_headers_and_nondefault_indexes_use_physical_positions():
    frame = pd.DataFrame([["正常", " 有空格"], ["不同", "\u00a0异常"]], columns=["重名", "重名"], index=[88, 88])
    result = analyze_sheet(frame)
    assert [column.column_index for column in result.columns] == [0, 1]
    assert [column.column_name for column in result.columns] == ["重名", "重名"]
    assert result.columns[0].affected_rows == 0
    assert result.columns[1].affected_rows == 2


def test_summary_matches_complete_column_analysis_without_retaining_issues():
    values = [" A ", "\u00a0", "A\t\n\r", None, "", 0, "0", "DUP", "DUP", pd.NA, float("nan")]
    series = pd.Series(values, dtype=object, name="字段", index=[5] * len(values))
    before = series.copy(deep=True)
    full = analyze_column(series)
    summary = analyze_column(series, include_details=False)
    assert summary.counts == full.counts
    assert summary.affected_rows == full.affected_rows == len({issue.excel_row for issue in full.issues})
    assert summary.issues == ()
    assert summary.duplicate_groups == ()
    pd.testing.assert_series_equal(series, before)


def test_all_rule_entry_points_still_agree():
    from app.core.detector import (
        detect_outer_spaces, detect_nbsp, detect_hidden_chars, detect_empty_values, detect_duplicates,
        detect_numeric_type_mixture, detect_leading_zero_risk, detect_date_anomalies, detect_fullwidth_chars,
    )
    series = pd.Series([" \u00a0A\t\n\r ", None, "", "DUP", "DUP", 1, 1.0, "1", "ABC 123"], dtype=object)
    expected = []
    for function in (detect_outer_spaces, detect_nbsp, detect_hidden_chars, detect_empty_values, detect_duplicates,
                     detect_numeric_type_mixture, detect_leading_zero_risk, detect_date_anomalies, detect_fullwidth_chars):
        expected.extend(function(series, 20))
    expected.sort(key=lambda issue: issue.excel_row)
    result = analyze_column(series, 20)
    assert result.issues == tuple(expected)
    assert result.counts == {name: sum(issue.problem_type == name for issue in expected) for name in PROBLEM_TYPES}


def test_scan_never_modifies_dataframe():
    frame = pd.DataFrame({
        "文本": [" A ", "B\u00a0", "X\tY", None],
        "数字": pd.Series([1, pd.NA, 2, 3], dtype="Int64"),
        "日期": [datetime(2026, 1, 1), pd.NaT, datetime(2026, 1, 2), datetime(2026, 1, 3)],
        "布尔": [True, False, True, False],
    })
    before = frame.copy(deep=True)
    analyze_sheet(frame)
    pd.testing.assert_frame_equal(frame, before)


def test_cancellation_never_returns_partial_results():
    frame = pd.DataFrame({"甲": [" A"], "乙": ["B"], "丙": ["C"]})
    completed = []
    with pytest.raises(ScanCancelled):
        analyze_sheet(
            frame, progress=lambda done, total, name: completed.append(done),
            is_cancelled=lambda: bool(completed),
        )
    assert completed == [1]


def test_cancellation_after_last_column_still_discards_result():
    completed = []
    with pytest.raises(ScanCancelled):
        analyze_sheet(
            pd.DataFrame({"甲": ["A"]}),
            progress=lambda *event: completed.append(event),
            is_cancelled=lambda: bool(completed),
        )


def test_pre_cancelled_scan_does_not_call_detector():
    with patch("app.core.sheet_analyzer.analyze_column") as detector:
        with pytest.raises(ScanCancelled):
            analyze_sheet(pd.DataFrame({"甲": ["A"]}), is_cancelled=lambda: True)
    detector.assert_not_called()


def test_detector_failure_does_not_mutate_input_or_return_partial_scan():
    frame = pd.DataFrame({"甲": [" A"], "乙": ["B"]})
    before = frame.copy(deep=True)
    with patch("app.core.sheet_analyzer.analyze_column", side_effect=RuntimeError("测试失败")):
        with pytest.raises(RuntimeError, match="测试失败"):
            analyze_sheet(frame)
    pd.testing.assert_frame_equal(frame, before)
