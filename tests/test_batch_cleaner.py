"""V0.9 多列规则、风险、预览、一次撤销和限量展示。"""

from dataclasses import replace
from unittest.mock import patch

import pandas as pd
import pytest

from app.core.annotation_builder import generate_duplicate_preview
from app.core.batch_cleaner import (
    BATCH_RULE_ORDER, BatchRuleError, generate_batch_preview, recommend_batch_configs,
)
from app.core.cleaner import CleaningRules, generate_cleaning_preview
from app.core.data_manager import DataStateError, WorkbookDataManager
from app.core.file_loader import ExcelPreview
from app.core.sheet_analyzer import ScanCancelled, analyze_sheet
from app.models.cleaning_action import BatchCleaningAction
from app.models.cleaning_rule_config import ColumnCleaningConfig


@pytest.fixture
def frame():
    return pd.DataFrame({
        "商品编码": [" Ｐ001 ", "P002\u00a0", "ABC 123"],
        "金额": ["1002", 0, 12.5],
        "订单日期": ["2026/10/11", "2026-10-12", "2026年10月13日"],
    }, index=[9, 9, 2], dtype=object)


def configs():
    return (
        ColumnCleaningConfig(0, "商品编码", True, ("outer_spaces", "fullwidth", "nbsp")),
        ColumnCleaningConfig(1, "金额", True, ("numeric_conversion",)),
        ColumnCleaningConfig(2, "订单日期", True, ("date_normalize",), date_confirmed=True),
    )


@pytest.fixture
def manager(tmp_path, frame):
    result = WorkbookDataManager()
    result.open_workbook(ExcelPreview(tmp_path / "测试.xlsx", ("订单", "其他"), "订单", frame))
    return result


def prepare(manager, rules=None, **kwargs):
    snapshot = manager.prepare_sheet_scan()
    result = generate_batch_preview(
        snapshot.dataframe, rules or configs(), snapshot.sheet_name,
        original_column_count=manager.original_df.shape[1], **kwargs,
    )
    return manager.prepare_batch_preview(snapshot, result)


def test_multiple_independent_rules_preview_is_readonly_and_rows_are_excel_positions(manager, frame):
    original, working = manager.original_df.copy(deep=True), manager.working_df.copy(deep=True)
    preview = prepare(manager).result
    assert preview.changed_cell_count == 5
    assert preview.changed_row_count == 3
    assert preview.affected_rows == (2, 3, 4)
    assert preview.column_counts == {0: 2, 1: 1, 2: 2}
    assert preview.column_changes[0].after_values.tolist() == ["P001", "P002", "ABC 123"]
    assert preview.column_changes[1].after_values.tolist() == [1002, 0, 12.5]
    assert preview.column_changes[2].after_values.tolist() == ["2026-10-11", "2026-10-12", "2026-10-13"]
    assert all(record.changed for record in preview.changed_records)
    assert len(preview.all_records) == 9
    pd.testing.assert_frame_equal(manager.original_df, original)
    pd.testing.assert_frame_equal(manager.working_df, working)
    pd.testing.assert_frame_equal(frame, original)
    assert not manager.undo_stack


def test_apply_is_one_action_and_one_undo_restores_every_column_dtype_and_rules(manager):
    before = manager.working_df.copy(deep=True)
    action = manager.apply_batch_preview(prepare(manager))
    assert isinstance(action, BatchCleaningAction)
    assert len(manager.undo_stack) == 1
    assert manager.current_data.operation_count == 1
    assert len(action.affected_cells) == 5
    assert set(action.rules_by_column) == {0, 1, 2}
    assert manager.working_df.iat[0, 1] == 1002
    assert manager.working_df.iat[0, 2] == "2026-10-11"
    assert manager.current_data.modified_columns == {0, 1, 2}
    assert set(manager.current_data.column_rules) == {0, 1, 2}
    pd.testing.assert_frame_equal(manager.original_df, before)
    assert manager.undo() is action
    pd.testing.assert_frame_equal(manager.working_df, before)
    assert manager.current_data.operation_count == 0
    assert not manager.current_data.column_rules
    assert not manager.is_modified


@pytest.mark.parametrize("value", ["00123", " 00123 ", "　００１２３　", "\u00a000123"])
def test_leading_zero_blocks_entire_column_including_after_character_rules(value):
    source = pd.DataFrame({"编码": [value, "1002", 0]}, dtype=object)
    config = ColumnCleaningConfig(0, "编码", True, ("numeric_conversion", "fullwidth", "outer_spaces", "nbsp"))
    with pytest.raises(BatchRuleError, match="前导0"):
        generate_batch_preview(source, [config], "订单")
    assert source.iat[1, 0] == "1002"


def test_recommendations_are_per_column_and_never_delete_duplicate_or_fill_empty():
    source = pd.DataFrame({
        "编码": ["00123", 123, None],
        "名称": [" P001", "P001\u00a0", "P001"],
        "金额": ["1002", 2, 0],
        "重复": ["X", "X", ""],
        "全角": ["Ｐ001", "正常", ""],
    }, dtype=object)
    recommended = recommend_batch_configs(source.columns, analyze_sheet(source), original_column_count=5)
    assert "numeric_conversion" not in recommended[0].rules
    assert any("前导0" in warning for warning in recommended[0].warnings)
    assert set(recommended[1].rules) == {"nbsp", "outer_spaces"}
    assert recommended[2].rules == ("numeric_conversion",)
    assert not recommended[3].enabled
    assert any("建议生成辅助列" in warning for warning in recommended[3].warnings)
    assert any("空值" in warning for warning in recommended[3].warnings)
    assert recommended[4].rules == ("fullwidth",)


def test_auxiliary_columns_are_visible_but_cannot_be_cleaned(manager):
    annotation = generate_duplicate_preview(manager.working_df, 0)
    manager.apply_annotation(manager.prepare_annotation_preview(annotation))
    recommended = recommend_batch_configs(
        manager.working_df.columns, analyze_sheet(manager.working_df), original_column_count=3,
    )
    assert recommended[-1].is_generated and not recommended[-1].enabled
    bad = replace(recommended[-1], enabled=True, rules=("outer_spaces",))
    with pytest.raises(BatchRuleError, match="辅助列"):
        generate_batch_preview(manager.working_df, [bad], "订单", original_column_count=3)


def test_numeric_zero_normal_text_and_missing_values_are_safe():
    source = pd.DataFrame({"值": ["ABC 123", 0, "0", 1002, 12.5, None, float("nan"), ""]}, dtype=object)
    before = source.copy(deep=True)
    preview = generate_batch_preview(source, [ColumnCleaningConfig(0, "值", True, ("numeric_conversion",))], "S")
    after = preview.column_changes[0].after_values
    assert after.iloc[:5].tolist() == ["ABC 123", 0, 0, 1002, 12.5]
    assert after.iloc[5] is None and pd.isna(after.iloc[6]) and after.iloc[7] == ""
    assert preview.changed_cell_count == 1
    pd.testing.assert_frame_equal(source, before)


def test_rule_order_is_fixed_even_if_configuration_order_is_reversed():
    config = ColumnCleaningConfig(0, "值", True, tuple(reversed(BATCH_RULE_ORDER[:6])))
    preview = generate_batch_preview(pd.DataFrame({"值": ["\u00a0\t\n\r　Ｐ001　"]}), [config], "S")
    assert preview.column_changes[0].after_values.iloc[0] == "P001"
    text = preview.changed_records[0].rules
    assert text.index("CHAR(160)") < text.index("Tab") < text.index("全角") < text.index("前后")


@pytest.mark.parametrize("rules,confirmed,error", [
    (("numeric_conversion", "date_normalize"), True, "不能同时"),
    (("date_normalize",), False, "确认"),
    (("delete_duplicates",), False, "不支持"),
    ((), False, "至少一项"),
])
def test_conflicting_or_unsupported_rules_are_rejected(rules, confirmed, error):
    config = ColumnCleaningConfig(0, "值", True, rules, date_confirmed=confirmed)
    with pytest.raises(BatchRuleError, match=error):
        generate_batch_preview(pd.DataFrame({"值": ["1002"]}), [config], "S")


def test_repeated_run_is_noop_without_new_history(manager):
    manager.apply_batch_preview(prepare(manager))
    before = manager.working_df.copy(deep=True)
    second = prepare(manager)
    assert second.result.changed_cell_count == 0
    assert manager.apply_batch_preview(second) is None
    assert len(manager.undo_stack) == 1
    pd.testing.assert_frame_equal(manager.working_df, before)


def test_limited_samples_still_apply_all_changes_and_undo(manager):
    frame = pd.DataFrame({"A": [" X "] * 2500, "B": [" Y "] * 2500})
    manager.open_workbook(ExcelPreview(manager.file_path, ("订单",), "订单", frame))
    rules = (ColumnCleaningConfig(0, "A", True, ("outer_spaces",)),
             ColumnCleaningConfig(1, "B", True, ("outer_spaces",)))
    preview = prepare(manager, rules)
    assert len(preview.result.changed_records) == len(preview.result.all_records) == 1000
    assert preview.result.changed_cell_count == 5000
    assert preview.result.changed_row_count == 2500
    manager.apply_batch_preview(preview)
    assert manager.working_df.iat[2499, 1] == "Y"
    manager.undo()
    pd.testing.assert_frame_equal(manager.working_df, frame)


def test_cancel_and_progress_never_publish_partial_result(frame):
    events = []
    with pytest.raises(ScanCancelled):
        generate_batch_preview(frame, configs(), "S",
                               progress=lambda *event: events.append(event),
                               is_cancelled=lambda: bool(events))
    assert events == [(1, 3, "商品编码")]


def test_stale_preview_and_stale_worker_snapshot_are_rejected(manager):
    snapshot = manager.prepare_sheet_scan()
    preview = prepare(manager)
    manager.apply_batch_preview(preview)
    with pytest.raises(DataStateError, match="失效"):
        manager.apply_batch_preview(preview)
    with pytest.raises(DataStateError, match="失效"):
        manager.prepare_batch_preview(snapshot, preview.result)


def test_truncated_preview_data_is_rejected(manager):
    snapshot = manager.prepare_sheet_scan()
    preview = prepare(manager).result
    change = preview.column_changes[0]
    truncated = replace(change, after_values=change.after_values.iloc[:-1])
    invalid = replace(preview, column_changes=(truncated,) + preview.column_changes[1:])
    with pytest.raises(DataStateError, match="不一致"):
        manager.prepare_batch_preview(snapshot, invalid)


def test_failed_second_assignment_does_not_partially_apply(manager):
    preview = prepare(manager)
    before = manager.working_df.copy(deep=True)
    original = pd.DataFrame.isetitem
    calls = []
    def fail_second(dataframe, index, value):
        calls.append(index)
        if len(calls) == 2:
            raise MemoryError("测试内存错误")
        return original(dataframe, index, value)
    with patch.object(pd.DataFrame, "isetitem", fail_second), pytest.raises(MemoryError):
        manager.apply_batch_preview(preview)
    pd.testing.assert_frame_equal(manager.working_df, before)
    assert not manager.undo_stack and manager.current_data.revision == 0


def test_restore_and_undo_restore_keep_batch_as_one_history_entry(manager):
    manager.apply_batch_preview(prepare(manager))
    cleaned = manager.working_df.copy(deep=True)
    manager.restore_original()
    pd.testing.assert_frame_equal(manager.working_df, manager.original_df)
    manager.undo()
    pd.testing.assert_frame_equal(manager.working_df, cleaned)
    manager.undo()
    pd.testing.assert_frame_equal(manager.working_df, manager.original_df)


def test_duplicate_headers_and_indices_keep_physical_column_identity(tmp_path):
    frame = pd.DataFrame([[" A ", " B "], [" C ", " D "]], columns=["同名", "同名"], index=[7, 7])
    manager = WorkbookDataManager()
    manager.open_workbook(ExcelPreview(tmp_path / "同名.xlsx", ("S",), "S", frame))
    rules = [ColumnCleaningConfig(0, "同名", True, ("outer_spaces",)),
             ColumnCleaningConfig(1, "同名", True, ("outer_spaces",))]
    preview = prepare(manager, rules)
    manager.apply_batch_preview(preview)
    assert manager.working_df.iloc[0].tolist() == ["A", "B"]
    assert preview.result.affected_rows == (2, 3)
    manager.undo()
    pd.testing.assert_frame_equal(manager.working_df, frame)


def test_batch_undo_returns_to_its_sheet_and_preserves_other_history(manager, frame):
    preview = generate_cleaning_preview(manager.working_df.iloc[:, 0], CleaningRules(outer_spaces=True))
    manager.apply_preview(manager.prepare_preview(0, preview))
    first_clean = manager.working_df.copy(deep=True)
    manager.apply_batch_preview(prepare(manager))
    manager.add_sheet(ExcelPreview(manager.file_path, ("订单", "其他"), "其他", frame))
    manager.undo()
    assert manager.current_sheet == "订单"
    assert len(manager.undo_stack) == 1
    pd.testing.assert_frame_equal(manager.working_df, first_clean)
