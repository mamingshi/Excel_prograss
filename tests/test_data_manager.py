"""验证工作副本、预览校验、多列多 Sheet 历史和原始快照保护。"""

from dataclasses import replace
from unittest.mock import patch

import pandas as pd
import pytest

from app.core.cleaner import CleaningRules, generate_cleaning_preview
from app.core.data_manager import DataStateError, WorkbookDataManager
from app.core.file_loader import ExcelPreview
from app.models.cleaning_action import CleaningAction, RestoreAction


RULES = CleaningRules(nbsp=True, tab=True, newline=True, carriage_return=True, outer_spaces=True)


@pytest.fixture
def source():
    return pd.DataFrame({
        "A": [" P001 ", "P002\u00a0", "ABC\t123", 1002, 0, 12.5, None, float("nan"), "", "DUP", "DUP", "0"],
        "B": [" B001 ", "B002\u00a0", "A\nB"] + ["正常"] * 9,
    }, dtype=object)


@pytest.fixture
def manager(tmp_path, source):
    data = WorkbookDataManager()
    data.open_workbook(ExcelPreview(tmp_path / "测试.xlsx", ("S1", "S2"), "S1", source))
    return data


def prepare(manager, column=0, rules=RULES):
    result = generate_cleaning_preview(manager.working_df.iloc[:, column], rules)
    return manager.prepare_preview(column, result)


def test_original_and_source_are_independent_from_working(manager, source):
    saved = source.copy(deep=True)
    assert manager.original_df is not manager.working_df
    assert manager.original_df is not source
    action = manager.apply_preview(prepare(manager))
    assert isinstance(action, CleaningAction)
    assert action.affected_indices == (0, 1, 2)
    assert manager.working_df.iloc[:3, 0].tolist() == ["P001", "P002", "ABC 123"]
    assert action.before_values.iloc[0] == " P001 "
    assert action.after_values.iloc[0] == "P001"
    assert action.timestamp.tzinfo is not None
    pd.testing.assert_frame_equal(manager.original_df, saved)
    pd.testing.assert_frame_equal(source, saved)
    assert manager.is_modified
    manager.undo()
    pd.testing.assert_frame_equal(manager.working_df, saved)
    assert not manager.is_modified
    assert not manager.undo_stack


def test_two_columns_undo_in_reverse_order(manager):
    original = manager.original_df.copy(deep=True)
    manager.apply_preview(prepare(manager, 0))
    first = manager.working_df.copy(deep=True)
    manager.apply_preview(prepare(manager, 1))
    assert manager.current_data.operation_count == 2
    assert manager.undo().column_index == 1
    pd.testing.assert_frame_equal(manager.working_df, first)
    assert manager.current_data.operation_count == 1
    assert manager.undo().column_index == 0
    pd.testing.assert_frame_equal(manager.working_df, original)
    assert manager.current_data.operation_count == 0


def test_numeric_missing_and_duplicate_values_are_preserved(manager):
    before = manager.working_df.iloc[:, 0].copy(deep=True)
    manager.apply_preview(prepare(manager))
    after = manager.working_df.iloc[:, 0]
    for position in (3, 4, 5, 11):
        assert after.iloc[position] == before.iloc[position]
        assert type(after.iloc[position]) is type(before.iloc[position])
    assert after.iloc[6] is None
    assert pd.isna(after.iloc[7])
    assert after.iloc[8] == ""
    assert after.iloc[9:11].tolist() == ["DUP", "DUP"]
    assert len(manager.working_df) == len(before)
    manager.undo()
    pd.testing.assert_series_equal(manager.working_df.iloc[:, 0], before)


def test_reapplying_same_rules_is_a_noop(manager):
    old_preview = prepare(manager)
    manager.apply_preview(old_preview)
    after = manager.working_df.copy(deep=True)
    with pytest.raises(DataStateError, match="失效"):
        manager.apply_preview(old_preview)
    current = prepare(manager)
    assert current.result.changed_count == 0
    assert manager.apply_preview(current) is None
    assert len(manager.undo_stack) == 1
    assert manager.current_data.operation_count == 1
    pd.testing.assert_frame_equal(manager.working_df, after)


def test_restore_all_columns_and_undo_restore(manager):
    manager.apply_preview(prepare(manager, 0))
    manager.apply_preview(prepare(manager, 1))
    cleaned = manager.working_df.copy(deep=True)
    action = manager.restore_original()
    assert isinstance(action, RestoreAction)
    pd.testing.assert_frame_equal(manager.working_df, manager.original_df)
    assert not manager.is_modified
    assert manager.current_data.operation_count == 0
    manager.undo()
    pd.testing.assert_frame_equal(manager.working_df, cleaned)
    assert manager.current_data.operation_count == 2
    manager.undo()
    manager.undo()
    pd.testing.assert_frame_equal(manager.working_df, manager.original_df)
    assert not manager.undo_stack


def test_per_sheet_copies_and_global_history(manager, source):
    manager.apply_preview(prepare(manager))
    first_cleaned = manager.working_df.copy(deep=True)
    manager.add_sheet(ExcelPreview(manager.file_path, manager.sheet_names, "S2", source))
    manager.apply_preview(prepare(manager, 1))
    second_cleaned = manager.working_df.copy(deep=True)
    manager.select_sheet("S1")
    pd.testing.assert_frame_equal(manager.working_df, first_cleaned)
    assert manager.undo().sheet_name == "S2"
    assert manager.current_sheet == "S2"
    pd.testing.assert_frame_equal(manager.working_df, source)
    assert manager.is_modified  # S1 的修改仍保留。
    manager.undo()
    assert manager.current_sheet == "S1"
    assert not manager.is_modified
    manager.select_sheet("S2")
    assert not manager.working_df.equals(second_cleaned)


def test_restore_affects_only_current_sheet(manager, source):
    manager.apply_preview(prepare(manager))
    first_cleaned = manager.working_df.copy(deep=True)
    manager.add_sheet(ExcelPreview(manager.file_path, manager.sheet_names, "S2", source))
    manager.apply_preview(prepare(manager, 1))
    manager.restore_original()
    pd.testing.assert_frame_equal(manager.working_df, source)
    pd.testing.assert_frame_equal(manager.sheet_data["S1"].working_df, first_cleaned)
    assert manager.is_modified


@pytest.mark.parametrize("same_path", [True, False])
def test_opening_file_resets_history_and_invalidates_preview(manager, source, tmp_path, same_path):
    manager.apply_preview(prepare(manager))
    pending = prepare(manager, 1)
    path = manager.file_path if same_path else tmp_path / "新文件.xlsx"
    manager.open_workbook(ExcelPreview(path, ("新表",), "新表", source))
    assert list(manager.sheet_data) == ["新表"]
    assert manager.undo() is None
    assert not manager.is_modified
    assert not manager.is_preview_current(pending)
    with pytest.raises(DataStateError):
        manager.apply_preview(pending)
    pd.testing.assert_frame_equal(manager.working_df, source)


def test_preview_for_another_sheet_is_rejected(manager, source):
    pending = prepare(manager)
    manager.add_sheet(ExcelPreview(manager.file_path, manager.sheet_names, "S2", source))
    with pytest.raises(DataStateError, match="失效"):
        manager.apply_preview(pending)
    assert not manager.undo_stack


def test_old_snapshot_cannot_be_staged_after_apply(manager):
    old = prepare(manager)
    manager.apply_preview(old)
    with pytest.raises(DataStateError, match="不一致"):
        manager.prepare_preview(0, old.result)


def test_preview_shape_and_display_must_match(manager):
    preview = prepare(manager)
    wrong = replace(preview.result, column_name="错误列")
    with pytest.raises(DataStateError, match="不匹配"):
        manager.prepare_preview(0, wrong)
    preview.result.preview_series.iloc[0] = "未展示的值"
    with pytest.raises(DataStateError, match="不一致"):
        manager.prepare_preview(0, preview.result)


def test_failed_assignment_leaves_data_and_history_unchanged(manager):
    saved = manager.working_df.copy(deep=True)
    pending = prepare(manager)
    with patch.object(manager, "_replace_column", side_effect=RuntimeError("模拟赋值失败")):
        with pytest.raises(RuntimeError):
            manager.apply_preview(pending)
    pd.testing.assert_frame_equal(manager.working_df, saved)
    assert not manager.undo_stack
    assert manager.is_preview_current(pending)
    manager.apply_preview(pending)
    cleaned = manager.working_df.copy(deep=True)
    with patch.object(manager, "_replace_column", side_effect=RuntimeError("模拟撤销失败")):
        with pytest.raises(RuntimeError):
            manager.undo()
    pd.testing.assert_frame_equal(manager.working_df, cleaned)
    assert len(manager.undo_stack) == 1


def test_duplicate_index_and_column_names_use_positions(tmp_path):
    frame = pd.DataFrame([[" A ", " B "], [" C ", " D "]], columns=["编码", "编码"], index=["r", "r"], dtype=object)
    manager = WorkbookDataManager()
    manager.open_workbook(ExcelPreview(tmp_path / "测试.xlsx", ("表",), "表", frame))
    manager.apply_preview(prepare(manager, 1))
    assert manager.working_df.iloc[:, 0].tolist() == [" A ", " C "]
    assert manager.working_df.iloc[:, 1].tolist() == ["B", "D"]
    manager.undo()
    pd.testing.assert_frame_equal(manager.working_df, frame)


@pytest.mark.parametrize("dtype,values", [("int64", [0, 1002]), ("float64", [12.5, float("nan")]), ("string", [" P001 ", pd.NA])])
def test_column_dtype_is_preserved(tmp_path, dtype, values):
    frame = pd.DataFrame({"A": pd.Series(values, dtype=dtype)})
    manager = WorkbookDataManager()
    manager.open_workbook(ExcelPreview(tmp_path / "测试.xlsx", ("表",), "表", frame))
    manager.apply_preview(prepare(manager))
    assert manager.working_df["A"].dtype == frame["A"].dtype
    if dtype == "string":
        assert manager.working_df.iat[0, 0] == "P001"
        manager.undo()
    pd.testing.assert_frame_equal(manager.working_df, frame)


def test_empty_history_and_unmodified_restore_are_safe(manager):
    assert manager.undo() is None
    assert manager.restore_original() is None
    assert not manager.undo_stack
    empty = WorkbookDataManager()
    assert empty.undo() is None
    with pytest.raises(DataStateError, match="请先打开"):
        empty.restore_original()
    with pytest.raises(DataStateError):
        manager.select_sheet("不存在")
