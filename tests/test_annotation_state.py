"""辅助列工作状态、历史、恢复、源字段变化联动和失败原子性。"""

from unittest.mock import patch

import pandas as pd
import pytest

from app.core.annotation_builder import generate_duplicate_preview
from app.core.cleaner import CleaningRules, generate_cleaning_preview
from app.core.data_manager import DataStateError, WorkbookDataManager
from app.core.exporter import ExcelExportError, prepare_export
from app.core.file_loader import ExcelPreview
from app.core.sheet_analyzer import analyze_sheet
from app.models.cleaning_action import AnnotationAction


@pytest.fixture
def manager(tmp_path):
    manager = WorkbookDataManager()
    manager.open_workbook(ExcelPreview(
        tmp_path / "原始.xlsx", ("订单", "其他"), "订单",
        pd.DataFrame({"商品编码": [" P001", "P002", "P001", "P001"], "金额": [1, 2, 3, 4]}),
    ))
    return manager


def preview(manager, column=0):
    return manager.prepare_annotation_preview(generate_duplicate_preview(manager.working_df, column))


def apply(manager, column=0):
    return manager.apply_annotation(preview(manager, column))


def clean(manager):
    result = generate_cleaning_preview(manager.working_df.iloc[:, 0], CleaningRules(outer_spaces=True))
    return manager.apply_preview(manager.prepare_preview(0, result))


def test_preview_alone_does_not_change_data_history_or_export(manager):
    original = manager.original_df.copy(deep=True)
    working = manager.working_df.copy(deep=True)
    prepared = preview(manager)
    assert prepared.result.values.tolist() == ["正常", "正常", "重复", "重复"]
    pd.testing.assert_frame_equal(manager.original_df, original)
    pd.testing.assert_frame_equal(manager.working_df, working)
    assert not manager.undo_stack
    assert not manager.is_modified
    with pytest.raises(ExcelExportError):
        prepare_export(manager.get_export_snapshot())


def test_apply_appends_only_to_working_and_undo_removes_exact_column(manager):
    original = manager.original_df.copy(deep=True)
    action = apply(manager)
    assert isinstance(action, AnnotationAction)
    assert manager.working_df.columns.tolist() == ["商品编码", "金额", "商品编码_重复检查"]
    assert manager.working_df.iloc[:, -1].tolist() == ["正常", "正常", "重复", "重复"]
    assert manager.is_modified and manager.current_data.operation_count == 1
    pd.testing.assert_frame_equal(manager.working_df.iloc[:, :2], original)
    pd.testing.assert_frame_equal(manager.original_df, original)
    plan = prepare_export(manager.get_export_snapshot())
    assert plan.changes == ()
    assert plan.annotation_column_count == 1
    assert plan.sheet_count == 1
    assert manager.undo() is action
    pd.testing.assert_frame_equal(manager.working_df, original)
    assert not manager.current_data.annotations
    assert not manager.is_modified and not manager.undo_stack


def test_multiple_additions_get_unique_names_and_lifo_undo(manager):
    first = apply(manager)
    second = apply(manager)
    third = apply(manager)
    assert manager.working_df.columns[-3:].tolist() == [
        "商品编码_重复检查", "商品编码_重复检查_2", "商品编码_重复检查_3",
    ]
    assert manager.undo() is third
    assert manager.undo() is second
    assert manager.working_df.columns[-1] == first.column_name
    assert manager.undo() is first
    assert manager.working_df.shape[1] == 2


def test_restore_removes_all_helpers_and_undo_restore_restores_them(manager):
    apply(manager)
    clean(manager)
    apply(manager)
    before = manager.working_df.copy(deep=True)
    annotations = manager.current_data.annotations
    manager.restore_original()
    pd.testing.assert_frame_equal(manager.working_df, manager.original_df)
    assert manager.current_data.annotations == ()
    assert not manager.is_modified
    manager.undo()
    pd.testing.assert_frame_equal(manager.working_df, before)
    assert manager.current_data.annotations == annotations
    assert manager.current_data.operation_count == 3


def test_source_cleaning_and_undo_recalculate_helper_labels(manager):
    apply(manager)
    clean(manager)
    assert manager.working_df.iloc[:, -1].tolist() == ["重复", "正常", "重复", "重复"]
    manager.undo()
    assert manager.working_df.iloc[:, -1].tolist() == ["正常", "正常", "重复", "重复"]
    manager.undo()
    assert manager.working_df.shape[1] == 2


def test_clean_then_annotate_and_reverse_undo_preserve_original_edit(manager):
    clean(manager)
    cleaned = manager.working_df.copy(deep=True)
    apply(manager)
    manager.undo()
    pd.testing.assert_frame_equal(manager.working_df, cleaned)
    assert manager.current_data.is_modified
    assert manager.current_data.modified_columns == {0}
    manager.undo()
    pd.testing.assert_frame_equal(manager.working_df, manager.original_df)


def test_annotation_invalidates_overview_cache(manager):
    snapshot = manager.prepare_sheet_scan()
    manager.save_sheet_scan(snapshot, analyze_sheet(snapshot.dataframe))
    cache = manager.current_data.scan_cache
    apply(manager)
    assert cache.revision != manager.current_data.revision
    manager.undo()
    assert cache.revision != manager.current_data.revision


def test_stale_preview_after_apply_sheet_switch_or_reopen_is_rejected(manager):
    prepared = preview(manager)
    clean(manager)
    with pytest.raises(DataStateError, match="失效"):
        manager.apply_annotation(prepared)
    prepared = preview(manager)
    manager.add_sheet(ExcelPreview(manager.file_path, manager.sheet_names, "其他", pd.DataFrame({"字段": ["A"]})))
    with pytest.raises(DataStateError, match="失效"):
        manager.apply_annotation(prepared)
    manager.open_workbook(ExcelPreview(manager.file_path, manager.sheet_names, "订单", manager.sheet_data["订单"].original_df))
    with pytest.raises(DataStateError, match="失效"):
        manager.apply_annotation(prepared)


def test_helpers_cannot_be_cleaned_or_annotated_again(manager):
    apply(manager)
    result = generate_cleaning_preview(manager.working_df.iloc[:, 2], CleaningRules())
    with pytest.raises(DataStateError, match="不能直接清洗"):
        manager.prepare_preview(2, result)
    with pytest.raises(DataStateError, match="原始字段"):
        preview(manager, 2)


def test_each_sheet_keeps_helpers_and_cross_sheet_undo_returns_to_target(manager):
    apply(manager)
    manager.add_sheet(ExcelPreview(manager.file_path, manager.sheet_names, "其他", pd.DataFrame({"客户": ["甲", "甲"]})))
    apply(manager)
    assert manager.undo().sheet_name == "其他"
    assert manager.working_df.shape[1] == 1
    manager.undo()
    assert manager.current_sheet == "订单"
    assert manager.working_df.shape[1] == 2


def test_apply_failure_does_not_change_state_or_history(manager):
    prepared = preview(manager)
    before = manager.working_df.copy(deep=True)
    with patch.object(pd.DataFrame, "insert", side_effect=RuntimeError("模拟内存错误")):
        with pytest.raises(RuntimeError):
            manager.apply_annotation(prepared)
    pd.testing.assert_frame_equal(manager.working_df, before)
    assert not manager.undo_stack and not manager.current_data.annotations


def test_tampered_preview_is_rejected(manager):
    prepared = preview(manager)
    prepared.result.values.iloc[0] = "重复"
    with pytest.raises(DataStateError, match="不一致"):
        manager.apply_annotation(prepared)
    assert not manager.is_modified


def test_snapshot_is_independent_of_later_undo(manager):
    apply(manager)
    snapshot = manager.get_export_snapshot()
    manager.undo()
    plan = prepare_export(snapshot)
    assert plan.annotation_column_count == 1
    assert snapshot.sheets[0].working_df.shape[1] == 3
