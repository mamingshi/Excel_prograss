"""工作副本扫描快照、各 Sheet 缓存与撤销/恢复后的版本失效。"""

import pandas as pd
import pytest

from app.core.cleaner import CleaningRules, generate_cleaning_preview
from app.core.data_manager import DataStateError, WorkbookDataManager
from app.core.file_loader import ExcelPreview
from app.core.sheet_analyzer import analyze_sheet


@pytest.fixture
def manager(tmp_path):
    manager = WorkbookDataManager()
    manager.open_workbook(ExcelPreview(tmp_path / "测试.xlsx", ("甲", "乙"), "甲", pd.DataFrame({"编码": [" P001", "P002"]})))
    return manager


def scan(manager):
    snapshot = manager.prepare_sheet_scan()
    result = analyze_sheet(snapshot.dataframe)
    manager.save_sheet_scan(snapshot, result)
    return snapshot, result


def clean(manager):
    preview = generate_cleaning_preview(manager.working_df.iloc[:, 0], CleaningRules(outer_spaces=True))
    return manager.apply_preview(manager.prepare_preview(0, preview))


def test_scans_working_copy_and_application_invalidates_cache(manager):
    snapshot, before = scan(manager)
    assert before.columns[0].outer_spaces == 1
    cache = manager.current_data.scan_cache
    clean(manager)
    assert manager.current_data.scan_cache is cache
    assert cache.revision != manager.current_data.revision
    assert not manager.is_sheet_scan_current(snapshot)
    _, after = scan(manager)
    assert after.columns[0].outer_spaces == 0
    assert after.columns[0].quality_score == 100
    assert manager.original_df.iat[0, 0] == " P001"


def test_undo_restore_and_undo_restore_invalidate_cached_results(manager):
    clean(manager)
    scan(manager)
    cache = manager.current_data.scan_cache
    manager.restore_original()
    assert cache.revision != manager.current_data.revision
    scan(manager)
    cache = manager.current_data.scan_cache
    manager.undo()
    assert cache.revision != manager.current_data.revision
    scan(manager)
    cache = manager.current_data.scan_cache
    manager.undo()
    assert cache.revision != manager.current_data.revision


def test_no_changes_and_preview_only_do_not_invalidate_cache(manager):
    snapshot, _ = scan(manager)
    no_change = generate_cleaning_preview(manager.working_df.iloc[:, 0], CleaningRules())
    prepared = manager.prepare_preview(0, no_change)
    assert manager.apply_preview(prepared) is None
    assert manager.is_sheet_scan_current(snapshot)
    assert manager.current_data.scan_cache.revision == manager.current_data.revision


def test_each_sheet_keeps_its_own_result_and_reopening_clears_all(manager):
    _, first_result = scan(manager)
    other = ExcelPreview(manager.file_path, manager.sheet_names, "乙", pd.DataFrame({"名称": ["正常"]}))
    manager.add_sheet(other)
    _, other_result = scan(manager)
    manager.select_sheet("甲")
    assert manager.current_data.scan_cache.result is first_result
    clean(manager)
    manager.select_sheet("乙")
    assert manager.current_data.scan_cache.result is other_result
    assert manager.current_data.scan_cache.revision == manager.current_data.revision
    manager.open_workbook(other)
    assert list(manager.sheet_data) == ["乙"]
    assert manager.current_data.scan_cache is None


def test_old_results_rejected_after_edit_switch_and_new_workbook(manager):
    snapshot, result = scan(manager)
    clean(manager)
    with pytest.raises(DataStateError, match="变化"):
        manager.save_sheet_scan(snapshot, result)
    snapshot = manager.prepare_sheet_scan()
    manager.add_sheet(ExcelPreview(manager.file_path, manager.sheet_names, "乙", pd.DataFrame({"名称": ["正常"]})))
    with pytest.raises(DataStateError, match="变化"):
        manager.save_sheet_scan(snapshot, result)
    manager.open_workbook(ExcelPreview(manager.file_path, manager.sheet_names, "甲", pd.DataFrame({"编码": [" P001", "P002"]})))
    with pytest.raises(DataStateError, match="变化"):
        manager.save_sheet_scan(snapshot, result)


def test_scan_snapshot_stays_stable_when_manager_replaces_dataframe(manager):
    snapshot = manager.prepare_sheet_scan()
    before = snapshot.dataframe.copy(deep=True)
    clean(manager)
    pd.testing.assert_frame_equal(snapshot.dataframe, before)
    assert manager.working_df.iat[0, 0] == "P001"


def test_no_file_cannot_start_scan():
    with pytest.raises(DataStateError, match="打开"):
        WorkbookDataManager().prepare_sheet_scan()
