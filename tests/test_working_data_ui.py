"""真实工作簿上的确认、应用、撤销、恢复和 Sheet 状态验证。"""

import hashlib
from unittest.mock import patch

from openpyxl import Workbook
import pandas as pd
from PySide6.QtCore import Qt
from PySide6.QtTest import QTest
from PySide6.QtWidgets import QMessageBox

from app.ui.main_window import MainWindow
from tests.test_cleaning_ui import window, analyze, generate, wait_for_idle, click_page


def apply(window):
    with patch("app.ui.main_window.confirm_apply", return_value=True) as confirmation:
        QTest.mouseClick(window.cleaning_preview_panel.apply_button, Qt.MouseButton.LeftButton)
    confirmation.assert_called_once()
    wait_for_idle(window)


def undo(window):
    QTest.mouseClick(window.undo_button, Qt.MouseButton.LeftButton)
    wait_for_idle(window)


def test_apply_refreshes_working_data_and_diagnostics_without_writing_file(window):
    original = window._preview.dataframe.copy(deep=True)
    digest = hashlib.sha256(window._preview.file_path.read_bytes()).digest()
    assert not window.undo_button.isEnabled()
    assert not window.restore_button.isEnabled()
    assert not window.cleaning_preview_panel.apply_button.isEnabled()
    analyze(window)
    generate(window)
    assert window.cleaning_preview_panel.apply_button.isEnabled()
    apply(window)
    assert window.data_manager.working_df.iat[0, 0] == "P001"
    assert window.table.model().data(window.table.model().index(0, 0)) == "P001"
    assert window.cleaning_preview_panel.result is None
    assert not window.cleaning_preview_panel.apply_button.isEnabled()
    assert window.diagnosis_panel.result.counts["前后空格"] == 0
    assert window.diagnosis_panel.result.counts["CHAR(160)"] == 0
    assert window.diagnosis_panel.result.counts["Tab / 换行"] == 0
    assert window.diagnosis_panel.result.counts["空值"] == 2
    assert window.data_manager.working_df.iloc[10:12, 0].tolist() == ["DUP", "DUP"]
    assert len(window.data_manager.working_df) == len(original)
    assert "已修改（1 次操作）" in window.modification_label.text()
    assert window.windowTitle().endswith(" *")
    assert window.statusBar().currentMessage() == "已清洗商品编码列，共修改 6 行"
    assert window.undo_button.isEnabled() and window.restore_button.isEnabled()
    pd.testing.assert_frame_equal(window.data_manager.original_df, original)
    pd.testing.assert_frame_equal(window._preview.dataframe, original)
    assert hashlib.sha256(window._preview.file_path.read_bytes()).digest() == digest


def test_actual_confirmation_contains_scope_rules_and_cancel_default(window):
    analyze(window)
    generate(window)
    before = window.data_manager.working_df.copy(deep=True)
    def inspect_and_cancel(dialog):
        assert "工作表：问题数据" in dialog.text()
        assert "字段：商品编码" in dialog.text()
        assert "影响数据：6 行" in dialog.text()
        assert "CHAR(160)" in dialog.text()
        assert "Tab" in dialog.text()
        assert "不会覆盖原 Excel 文件" in dialog.informativeText()
        assert "可以撤销" in dialog.informativeText()
        assert dialog.defaultButton().text() == "取消"
        assert {button.text() for button in dialog.buttons()} == {"取消", "确认应用"}
        dialog.defaultButton().click()
        return 0
    with patch.object(QMessageBox, "exec", inspect_and_cancel):
        window._apply_cleaning()
    pd.testing.assert_frame_equal(window.data_manager.working_df, before)
    assert not window.data_manager.undo_stack
    assert window.cleaning_preview_panel.apply_button.isEnabled()


def test_undo_restores_column_and_disables_empty_history(window):
    digest = hashlib.sha256(window.data_manager.file_path.read_bytes()).digest()
    before = window.data_manager.working_df.copy(deep=True)
    analyze(window)
    generate(window)
    apply(window)
    undo(window)
    pd.testing.assert_frame_equal(window.data_manager.working_df, before)
    assert window.diagnosis_panel.result.counts["CHAR(160)"] == 2
    assert not window.undo_button.isEnabled()
    assert not window.restore_button.isEnabled()
    assert window.windowTitle() == "Excel 数据清洗助手"
    assert window.modification_label.text() == "当前状态：未修改"
    assert hashlib.sha256(window.data_manager.file_path.read_bytes()).digest() == digest


def test_restore_confirmation_cancellation_and_undo_restore(window):
    digest = hashlib.sha256(window.data_manager.file_path.read_bytes()).digest()
    analyze(window)
    generate(window)
    apply(window)
    cleaned = window.data_manager.working_df.copy(deep=True)
    with patch("app.ui.main_window.confirm_restore", return_value=False) as confirmation:
        window._restore_original()
    confirmation.assert_called_once_with(window, "问题数据")
    pd.testing.assert_frame_equal(window.data_manager.working_df, cleaned)
    with patch("app.ui.main_window.confirm_restore", return_value=True):
        QTest.mouseClick(window.restore_button, Qt.MouseButton.LeftButton)
    wait_for_idle(window)
    pd.testing.assert_frame_equal(window.data_manager.working_df, window.data_manager.original_df)
    assert window.undo_button.isEnabled()
    assert not window.restore_button.isEnabled()
    undo(window)
    pd.testing.assert_frame_equal(window.data_manager.working_df, cleaned)
    undo(window)
    assert not window.undo_button.isEnabled()
    assert hashlib.sha256(window.data_manager.file_path.read_bytes()).digest() == digest


def test_sheet_cache_preserves_edits_and_undo_targets_its_sheet(window):
    analyze(window)
    generate(window)
    apply(window)
    cleaned = window.data_manager.working_df.copy(deep=True)
    window.sheet_combo.setCurrentText("正常表")
    wait_for_idle(window)
    assert window.modification_label.text() == "当前状态：未修改"
    assert window.windowTitle().endswith(" *")
    window.sheet_combo.setCurrentText("问题数据")
    wait_for_idle(window)
    pd.testing.assert_frame_equal(window.data_manager.working_df, cleaned)
    assert window.table.model().data(window.table.model().index(0, 0)) == "P001"
    window.sheet_combo.setCurrentText("正常表")
    wait_for_idle(window)
    window.data_manager.file_path.unlink()  # 两个已缓存 Sheet 的操作不依赖磁盘读取。
    undo(window)
    assert window.sheet_combo.currentText() == "问题数据"
    pd.testing.assert_frame_equal(window.data_manager.working_df, window.data_manager.original_df)


def test_rule_change_clears_apply_permission(window):
    analyze(window)
    generate(window)
    window.cleaning_panel.rule_boxes["tab"].setChecked(False)
    assert not window.cleaning_preview_panel.apply_button.isEnabled()
    with patch("app.ui.main_window.confirm_apply") as confirmation:
        window._apply_cleaning()
    confirmation.assert_not_called()
    assert not window.data_manager.undo_stack


def test_same_rules_after_apply_have_no_changes_and_no_new_history(window):
    analyze(window)
    generate(window)
    apply(window)
    for box in window.cleaning_panel.rule_boxes.values():
        box.setChecked(True)
    generate(window)
    assert window.cleaning_preview_panel.result.changed_count == 0
    assert not window.cleaning_preview_panel.apply_button.isEnabled()
    with patch("app.ui.main_window.confirm_apply") as confirmation:
        window._apply_cleaning()
    confirmation.assert_not_called()
    assert len(window.data_manager.undo_stack) == 1


def test_reopen_same_file_resets_copies_history_and_state(window):
    analyze(window)
    generate(window)
    apply(window)
    path = window.data_manager.file_path
    window._start_load(str(path))
    wait_for_idle(window)
    assert not window.data_manager.undo_stack
    assert not window.undo_button.isEnabled()
    assert not window.data_manager.is_modified
    assert window.diagnosis_panel.result is None
    assert window.cleaning_preview_panel.result is None
    assert window.data_manager.working_df.iat[0, 0] == " P001 "
    assert window.windowTitle() == "Excel 数据清洗助手"


def test_new_file_clears_old_sheet_state_and_history(window, tmp_path):
    analyze(window)
    generate(window)
    apply(window)
    path = tmp_path / "新的.xlsx"
    workbook = Workbook()
    workbook.active.title = "新表"
    workbook.active.append(["新字段"])
    workbook.active.append(["新值"])
    workbook.save(path)
    workbook.close()
    window._start_load(str(path))
    wait_for_idle(window)
    assert list(window.data_manager.sheet_data) == ["新表"]
    assert not window.data_manager.undo_stack
    assert not window.data_manager.is_modified
    assert window.cleaning_preview_panel.result is None
    assert window.diagnosis_panel.result is None


def test_failed_file_open_preserves_working_data_and_undo(window, tmp_path):
    analyze(window)
    generate(window)
    apply(window)
    cleaned = window.data_manager.working_df.copy(deep=True)
    with patch.object(QMessageBox, "exec", return_value=QMessageBox.StandardButton.Ok):
        window._start_load(str(tmp_path / "不存在.xlsx"))
        wait_for_idle(window)
    pd.testing.assert_frame_equal(window.data_manager.working_df, cleaned)
    assert window.undo_button.isEnabled()
    undo(window)
    assert not window.data_manager.is_modified


def test_apply_failure_is_atomic_and_keeps_preview_available(window):
    analyze(window)
    generate(window)
    before = window.data_manager.working_df.copy(deep=True)
    with patch.object(window.data_manager, "_replace_column", side_effect=RuntimeError("测试异常")):
        with patch("app.ui.main_window.confirm_apply", return_value=True):
            with patch.object(QMessageBox, "exec", return_value=QMessageBox.StandardButton.Ok):
                window._apply_cleaning()
    pd.testing.assert_frame_equal(window.data_manager.working_df, before)
    assert not window.data_manager.undo_stack
    assert window.cleaning_preview_panel.apply_button.isEnabled()


def test_failed_recheck_never_shows_old_diagnosis(window):
    analyze(window)
    generate(window)
    with patch("app.ui.analysis_worker.analyze_column", side_effect=RuntimeError("重新检测失败")):
        with patch.object(QMessageBox, "exec", return_value=QMessageBox.StandardButton.Ok):
            apply(window)
    assert window.data_manager.working_df.iat[0, 0] == "P001"
    assert window.diagnosis_panel.result is None
    assert window.cleaning_preview_panel.result is None
    assert window.undo_button.isEnabled()
    assert "重新检测失败" in window.statusBar().currentMessage()
    window._start_analysis()
    wait_for_idle(window)
    assert window.diagnosis_panel.result.counts["CHAR(160)"] == 0


def test_selection_changed_during_confirmation_cannot_apply_old_preview(window):
    analyze(window)
    generate(window)
    def change_selection(*args):
        window._select_column(1)
        return True
    with patch("app.ui.main_window.confirm_apply", side_effect=change_selection):
        window._apply_cleaning()
    assert not window.data_manager.undo_stack
    pd.testing.assert_frame_equal(window.data_manager.working_df, window.data_manager.original_df)


def test_actions_without_open_file_or_preview_are_safe(qapp):
    window = MainWindow()
    try:
        assert not window.undo_button.isEnabled()
        assert not window.restore_button.isEnabled()
        assert not window.cleaning_preview_panel.apply_button.isEnabled()
        window._apply_cleaning()
        window._undo_last_action()
        window._restore_original()
        assert not window.data_manager.undo_stack
    finally:
        window.close()
        window.deleteLater()
        qapp.processEvents()
