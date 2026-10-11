"""真实 Qt 交互验证辅助列生成、确认应用、撤销恢复和两种导出模式。"""

import hashlib
import time
from unittest.mock import patch

from openpyxl import load_workbook
import pandas as pd
import pytest
from PySide6.QtCore import Qt
from PySide6.QtTest import QTest
from PySide6.QtWidgets import QMessageBox

from app.core.annotation_builder import build_duplicate_formula, generate_duplicate_preview
from app.core.exporter import prepare_export
from app.models.export_data import ExportMode
from app.ui.confirmation_dialogs import ExportConfirmationDialog
from app.ui.main_window import MainWindow
from tests.test_annotation_export import session
from tests.test_sheet_scan_ui import wait_for_idle


@pytest.fixture
def window(session, qapp):
    source, _ = session
    widget = MainWindow()
    widget.show()
    qapp.processEvents()
    widget._start_load(str(source))
    wait_for_idle(widget)
    yield widget
    wait_for_idle(widget)
    widget.close()
    widget.deleteLater()
    qapp.processEvents()


def analyze(window, column=3):
    window._select_column(column)
    window._start_analysis()
    wait_for_idle(window)


def generate(window):
    window._set_cleaning_sidebar_visible(True)
    window.cleaning_scroll.ensureWidgetVisible(window.cleaning_panel.annotation_button)
    QTest.qWait(20)
    QTest.mouseClick(window.cleaning_panel.annotation_button, Qt.MouseButton.LeftButton)
    assert window._annotation_worker is not None
    wait_for_idle(window)


def apply(window):
    with patch("app.ui.main_window.confirm_annotation", return_value=True) as confirmation:
        QTest.mouseClick(window.annotation_preview_panel.apply_button, Qt.MouseButton.LeftButton)
    confirmation.assert_called_once()
    wait_for_idle(window)


def test_preview_shows_values_labels_formulas_without_modifying_data(window):
    before = window.data_manager.working_df.copy(deep=True)
    digest = hashlib.sha256(window.data_manager.file_path.read_bytes()).digest()
    assert not window.cleaning_panel.annotation_button.isEnabled()
    analyze(window)
    assert window.cleaning_panel.annotation_button.isEnabled()
    generate(window)
    panel = window.annotation_preview_panel
    assert window.results_tabs.currentWidget() is panel
    assert panel.model.rowCount() == 8
    assert panel.model.data(panel.model.index(0, 0)) == 2
    assert panel.model.data(panel.model.index(0, 1)) == '"P001"'
    assert panel.model.data(panel.model.index(0, 2)) == "重复"
    assert panel.model.data(panel.model.index(3, 2)) == "空值"
    assert panel.model.data(panel.model.index(0, 3)) == build_duplicate_formula(3, 2, 9)
    assert not panel.model.flags(panel.model.index(0, 0)) & Qt.ItemFlag.ItemIsEditable
    assert panel.apply_button.isEnabled()
    assert not window.export_button.isEnabled()
    assert not window.data_manager.undo_stack
    pd.testing.assert_frame_equal(window.data_manager.working_df, before)
    assert hashlib.sha256(window.data_manager.file_path.read_bytes()).digest() == digest


def test_apply_adds_rightmost_column_and_undo_removes_it(window):
    analyze(window)
    generate(window)
    apply(window)
    assert window.results_tabs.currentWidget() is window.raw_data_page
    assert window.data_manager.working_df.columns[-1] == "商品编码_重复检查"
    assert window.table.model().columnCount() == 5
    assert window.selected_column == 4
    assert window.export_button.isEnabled()
    assert window.undo_button.isEnabled() and window.restore_button.isEnabled()
    assert window.annotation_preview_panel.result is None
    assert not window.annotation_preview_panel.apply_button.isEnabled()
    window._undo_last_action()
    wait_for_idle(window)
    assert window.table.model().columnCount() == 4
    assert not window.data_manager.is_modified
    assert not window.export_button.isEnabled()
    assert window.selected_column == 3
    assert window.diagnosis_panel.result.counts["重复值"] == 2


def test_confirmation_is_explicit_cancel_by_default_and_does_not_delete_records(window):
    analyze(window)
    generate(window)
    before = window.data_manager.working_df.copy(deep=True)

    def cancel(dialog):
        assert "新增辅助列：商品编码_重复检查" in dialog.text()
        assert "数据区域最右侧" in dialog.text()
        assert "重复记录：2" in dialog.text()
        assert "不删除、合并或修改原有记录" in dialog.informativeText()
        assert dialog.defaultButton().text() == "取消"
        assert {button.text() for button in dialog.buttons()} == {"取消", "确认应用辅助列"}
        dialog.defaultButton().click()
        return 0

    with patch.object(QMessageBox, "exec", cancel):
        window._apply_annotation()
    pd.testing.assert_frame_equal(window.data_manager.working_df, before)
    assert not window.data_manager.undo_stack
    assert window.annotation_preview_panel.apply_button.isEnabled()


def test_selection_or_sheet_change_clears_old_annotation_preview(window):
    analyze(window)
    generate(window)
    window._select_column(1)
    assert window.annotation_preview_panel.result is None
    assert not window.annotation_preview_panel.apply_button.isEnabled()
    analyze(window)
    generate(window)
    window.sheet_combo.setCurrentText("统计")
    wait_for_idle(window)
    assert window.annotation_preview_panel.result is None


def test_selection_change_during_confirmation_cannot_apply_old_preview(window):
    analyze(window)
    generate(window)

    def change_selection(*args):
        window._select_column(1)
        return True

    with patch("app.ui.main_window.confirm_annotation", side_effect=change_selection):
        window._apply_annotation()
    assert window.data_manager.working_df.shape[1] == 4
    assert not window.data_manager.undo_stack


def test_rule_changes_do_not_invalidate_independent_annotation_preview(window):
    analyze(window)
    generate(window)
    preview = window._prepared_annotation
    window.cleaning_panel.rule_boxes["outer_spaces"].setChecked(True)
    assert window._prepared_annotation is preview
    assert window.annotation_preview_panel.apply_button.isEnabled()


def test_normal_column_and_generated_column_do_not_enable_annotation(window):
    analyze(window, 0)
    assert not window.cleaning_panel.annotation_button.isEnabled()
    window._start_annotation_preview()
    assert window._annotation_worker is None
    analyze(window)
    generate(window)
    apply(window)
    window._start_analysis()
    wait_for_idle(window)
    assert not window.cleaning_panel.annotation_button.isEnabled()
    assert not window.cleaning_panel.generate_button.isEnabled()
    window._start_annotation_preview()
    window._start_cleaning_preview()
    assert window._annotation_worker is None and window._cleaning_worker is None


def test_restore_removes_all_helpers_and_undo_restore_returns_them(window):
    analyze(window)
    generate(window)
    apply(window)
    analyze(window)
    generate(window)
    apply(window)
    assert window.data_manager.working_df.columns[-1] == "商品编码_重复检查_2"
    before = window.data_manager.working_df.copy(deep=True)
    with patch("app.ui.main_window.confirm_restore", return_value=True):
        window._restore_original()
    wait_for_idle(window)
    assert window.data_manager.working_df.shape[1] == 4
    assert not window.data_manager.current_data.annotations
    window._undo_last_action()
    wait_for_idle(window)
    pd.testing.assert_frame_equal(window.data_manager.working_df, before)
    assert len(window.data_manager.current_data.annotations) == 2


@pytest.mark.parametrize("mode", list(ExportMode))
def test_annotation_only_can_export_through_existing_ui(window, tmp_path, mode):
    analyze(window)
    generate(window)
    apply(window)
    path = tmp_path / f"界面导出_{mode}.xlsx"
    digest = hashlib.sha256(window.data_manager.file_path.read_bytes()).digest()
    with (patch("app.ui.main_window.QFileDialog.getSaveFileName", return_value=(str(path), "")),
          patch("app.ui.main_window.confirm_export", return_value=mode),
          patch("app.ui.main_window.show_export_success") as success):
        window._export_results()
        wait_for_idle(window)
    success.assert_called_once()
    assert success.call_args.args[1].annotation_column_count == 1
    book = load_workbook(path)
    assert book["订单"]["E2"].value == build_duplicate_formula(3, 2, 9)
    book.close()
    assert hashlib.sha256(window.data_manager.file_path.read_bytes()).digest() == digest
    assert window.undo_button.isEnabled()


def test_export_confirmation_describes_helpers_for_both_modes(window, tmp_path):
    analyze(window)
    generate(window)
    apply(window)
    dialog = ExportConfirmationDialog(window, prepare_export(window.data_manager.get_export_snapshot()), tmp_path / "输出.xlsx")
    try:
        assert "新增辅助列：1" in dialog.summary_label.text()
        assert "最右侧追加" in dialog.mode_note_label.text()
        dialog.update_radio.setChecked(True)
        assert "最右侧追加" in dialog.mode_note_label.text()
        assert "写入公式" in dialog.mode_note_label.text()
    finally:
        dialog.close()
        dialog.deleteLater()


def test_annotation_worker_error_preserves_data_and_recovers_controls(window):
    analyze(window)
    before = window.data_manager.working_df.copy(deep=True)
    with patch("app.ui.annotation_worker.generate_duplicate_preview", side_effect=RuntimeError("测试错误")):
        with patch.object(QMessageBox, "exec", return_value=QMessageBox.StandardButton.Ok):
            generate(window)
    assert window.annotation_preview_panel.result is None
    assert not window.annotation_preview_panel.apply_button.isEnabled()
    assert window.cleaning_panel.annotation_button.isEnabled()
    pd.testing.assert_frame_equal(window.data_manager.working_df, before)


def test_close_during_annotation_preview_waits_for_thread(window):
    analyze(window)

    def slow_preview(*args):
        time.sleep(0.05)
        return generate_duplicate_preview(*args)

    with patch("app.ui.annotation_worker.generate_duplicate_preview", side_effect=slow_preview):
        window._start_annotation_preview()
        window.close()
        assert window._close_requested
        wait_for_idle(window)
    assert not window.isVisible()
