"""真实 Qt 操作验证批量工作台、后台任务、确认、一次撤销及旧入口联动。"""

import hashlib
import time
from unittest.mock import patch

from openpyxl import Workbook
import pandas as pd
import pytest
from PySide6.QtCore import QPoint, Qt
from PySide6.QtTest import QTest
from PySide6.QtWidgets import QDialog, QLabel, QPushButton, QStyle, QStyleOptionViewItem

from app.core.file_loader import ExcelPreview
from app.core.sheet_analyzer import ScanCancelled
from app.ui.main_window import MainWindow
from tests.test_sheet_scan_ui import wait_for_idle, wait_until


@pytest.fixture
def window(tmp_path, qapp):
    path = tmp_path / "批量工作台.xlsx"
    book = Workbook()
    sheet = book.active
    sheet.title = "订单"
    sheet.append(["商品编码", "金额", "订单日期", "备注"])
    rows = [
        [" Ｐ001 ", "1002", "2026/10/11", "备注\n"],
        ["P002\u00a0", 2, "2026-10-12", "普通备注"],
        ["P003", 0, "2026年10月13日", "普通备注"],
        ["P003", 12.5, "abc", "普通备注"],
        ["00123", None, "2026-02-30", "普通备注"],
    ]
    for row in rows:
        sheet.append(row)
    book.create_sheet("其他").append(["名称"])
    book.save(path)
    book.close()
    widget = MainWindow()
    widget.show()
    widget._start_load(str(path))
    wait_for_idle(widget)
    yield widget
    wait_for_idle(widget)
    widget.close()
    widget.deleteLater()
    qapp.processEvents()


def scan(window):
    window._open_batch_workspace()
    QTest.mouseClick(window.batch_panel.scan_button, Qt.MouseButton.LeftButton)
    wait_for_idle(window)


def configure(window):
    panel = window.batch_panel
    panel.config_model.update_config(3, enabled=False)
    panel.config_table.selectRow(2)
    QTest.qWait(5)
    assert panel.date_confirm_box.isVisible()
    QTest.mouseClick(panel.date_confirm_box, Qt.MouseButton.LeftButton)
    assert panel.configs[2].date_confirmed


def generate(window):
    QTest.mouseClick(window.batch_panel.generate_button, Qt.MouseButton.LeftButton)
    wait_for_idle(window)


def ready(window):
    scan(window)
    configure(window)
    generate(window)
    assert window.batch_panel.result is not None


def test_scan_populates_independent_config_and_reuses_valid_cache(window):
    panel = window.batch_panel
    assert not panel.scanned and not panel.generate_button.isEnabled()
    assert len(panel.configs) == 4
    window._start_sheet_scan()
    wait_for_idle(window)
    cache = window.data_manager.current_data.scan_cache
    QTest.mouseClick(window.sheet_overview_panel.batch_button, Qt.MouseButton.LeftButton)
    assert window.results_tabs.currentWidget() is panel
    assert panel.scanned and panel.generate_button.isEnabled()
    assert window._scan_worker is None
    assert window.data_manager.current_data.scan_cache is cache
    assert "numeric_conversion" not in panel.configs[0].rules
    assert panel.configs[1].rules == ("numeric_conversion",)
    assert panel.configs[2].rules == ("date_normalize",)
    assert not panel.configs[2].date_confirmed
    panel.config_table.selectRow(0)
    assert "前导0" in panel.warning_label.text()
    assert "建议生成辅助列" in panel.warning_label.text()
    assert panel.duplicate_button.isVisible() and panel.duplicate_button.isEnabled()


def test_configuration_checkbox_and_rule_editor_update_only_current_column(window):
    scan(window)
    panel = window.batch_panel
    before_other = panel.configs[1]
    index = panel.config_model.index(0, 0)
    option = QStyleOptionViewItem()
    option.initFrom(panel.config_table)
    panel.config_table.itemDelegate().initStyleOption(option, index)
    option.rect = panel.config_table.visualRect(index)
    indicator = panel.config_table.style().subElementRect(
        QStyle.SubElement.SE_ItemViewItemCheckIndicator, option, panel.config_table,
    )
    QTest.mouseClick(panel.config_table.viewport(), Qt.MouseButton.LeftButton, pos=indicator.center())
    assert not panel.configs[0].enabled
    assert panel.configs[1] == before_other
    QTest.mouseClick(panel.enabled_box, Qt.MouseButton.LeftButton, pos=QPoint(8, panel.enabled_box.height() // 2))
    assert panel.configs[0].enabled
    original_rules = panel.configs[0].rules
    box = panel.rule_boxes["fullwidth"]
    QTest.mouseClick(box, Qt.MouseButton.LeftButton, pos=QPoint(8, box.height() // 2))
    assert "fullwidth" not in panel.configs[0].rules and "fullwidth" in original_rules
    assert panel.configs[1] == before_other


def test_preview_apply_once_undo_and_source_protection(window):
    digest = hashlib.sha256(window.data_manager.file_path.read_bytes()).digest()
    original = window.data_manager.working_df.copy(deep=True)
    ready(window)
    panel = window.batch_panel
    assert window.results_tabs.currentWidget() is panel
    assert panel.result.changed_cell_count == 5
    assert panel.result.changed_row_count == 3
    assert panel.preview_model.rowCount() == 5
    assert panel.preview_model.columnCount() == 8
    assert not panel.preview_model.flags(panel.preview_model.index(0, 0)) & Qt.ItemFlag.ItemIsEditable
    assert panel.preview_model.data(panel.preview_model.index(2, 3)) == "'1002'"
    assert panel.preview_model.data(panel.preview_model.index(2, 4)) == "1002"
    assert panel.apply_button.isEnabled()
    pd.testing.assert_frame_equal(window.data_manager.working_df, original)
    QTest.mouseClick(panel.show_unchanged_box, Qt.MouseButton.LeftButton)
    assert panel.preview_model.rowCount() == 15
    assert panel.preview_model.data(panel.preview_model.index(13, 7)) == "无法识别日期，保留原值"
    with patch("app.ui.main_window.confirm_batch_apply", return_value=True) as confirm:
        QTest.mouseClick(panel.apply_button, Qt.MouseButton.LeftButton)
    confirm.assert_called_once()
    wait_for_idle(window)
    assert window.data_manager.working_df.iat[0, 0] == "P001"
    assert window.data_manager.working_df.iat[0, 1] == 1002
    assert window.data_manager.working_df.iat[4, 0] == "00123"
    assert window.data_manager.working_df.iat[3, 2] == "abc"
    assert window.data_manager.working_df.iat[0, 3] == "备注\n"
    assert window.results_tabs.currentWidget() is window.raw_data_page
    assert panel.result is None and window._prepared_batch is None
    assert len(window.data_manager.undo_stack) == 1
    assert window.undo_button.isEnabled() and window.export_button.isEnabled()
    assert "修改 5 个单元格" in window.statusBar().currentMessage()
    assert "待更新" in window.results_tabs.tabText(window.results_tabs.indexOf(window.sheet_overview_panel))
    assert not panel.scanned
    QTest.mouseClick(window.undo_button, Qt.MouseButton.LeftButton)
    wait_for_idle(window)
    pd.testing.assert_frame_equal(window.data_manager.working_df, original)
    pd.testing.assert_frame_equal(window.data_manager.original_df, original)
    assert not window.data_manager.undo_stack
    assert not window.export_button.isEnabled()
    assert "整批清洗" in window.statusBar().currentMessage()
    assert hashlib.sha256(window.data_manager.file_path.read_bytes()).digest() == digest


def test_date_confirmation_is_required_and_leading_zero_rule_is_blocked(window):
    scan(window)
    before = window.data_manager.working_df.copy(deep=True)
    with patch.object(window, "_show_error_dialog") as error:
        generate(window)
    assert "确认" in error.call_args.args[-1]
    assert window.batch_panel.result is None
    configure(window)
    window.batch_panel.config_model.update_config(
        0, rules=window.batch_panel.configs[0].rules + ("numeric_conversion",),
    )
    with patch.object(window, "_show_error_dialog") as error:
        generate(window)
    assert "前导0" in error.call_args.args[-1]
    assert not window.batch_panel.apply_button.isEnabled()
    pd.testing.assert_frame_equal(window.data_manager.working_df, before)
    assert not window.data_manager.undo_stack


def test_actual_confirmation_cancel_is_default_and_changes_nothing(window):
    ready(window)
    before = window.data_manager.working_df.copy(deep=True)
    def inspect(dialog):
        assert dialog.windowTitle() == "确认应用批量清洗"
        text = "\n".join(label.text() for label in dialog.findChildren(QLabel))
        assert "工作表：订单" in text and "字段：3" in text
        assert "修改单元格：5" in text and "修改数据行：3" in text
        assert "不会覆盖原 Excel 文件" in text
        assert [button.text() for button in dialog.findChildren(QPushButton) if button.isDefault()] == ["取消"]
        return QDialog.DialogCode.Rejected
    with patch.object(QDialog, "exec", inspect):
        window._apply_batch_cleaning()
    assert window.batch_panel.result is not None
    pd.testing.assert_frame_equal(window.data_manager.working_df, before)
    assert not window.data_manager.undo_stack


def test_changing_rules_invalidates_batch_preview_and_confirmation_is_rechecked(window):
    ready(window)
    panel = window.batch_panel
    before = window.data_manager.working_df.copy(deep=True)
    def change_during_confirmation(*args):
        panel.config_model.update_config(1, enabled=False)
        return True
    with patch("app.ui.main_window.confirm_batch_apply", side_effect=change_during_confirmation):
        window._apply_batch_cleaning()
    assert panel.result is None and window._prepared_batch is None
    assert not panel.apply_button.isEnabled()
    pd.testing.assert_frame_equal(window.data_manager.working_df, before)
    assert not window.data_manager.undo_stack


def test_cancel_background_preview_leaves_ui_responsive_and_data_unchanged(window):
    scan(window)
    configure(window)
    started = []
    def slow_preview(*args, is_cancelled, **kwargs):
        started.append(True)
        deadline = time.monotonic() + 5
        while not is_cancelled() and time.monotonic() < deadline:
            time.sleep(0.01)
        raise ScanCancelled()
    before = window.data_manager.working_df.copy(deep=True)
    with patch("app.ui.batch_worker.generate_batch_preview", side_effect=slow_preview):
        QTest.mouseClick(window.batch_panel.generate_button, Qt.MouseButton.LeftButton)
        wait_until(lambda: bool(started))
        assert window.batch_panel.cancel_button.isVisible()
        assert not window.open_button.isEnabled()
        assert not window.undo_button.isEnabled()
        QTest.mouseClick(window.batch_panel.cancel_button, Qt.MouseButton.LeftButton)
        wait_for_idle(window)
    assert window.batch_panel.result is None
    assert "已取消" in window.batch_panel.notice_label.text()
    pd.testing.assert_frame_equal(window.data_manager.working_df, before)


def test_duplicate_helper_button_uses_existing_preview_without_modification(window):
    scan(window)
    panel = window.batch_panel
    panel.config_table.selectRow(0)
    QTest.mouseClick(panel.duplicate_button, Qt.MouseButton.LeftButton)
    wait_for_idle(window)
    assert window.results_tabs.currentWidget() is window.annotation_preview_panel
    assert window.annotation_preview_panel.result.duplicate_count == 2
    assert window.data_manager.working_df.shape[1] == 4
    assert not window.data_manager.undo_stack


def test_large_preview_limit_is_display_only_and_sheet_switch_discards_preview(window):
    frame = pd.DataFrame({"编码": [" X "] * 1205})
    preview = ExcelPreview(window.data_manager.file_path, ("大表",), "大表", frame)
    window.data_manager.open_workbook(preview)
    window._display_sheet(preview)
    scan(window)
    generate(window)
    panel = window.batch_panel
    assert panel.preview_model.rowCount() == 1000
    assert "1205" in panel.sample_label.text() and "1000" in panel.sample_label.text()
    assert "应用仍处理全部变化" in panel.sample_label.text()
    with patch("app.ui.main_window.confirm_batch_apply", return_value=True):
        window._apply_batch_cleaning()
    assert window.data_manager.working_df.iat[1204, 0] == "X"
    assert len(window.data_manager.undo_stack) == 1
    window._undo_last_action()
    assert window.data_manager.working_df.iat[1204, 0] == " X "
