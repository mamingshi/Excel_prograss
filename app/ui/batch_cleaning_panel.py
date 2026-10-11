"""批量工作台：独立逐列规则编辑，配置与预览分开占用完整工作区。"""

from PySide6.QtCore import QSignalBlocker, Qt, Signal
from PySide6.QtWidgets import (
    QAbstractItemView, QCheckBox, QHBoxLayout, QHeaderView, QLabel, QProgressBar,
    QPushButton, QScrollArea, QSplitter, QTableView, QTabWidget, QVBoxLayout, QWidget,
)

from app.core.batch_cleaner import BATCH_RULE_ORDER
from app.core.cleaner import rule_label
from app.models.batch_models import BatchConfigModel, BatchPreviewModel
from app.ui.cell_preview_dialog import CellPreviewDialog
from app.ui.theme import primary_button, set_tone, style_table


def _label(text=""):
    label = QLabel(text)
    label.setTextFormat(Qt.TextFormat.PlainText)
    label.setWordWrap(True)
    return label


def _table(model):
    table = QTableView()
    style_table(table)
    table.setModel(model)
    table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
    table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
    table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
    table.setAlternatingRowColors(True)
    table.setWordWrap(False)
    table.setHorizontalScrollMode(QAbstractItemView.ScrollMode.ScrollPerPixel)
    table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
    table.horizontalHeader().sectionDoubleClicked.connect(table.resizeColumnToContents)
    return table


class BatchCleaningPanel(QWidget):
    scan_requested = Signal()
    preview_requested = Signal()
    apply_requested = Signal()
    cancel_requested = Signal()
    configuration_changed = Signal()
    duplicate_requested = Signal(int)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.result = None
        self.scanned = False
        self._has_sheet = False
        self._busy = False
        self._selected_row = -1
        self._cancel_pending = False
        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 12, 12, 12)
        header = QHBoxLayout()
        self.summary_label = _label("当前工作表：未选择")
        self.summary_label.setProperty("role", "sectionTitle")
        header.addWidget(self.summary_label, 1)
        self.scan_button = QPushButton("扫描并生成建议")
        self.scan_button.clicked.connect(self.scan_requested.emit)
        header.addWidget(self.scan_button)
        self.cancel_button = QPushButton("取消任务")
        self.cancel_button.clicked.connect(self._request_cancel)
        self.cancel_button.hide()
        header.addWidget(self.cancel_button)
        layout.addLayout(header)
        self.notice_label = _label("请先打开 Excel 文件。")
        self.notice_label.setProperty("role", "badge")
        layout.addWidget(self.notice_label)
        self.progress_bar = QProgressBar()
        self.progress_bar.hide()
        layout.addWidget(self.progress_bar)
        self.pages = QTabWidget()
        layout.addWidget(self.pages, 1)

        config_page = QWidget()
        config_layout = QVBoxLayout(config_page)
        hint = _label("先选择是否处理，再点击字段修改专属规则。重复值、空值和前导0仍需人工判断。")
        hint.setProperty("role", "muted")
        config_layout.addWidget(hint)
        self.config_model = BatchConfigModel(self)
        self.config_table = _table(self.config_model)
        for index, width in enumerate((70, 150, 230, 280, 280)):
            self.config_table.setColumnWidth(index, width)
        editor = QWidget()
        editor_layout = QVBoxLayout(editor)
        self.editor_title = _label("选择字段以配置规则")
        self.editor_title.setProperty("role", "sectionTitle")
        editor_layout.addWidget(self.editor_title)
        self.enabled_box = QCheckBox("处理此字段")
        self.enabled_box.toggled.connect(self._edit_rules)
        editor_layout.addWidget(self.enabled_box)
        self.duplicate_button = QPushButton("生成重复检查列")
        self.duplicate_button.clicked.connect(self._request_duplicate)
        self.duplicate_button.hide()
        editor_layout.addWidget(self.duplicate_button)
        self.rule_boxes = {}
        for rule in BATCH_RULE_ORDER:
            box = QCheckBox(rule_label(rule))
            box.toggled.connect(self._edit_rules)
            self.rule_boxes[rule] = box
            editor_layout.addWidget(box)
        self.date_confirm_box = QCheckBox("我确认此字段是日期列")
        self.date_confirm_box.toggled.connect(self._edit_rules)
        editor_layout.addWidget(self.date_confirm_box)
        self.warning_label = _label()
        self.warning_label.setProperty("role", "muted")
        editor_layout.addWidget(self.warning_label)
        editor_layout.addStretch()
        self.editor_scroll = QScrollArea()
        self.editor_scroll.setWidgetResizable(True)
        self.editor_scroll.setMinimumWidth(270)
        self.editor_scroll.setWidget(editor)
        splitter = QSplitter(Qt.Orientation.Horizontal)
        splitter.addWidget(self.config_table)
        splitter.addWidget(self.editor_scroll)
        splitter.setChildrenCollapsible(False)
        splitter.setStretchFactor(0, 1)
        splitter.setSizes([760, 280])
        config_layout.addWidget(splitter, 1)
        actions = QHBoxLayout()
        self.configuration_label = _label()
        actions.addWidget(self.configuration_label, 1)
        self.generate_button = QPushButton("生成批量预览")
        primary_button(self.generate_button, "preview")
        self.generate_button.clicked.connect(self.preview_requested.emit)
        actions.addWidget(self.generate_button)
        config_layout.addLayout(actions)
        self.pages.addTab(config_page, "字段配置")

        preview_page = QWidget()
        preview_layout = QVBoxLayout(preview_page)
        self.stats_label = _label("尚未生成批量预览")
        self.stats_label.setProperty("role", "sectionTitle")
        preview_layout.addWidget(self.stats_label)
        self.column_stats_label = _label()
        self.column_stats_label.setMaximumHeight(52)
        preview_layout.addWidget(self.column_stats_label)
        preview_actions = QHBoxLayout()
        self.show_unchanged_box = QCheckBox("显示未变化数据")
        self.show_unchanged_box.toggled.connect(self._show_records)
        preview_actions.addWidget(self.show_unchanged_box)
        self.sample_label = _label()
        preview_actions.addWidget(self.sample_label, 1)
        self.apply_button = QPushButton("应用全部清洗")
        primary_button(self.apply_button, "check")
        self.apply_button.clicked.connect(self.apply_requested.emit)
        preview_actions.addWidget(self.apply_button)
        preview_layout.addLayout(preview_actions)
        self.preview_model = BatchPreviewModel(self)
        self.preview_table = _table(self.preview_model)
        for index, width in enumerate((115, 150, 90, 200, 200, 300, 85, 260)):
            self.preview_table.setColumnWidth(index, width)
        self.preview_table.doubleClicked.connect(self._show_cell)
        preview_layout.addWidget(self.preview_table, 1)
        footer = _label("预览仅显示临时结果。应用需确认，支持一次撤销整批；导出沿用原有两种模式。")
        footer.setProperty("role", "muted")
        preview_layout.addWidget(footer)
        self.pages.addTab(preview_page, "批量预览")
        self.config_table.selectionModel().currentRowChanged.connect(self._select_config)
        self.config_model.configuration_changed.connect(self._configuration_changed)
        self.set_busy(False)

    @property
    def configs(self):
        return self.config_model.configs

    @property
    def current_table(self):
        return self.config_table if self.pages.currentIndex() == 0 else self.preview_table

    def set_sheet(self, sheet_name, configs, *, scanned=False):
        self._has_sheet = sheet_name is not None
        self.scanned = scanned
        self.summary_label.setText(f"当前工作表：{sheet_name or '未选择'} · 字段：{len(configs)}")
        self._selected_row = -1
        self.clear_preview()
        self.config_model.set_configs(configs)
        self.pages.setCurrentIndex(0)
        self.show_notice(
            "建议已生成。请逐列核对规则；日期标准化必须确认字段含义。"
            if scanned else "建议先扫描当前工作表。点击“扫描并生成建议”，再配置各字段规则。"
        )
        if configs:
            self.config_table.selectRow(0)
        self._load_editor()
        self._update_configuration_summary()
        self.set_busy(self._busy)

    def show_notice(self, text, tone="neutral"):
        self.notice_label.setText(text)
        set_tone(self.notice_label, tone)

    def clear_preview(self):
        self.result = None
        self.preview_model.set_records(())
        self.stats_label.setText("尚未生成批量预览")
        self.column_stats_label.clear()
        self.sample_label.clear()
        self.apply_button.setEnabled(False)

    def show_result(self, result):
        self.result = result
        self.stats_label.setText(
            f"工作表：{result.sheet_name} · 参与字段：{len(result.configs)}"
            f" · 预计修改行：{result.changed_row_count} · 预计修改单元格：{result.changed_cell_count}"
        )
        text = "；".join(f"{config.column_name}：{result.column_counts[config.column_index]}"
                       for config in result.configs)
        self.column_stats_label.setText(text)
        self.column_stats_label.setToolTip(text)
        with QSignalBlocker(self.show_unchanged_box):
            self.show_unchanged_box.setChecked(False)
        self._show_records()
        self.pages.setCurrentIndex(1)
        self.show_notice(
            f"预览完成，仅展示临时结果。因安全保护保留 {result.protected_count} 个单元格；"
            "重复、空值和其他未处理问题仍需人工核对。", "success",
        )

    def set_busy(self, busy, *, task_active=False):
        self._busy = busy
        self.config_table.setEnabled(not busy and self._has_sheet)
        editable = (0 <= self._selected_row < len(self.configs)
                    and not self.configs[self._selected_row].is_generated)
        self.editor_scroll.setEnabled(not busy and editable)
        self.duplicate_button.setEnabled(not busy and editable)
        self.scan_button.setEnabled(not busy and self._has_sheet)
        self.generate_button.setEnabled(not busy and self.scanned and any(c.enabled for c in self.configs))
        self.cancel_button.setVisible(task_active)
        self.cancel_button.setEnabled(task_active and not self._cancel_pending)
        self.progress_bar.setVisible(task_active)
        if not task_active:
            self._cancel_pending = False

    def begin_task(self, message, total):
        self._cancel_pending = False
        self.progress_bar.setRange(0, max(1, total))
        self.progress_bar.setValue(0)
        self.show_notice(message)

    def show_progress(self, completed, total, column):
        self.progress_bar.setRange(0, max(1, total))
        self.progress_bar.setValue(completed)
        if not self._cancel_pending:
            self.show_notice(f"正在处理：{column} · {completed} / {total} 字段")

    def _request_cancel(self):
        self._cancel_pending = True
        self.cancel_button.setEnabled(False)
        self.show_notice("正在取消，将在当前字段完成后停止；工作副本不会修改。")
        self.cancel_requested.emit()

    def _select_config(self, current, previous):
        self._selected_row = current.row() if current.isValid() else -1
        self._load_editor()

    def _load_editor(self):
        valid = 0 <= self._selected_row < len(self.configs)
        config = self.configs[self._selected_row] if valid else None
        self.editor_title.setText(f"字段：{config.column_name}" if config else "选择字段以配置规则")
        controls = [self.enabled_box, *self.rule_boxes.values(), self.date_confirm_box]
        blockers = [QSignalBlocker(control) for control in controls]
        self.enabled_box.setChecked(bool(config and config.enabled))
        for rule, box in self.rule_boxes.items():
            box.setChecked(bool(config and rule in config.rules))
        self.date_confirm_box.setChecked(bool(config and config.date_confirmed))
        del blockers
        self.editor_scroll.setEnabled(valid and not self._busy and not config.is_generated)
        warnings = "\n".join(config.warnings) if config else ""
        self.warning_label.setText(warnings or "仅处理明确勾选的规则，不删除重复、不填充空值。")
        self.duplicate_button.setVisible(bool(config and any("重复值" in note for note in config.warnings)))
        self.duplicate_button.setEnabled(valid and not self._busy)
        self.date_confirm_box.setVisible(bool(config and "date_normalize" in config.rules))

    def _edit_rules(self, *_args):
        if self._selected_row < 0:
            return
        rules = tuple(rule for rule, box in self.rule_boxes.items() if box.isChecked())
        self.config_model.update_config(
            self._selected_row, enabled=self.enabled_box.isChecked(), rules=rules,
            date_confirmed=self.date_confirm_box.isChecked() if "date_normalize" in rules else False,
        )

    def _configuration_changed(self):
        self.clear_preview()
        self._load_editor()
        self._update_configuration_summary()
        self.set_busy(self._busy)
        self.configuration_changed.emit()

    def _update_configuration_summary(self):
        self.configuration_label.setText(f"已启用 {sum(config.enabled for config in self.configs)} / {len(self.configs)} 个字段")

    def _show_records(self, *_args):
        if self.result is None:
            return
        show_all = self.show_unchanged_box.isChecked()
        records = self.result.all_records if show_all else self.result.changed_records
        total = self.result.row_count * len(self.result.configs) if show_all else self.result.changed_cell_count
        self.preview_model.set_records(records)
        self.sample_label.setText(
            f"共有 {total} 项{'数据' if show_all else '变化'}，当前显示前 {len(records)} 项。"
            + ("应用仍处理全部变化。" if total > len(records) else "")
        )

    def _request_duplicate(self):
        if 0 <= self._selected_row < len(self.configs):
            self.duplicate_requested.emit(self.configs[self._selected_row].column_index)

    def _show_cell(self, index):
        text = self.preview_model.data(index)
        dialog = CellPreviewDialog(self.preview_model.HEADERS[index.column()], index.row() + 1, text or "", self, "预览序号")
        dialog.exec()
