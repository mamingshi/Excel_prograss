"""清洗前后对比及统计，默认仅展示发生变化的记录。"""

import pandas as pd
from PySide6.QtCore import QSignalBlocker, Signal, Slot, Qt
from PySide6.QtWidgets import (
    QAbstractItemView,
    QCheckBox,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QPushButton,
    QTableView,
    QVBoxLayout,
)

from app.core.cleaner import CleaningPreview
from app.models.cleaning_preview_model import CleaningPreviewModel
from app.ui.cell_preview_dialog import CellPreviewDialog
from app.ui.theme import primary_button, style_table


class CleaningPreviewPanel(QGroupBox):
    apply_requested = Signal()
    COLUMNS = ("Excel行号", "原始值", "清洗后值", "发生的操作", "是否变化")

    def __init__(self, parent=None) -> None:
        super().__init__("", parent)
        self.setObjectName("cleaningPreviewPanel")
        self.result: CleaningPreview | None = None
        self._visible_records = ()
        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(8)
        summary = QHBoxLayout()
        self.column_label = QLabel("当前列：未选择")
        self.column_label.setTextFormat(Qt.TextFormat.PlainText)
        self.column_label.setWordWrap(True)
        self.total_label = QLabel("总数据：—")
        self.changed_label = QLabel("发生变化：—")
        self.unchanged_label = QLabel("未变化：—")
        for label in (self.total_label, self.changed_label, self.unchanged_label):
            label.setProperty("role", "metric")
        self.changed_label.setProperty("tone", "success")
        for label in (self.column_label, self.total_label, self.changed_label, self.unchanged_label):
            summary.addWidget(label)
        summary.addStretch()
        self.apply_button = QPushButton("应用清洗")
        primary_button(self.apply_button, "check")
        self.apply_button.setToolTip("确认后将当前预览应用到工作副本；可以撤销，不会覆盖源文件。")
        self.apply_button.clicked.connect(self.apply_requested.emit)
        summary.addWidget(self.apply_button)
        layout.addLayout(summary)
        self.processed_label = QLabel("将被处理的问题：尚未生成预览")
        self.processed_label.setWordWrap(True)
        self.processed_label.setProperty("role", "muted")
        layout.addWidget(self.processed_label)
        self.show_unchanged_box = QCheckBox("显示未变化数据")
        self.show_unchanged_box.toggled.connect(self._refresh_table)
        layout.addWidget(self.show_unchanged_box)
        self.state_label = QLabel("生成预览后才可应用；仅修改工作副本，源 Excel 保持不变。")
        self.state_label.setWordWrap(True)
        self.state_label.setProperty("role", "muted")
        layout.addWidget(self.state_label)
        self.model = CleaningPreviewModel(self)
        self.table = QTableView()
        style_table(self.table)
        self.table.setModel(self.model)
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setAlternatingRowColors(True)
        self.table.setWordWrap(False)
        self.table.setHorizontalScrollMode(QAbstractItemView.ScrollMode.ScrollPerPixel)
        self.table.setMinimumHeight(120)
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        header.sectionDoubleClicked.connect(self.table.resizeColumnToContents)
        self.table.doubleClicked.connect(self._show_cell_content)
        layout.addWidget(self.table, 1)
        self.clear_preview()
        for column, width in enumerate((75, 140, 140, 300, 80)):
            self.table.setColumnWidth(column, width)

    @Slot()
    def clear_preview(self) -> None:
        self.apply_button.setEnabled(False)
        self.result = None
        self._visible_records = ()
        self.column_label.setText("当前列：未选择")
        self.total_label.setText("总数据：—")
        self.changed_label.setText("发生变化：—")
        self.unchanged_label.setText("未变化：—")
        self.processed_label.setText("将被处理的问题：尚未生成预览")
        with QSignalBlocker(self.show_unchanged_box):
            self.show_unchanged_box.setChecked(False)
        self.show_unchanged_box.setEnabled(False)
        self.state_label.setText("生成预览后才可应用；仅修改工作副本，源 Excel 保持不变。")
        self.model.set_dataframe(pd.DataFrame(columns=self.COLUMNS))

    @Slot(object)
    def show_result(self, result: CleaningPreview) -> None:
        self.result = result
        self.column_label.setText(f"当前列：{result.column_name}")
        self.total_label.setText(f"总数据：{result.row_count}")
        self.changed_label.setText(f"发生变化：{result.changed_count}")
        self.unchanged_label.setText(f"未变化：{result.unchanged_count}")
        counts = "    ".join(f"{name}：{count}" for name, count in result.processed_counts.items())
        self.processed_label.setText("将被处理的问题（按单元格计）：" + counts)
        self.show_unchanged_box.setEnabled(True)
        message = "仅内存预览；问题计数可重叠。字符串加引号，[NBSP] / [Tab] / [LF] / [CR] 仅为显示标记；双击单元格可查看全文。"
        protected = sum(bool(record.notes) for record in result.records)
        if protected:
            message = f"有 {protected} 行因前导0、精度或日期风险保留；保留原因已在表格中显示。" + message
        if result.changed_count == 0:
            message = "未发生变化，可勾选“显示未变化数据”查看全部记录。" + message
        self.state_label.setText(message)
        self._refresh_table()

    def _refresh_table(self, *_args) -> None:
        if self.result is None:
            return
        self._visible_records = tuple(
            record for record in self.result.records
            if record.changed or record.notes or self.show_unchanged_box.isChecked()
        )
        rows = [
            (record.excel_row, record.original_value, record.cleaned_value,
             "；".join(record.operations + record.notes) or "无", "是" if record.changed else "否")
            for record in self._visible_records
        ]
        self.model.set_dataframe(pd.DataFrame(rows, columns=self.COLUMNS, dtype=object))
        self.table.scrollToTop()
        self.table.horizontalScrollBar().setValue(0)

    def _show_cell_content(self, index) -> None:
        if not index.isValid() or index.column() not in (1, 2):
            return
        record = self._visible_records[index.row()]
        dialog = CellPreviewDialog(
            self.model.column_name(index.column()), record.excel_row,
            self.model.data(index), self, row_label="Excel行号",
        )
        dialog.exec()
        dialog.deleteLater()
