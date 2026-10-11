"""数据诊断区域：列信息、检测汇总和只读问题明细。"""

import pandas as pd
from PySide6.QtCore import Signal, Slot, Qt
from PySide6.QtWidgets import (
    QAbstractItemView,
    QGroupBox,
    QGridLayout,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QPushButton,
    QTableView,
    QVBoxLayout,
)

from app.core.detector import ColumnAnalysis, PROBLEM_TYPES
from app.models.dataframe_model import DataFrameModel
from app.ui.theme import primary_button, set_tone, style_table


class DiagnosisPanel(QGroupBox):
    analysis_requested = Signal()
    DETAIL_COLUMNS = ("Excel行号", "原始值", "问题类型", "问题说明")

    def __init__(self, parent=None) -> None:
        super().__init__("", parent)
        self.setObjectName("diagnosisPanel")
        self.result: ColumnAnalysis | None = None
        self._column_name: str | None = None
        self._busy = False
        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(8)

        controls = QHBoxLayout()
        self.column_label = QLabel("当前列：未选择")
        self.column_label.setTextFormat(Qt.TextFormat.PlainText)
        self.column_label.setWordWrap(True)
        self.row_count_label = QLabel("数据行数：0")
        self.row_count_label.setProperty("role", "muted")
        self.start_button = QPushButton("开始检测")
        primary_button(self.start_button, "scan")
        self.start_button.setEnabled(False)
        self.start_button.clicked.connect(self.analysis_requested.emit)
        controls.addWidget(self.column_label, 1)
        controls.addWidget(self.row_count_label)
        controls.addWidget(self.start_button)
        layout.addLayout(controls)

        summary = QGridLayout()
        self.count_labels = {}
        for index, problem_type in enumerate(PROBLEM_TYPES):
            label = QLabel(f"{problem_type}：— 条")
            label.setProperty("role", "metric")
            label.setProperty("tone", "neutral")
            label.setAlignment(Qt.AlignmentFlag.AlignCenter)
            self.count_labels[problem_type] = label
            summary.addWidget(label, index // 5, index % 5)
        layout.addLayout(summary)
        self.state_label = QLabel("请选择列后点击“开始检测”。")
        self.state_label.setWordWrap(True)
        self.state_label.setProperty("role", "muted")
        layout.addWidget(self.state_label)

        self.details_model = DataFrameModel(self)
        self.details_table = QTableView()
        style_table(self.details_table)
        self.details_table.setModel(self.details_model)
        self.details_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.details_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.details_table.setAlternatingRowColors(True)
        self.details_table.setWordWrap(False)
        self.details_table.setHorizontalScrollMode(QAbstractItemView.ScrollMode.ScrollPerPixel)
        self.details_table.setMinimumHeight(120)
        header = self.details_table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        header.sectionDoubleClicked.connect(self.details_table.resizeColumnToContents)
        layout.addWidget(self.details_table, 1)
        self._clear_details()
        for column, width in enumerate((85, 220, 160, 620)):
            self.details_table.setColumnWidth(column, width)

    def _clear_details(self) -> None:
        self.details_model.set_dataframe(pd.DataFrame(columns=self.DETAIL_COLUMNS))

    def set_column(self, column_name: str | None, row_count: int) -> None:
        """列或数据来源变化时清除旧诊断，避免把旧结果当作当前结果。"""
        self._column_name = column_name
        self.result = None
        self.column_label.setText(f"当前列：{column_name if column_name is not None else '未选择'}")
        self.row_count_label.setText(f"数据行数：{row_count}")
        for problem_type, label in self.count_labels.items():
            label.setText(f"{problem_type}：— 条")
            set_tone(label, "neutral")
        self.state_label.setText("请选择列后点击“开始检测”。" if column_name is None else "尚未检测；所有操作均为只读。")
        self._clear_details()
        self.set_busy(self._busy)

    def set_busy(self, busy: bool) -> None:
        self._busy = busy
        self.start_button.setEnabled(not busy and self._column_name is not None)

    @Slot(object)
    def show_result(self, result: ColumnAnalysis) -> None:
        self.result = result
        for problem_type, count in result.counts.items():
            self.count_labels[problem_type].setText(f"{problem_type}：{count} 条")
            set_tone(self.count_labels[problem_type], "warning" if count else "success")
        rows = [
            (issue.excel_row, issue.original_value, issue.problem_type, issue.description)
            for issue in result.issues
        ]
        self.details_model.set_dataframe(pd.DataFrame(rows, columns=self.DETAIL_COLUMNS, dtype=object))
        self.details_table.scrollToTop()
        self.details_table.horizontalScrollBar().setValue(0)
        if result.issues:
            affected_rows = len({issue.excel_row for issue in result.issues})
            self.state_label.setText(f"检测完成：{affected_rows} 个单元格存在问题，共 {len(result.issues)} 条明细；重复数量包含首次出现。悬停可查看完整内容。可在“展开清洗设置”中选择转换规则或生成异常检查列。")
        else:
            self.state_label.setText("检测完成：未发现本版本支持的九类问题。")
