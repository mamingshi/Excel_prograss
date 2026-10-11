"""异常标注预览：逐行展示原值、标签和公式，确认后才加入工作副本。"""

from PySide6.QtCore import Qt, Signal, Slot
from PySide6.QtWidgets import (
    QAbstractItemView, QHBoxLayout, QHeaderView, QLabel,
    QPushButton, QTableView, QVBoxLayout, QWidget,
)

from app.core.annotation_builder import AnnotationPreview
from app.models.annotation_preview_model import AnnotationPreviewModel
from app.ui.cell_preview_dialog import CellPreviewDialog
from app.ui.theme import primary_button, style_table


class AnnotationPreviewPanel(QWidget):
    apply_requested = Signal()

    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self.result: AnnotationPreview | None = None
        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 12, 12, 12)
        controls = QHBoxLayout()
        self.column_label = QLabel()
        self.column_label.setTextFormat(Qt.TextFormat.PlainText)
        self.column_label.setWordWrap(True)
        self.column_label.setProperty("role", "sectionTitle")
        controls.addWidget(self.column_label, 1)
        self.apply_button = QPushButton("应用辅助列")
        primary_button(self.apply_button, "check")
        self.apply_button.clicked.connect(self.apply_requested.emit)
        controls.addWidget(self.apply_button)
        layout.addLayout(controls)
        self.summary_label = QLabel()
        self.summary_label.setProperty("role", "badge")
        self.summary_label.setProperty("tone", "success")
        self.summary_label.setWordWrap(True)
        layout.addWidget(self.summary_label)
        self.note_label = QLabel(
            "只添加标注列，不删除重复记录。应用到数据区域最右侧，可撤销；导出时写入旧版 Excel 兼容公式。"
            "\n精确区分大小写和数字/文本；空值单独标记。公式限定到当前数据行，不计表头。双击单元格可查看全文。"
        )
        self.note_label.setWordWrap(True)
        self.note_label.setProperty("role", "muted")
        layout.addWidget(self.note_label)
        self.model = AnnotationPreviewModel(self)
        self.table = QTableView()
        style_table(self.table)
        self.table.setModel(self.model)
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setAlternatingRowColors(True)
        self.table.setWordWrap(False)
        self.table.setHorizontalScrollMode(QAbstractItemView.ScrollMode.ScrollPerPixel)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        self.table.horizontalHeader().sectionDoubleClicked.connect(self.table.resizeColumnToContents)
        self.table.doubleClicked.connect(self._show_content)
        layout.addWidget(self.table, 1)
        for column, width in enumerate((85, 230, 145, 800)):
            self.table.setColumnWidth(column, width)
        self.clear_preview()

    @Slot()
    def clear_preview(self) -> None:
        self.result = None
        self.apply_button.setEnabled(False)
        self.model.set_result(None)
        self.column_label.setText("异常标注 / 辅助列")
        self.summary_label.setText("选择原始字段并完成检测，在清洗设置中选择需要生成的检查列。")

    def show_result(self, result: AnnotationPreview) -> None:
        self.result = result
        self.model.set_result(result)
        self.column_label.setText(f"{result.source_name} → {result.annotation.column_name}")
        is_duplicate = result.annotation.kind == "duplicate"
        self.note_label.setText(
            "只添加标注列，不删除或修改原有记录。应用到数据区域最右侧，可撤销。"
            + ("导出时写入精确匹配公式，区分大小写和数字/文本。" if is_duplicate
               else "本列导出为检测标签；在 Excel 中修改原值后需要重新检测。")
            + "\n空值单独标记；双击单元格可查看全文。"
        )
        self.summary_label.setText(
            f"预览 · 数据行数：{result.row_count} · 重复：{result.duplicate_count}"
            f" · 空值：{result.empty_count} · 正常：{result.row_count - result.duplicate_count - result.empty_count}"
            " · 尚未添加到工作副本"
        )
        if not is_duplicate:
            self.summary_label.setText(
                f"预览 · 数据行数：{result.row_count} · 需关注：{result.problem_count}"
                f" · 空值：{result.empty_count} · 尚未添加到工作副本"
            )
        self.table.scrollToTop()

    def _show_content(self, index) -> None:
        if index.isValid():
            dialog = CellPreviewDialog(
                self.model.headerData(index.column(), Qt.Orientation.Horizontal),
                index.row() + 2, str(self.model.data(index)), self, row_label="Excel行号",
            )
            dialog.exec()
            dialog.deleteLater()
