"""展示单元格完整文本，支持只读选择和复制。"""

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QDialog, QDialogButtonBox, QLabel, QPlainTextEdit, QVBoxLayout


class CellPreviewDialog(QDialog):
    def __init__(self, column_name: str, row_number: int, content: str, parent=None, row_label: str = "数据行") -> None:
        super().__init__(parent)
        self.setWindowTitle("单元格完整内容")
        self.resize(800, 450)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(12)
        location = QLabel(f"{row_label}：{row_number}    字段：{column_name}")
        location.setTextFormat(Qt.TextFormat.PlainText)
        location.setWordWrap(True)
        location.setProperty("role", "sectionTitle")
        layout.addWidget(location)

        self.content_edit = QPlainTextEdit()
        self.content_edit.setReadOnly(True)
        self.content_edit.setPlainText(content)
        self.content_edit.setLineWrapMode(QPlainTextEdit.LineWrapMode.WidgetWidth)
        layout.addWidget(self.content_edit)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Close)
        buttons.button(QDialogButtonBox.StandardButton.Close).setText("关闭")
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)
