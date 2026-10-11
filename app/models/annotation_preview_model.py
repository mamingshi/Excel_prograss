"""辅助列预览的只读惰性模型：仅为可见行生成公式文本。"""

from PySide6.QtCore import QAbstractTableModel, QModelIndex, Qt
import pandas as pd

from app.core.annotation_builder import AnnotationPreview


class AnnotationPreviewModel(QAbstractTableModel):
    HEADERS = ("Excel行号", "原值", "辅助列结果", "公式 / 生成方式")

    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self.result: AnnotationPreview | None = None

    def set_result(self, result: AnnotationPreview | None) -> None:
        self.beginResetModel()
        self.result = result
        self.endResetModel()

    def rowCount(self, parent=QModelIndex()) -> int:
        return 0 if parent.isValid() or self.result is None else self.result.row_count

    def columnCount(self, parent=QModelIndex()) -> int:
        return 0 if parent.isValid() else len(self.HEADERS)

    def data(self, index, role=Qt.ItemDataRole.DisplayRole):
        if not index.isValid() or self.result is None:
            return None
        if role not in (Qt.ItemDataRole.DisplayRole, Qt.ItemDataRole.ToolTipRole, Qt.ItemDataRole.UserRole):
            return None
        position = index.row()
        if index.column() == 0:
            return position + 2
        if index.column() == 1:
            value = self.result.source_values.iloc[position]
            if role == Qt.ItemDataRole.UserRole:
                return value
            if isinstance(value, str):
                visible = value.replace("\u00a0", "[NBSP]").replace("\t", "[Tab]").replace("\n", "[LF]").replace("\r", "[CR]")
                return f'"{visible}"'
            return "空值" if pd.isna(value) else str(value)
        if index.column() == 2:
            return self.result.values.iloc[position]
        return self.result.formula(position)

    def headerData(self, section, orientation, role=Qt.ItemDataRole.DisplayRole):
        if role == Qt.ItemDataRole.DisplayRole:
            if orientation == Qt.Orientation.Horizontal and 0 <= section < len(self.HEADERS):
                return self.HEADERS[section]
            if orientation == Qt.Orientation.Vertical:
                return section + 1
        return None

    def flags(self, index):
        return Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable if index.isValid() else Qt.ItemFlag.NoItemFlags
