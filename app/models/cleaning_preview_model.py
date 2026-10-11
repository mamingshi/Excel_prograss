"""仅在显示层标记不可见字符，底层原始值及其类型保持不变。"""

import pandas as pd
from PySide6.QtCore import Qt

from app.models.dataframe_model import DataFrameModel


class CleaningPreviewModel(DataFrameModel):
    def data(self, index, role=Qt.ItemDataRole.DisplayRole):
        if not index.isValid():
            return None
        if role == Qt.ItemDataRole.UserRole:
            return self._dataframe.iat[index.row(), index.column()]
        if index.column() in (1, 2) and role in (Qt.ItemDataRole.DisplayRole, Qt.ItemDataRole.ToolTipRole):
            value = self._dataframe.iat[index.row(), index.column()]
            if isinstance(value, str):
                visible = value.replace("\u00a0", "[NBSP]").replace("\t", "[Tab]").replace("\n", "[LF]").replace("\r", "[CR]")
                return f'"{visible}"'
            if value is None:
                return "None"
            if value is pd.NA:
                return "<NA>"
            return "NaN" if pd.isna(value) else str(value)
        return super().data(index, role)
