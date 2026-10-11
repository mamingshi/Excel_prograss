"""工作表质量总览的只读模型，数值保留为整数以支持正确排序。"""

from PySide6.QtCore import QAbstractTableModel, QModelIndex, Qt
from PySide6.QtGui import QBrush, QColor

from app.core.sheet_analyzer import ColumnQuality, SheetAnalysis


class SheetOverviewModel(QAbstractTableModel):
    HEADERS = ("字段名", "总行数", "受影响行数", "前后空格", "CHAR(160)", "Tab/换行", "空值", "重复值", "质量分", "状态", "问题摘要", "文本数字混合", "前导0风险", "日期异常", "全角字符")

    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self.result: SheetAnalysis | None = None

    def set_result(self, result: SheetAnalysis | None) -> None:
        self.beginResetModel()
        self.result = result
        self.endResetModel()

    def rowCount(self, parent=QModelIndex()) -> int:
        return 0 if parent.isValid() or self.result is None else self.result.column_count

    def columnCount(self, parent=QModelIndex()) -> int:
        return 0 if parent.isValid() else len(self.HEADERS)

    @staticmethod
    def _values(column: ColumnQuality) -> tuple:
        return (
            column.column_name, column.total_rows, column.affected_rows,
            column.outer_spaces, column.nbsp, column.hidden_chars,
            column.empty, column.duplicates, column.quality_score,
            column.status, column.summary,
            column.numeric_mixture, column.leading_zero, column.date_anomalies, column.fullwidth,
        )

    def data(self, index, role=Qt.ItemDataRole.DisplayRole):
        if not index.isValid() or self.result is None:
            return None
        column = self.result.columns[index.row()]
        if role == Qt.ItemDataRole.DisplayRole:
            return self._values(column)[index.column()]
        if role == Qt.ItemDataRole.UserRole:
            # 排序后仍按物理列位置联动，不能用字段名定位重名列。
            return column.column_index
        if role == Qt.ItemDataRole.ToolTipRole:
            return (
                f"{column.column_name}\n前后空格：{column.outer_spaces} / CHAR(160)：{column.nbsp}"
                f"\nTab/换行：{column.hidden_chars} / 空值：{column.empty} / 重复值：{column.duplicates}"
                f"\n受影响行数：{column.affected_rows} / 质量分：{column.quality_score}"
                f"\n文本数字混合：{column.numeric_mixture} / 前导0：{column.leading_zero}"
                f"\n日期异常：{column.date_anomalies} / 全角字符：{column.fullwidth}"
                "\n检查当前九类问题；点击查看单列明细和清洗建议。"
            )
        if role == Qt.ItemDataRole.TextAlignmentRole and index.column() not in (0, 10):
            return Qt.AlignmentFlag.AlignCenter
        if role == Qt.ItemDataRole.ForegroundRole and index.column() in (8, 9):
            return QBrush(QColor("#08785d" if column.quality_score >= 85 else "#a46513"))
        return None

    def headerData(self, section, orientation, role=Qt.ItemDataRole.DisplayRole):
        if role == Qt.ItemDataRole.DisplayRole:
            if orientation == Qt.Orientation.Horizontal and 0 <= section < len(self.HEADERS):
                return self.HEADERS[section]
            if orientation == Qt.Orientation.Vertical:
                return section + 1
        if role == Qt.ItemDataRole.ToolTipRole and orientation == Qt.Orientation.Horizontal:
            if section == 8:
                return "round(100 × (1 − 受影响的唯一行数 / 总数据行数))；仅表头为 100 分。"
            if section == 7:
                return "排除空值，按原始值判断重复，数量包含首次出现。"
        return None

    def flags(self, index):
        if not index.isValid():
            return Qt.ItemFlag.NoItemFlags
        return Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable
