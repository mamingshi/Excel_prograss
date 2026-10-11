"""批量配置和限量预览的 Qt 模型；不创建逐单元格控件。"""

from dataclasses import replace

from PySide6.QtCore import QAbstractTableModel, QModelIndex, Qt, Signal

from app.core.batch_cleaner import rules_description
from app.core.type_utils import is_empty_value


class BatchConfigModel(QAbstractTableModel):
    configuration_changed = Signal()
    HEADERS = ("是否处理", "字段", "检测问题", "清洗规则", "需要人工判断")

    def __init__(self, parent=None):
        super().__init__(parent)
        self.configs = ()

    def set_configs(self, configs):
        self.beginResetModel()
        self.configs = tuple(configs)
        self.endResetModel()

    def update_config(self, row, **fields):
        if not 0 <= row < len(self.configs) or self.configs[row].is_generated:
            return
        config = replace(self.configs[row], **fields)
        if config == self.configs[row]:
            return
        values = list(self.configs)
        values[row] = config
        self.configs = tuple(values)
        self.dataChanged.emit(self.index(row, 0), self.index(row, len(self.HEADERS) - 1))
        self.configuration_changed.emit()

    def rowCount(self, parent=QModelIndex()):
        return 0 if parent.isValid() else len(self.configs)

    def columnCount(self, parent=QModelIndex()):
        return 0 if parent.isValid() else len(self.HEADERS)

    def headerData(self, section, orientation, role=Qt.ItemDataRole.DisplayRole):
        if role == Qt.ItemDataRole.DisplayRole:
            return self.HEADERS[section] if orientation == Qt.Orientation.Horizontal else str(section + 1)
        return None

    def data(self, index, role=Qt.ItemDataRole.DisplayRole):
        if not index.isValid():
            return None
        config = self.configs[index.row()]
        if role == Qt.ItemDataRole.CheckStateRole and index.column() == 0:
            return Qt.CheckState.Checked if config.enabled else Qt.CheckState.Unchecked
        if role in (Qt.ItemDataRole.DisplayRole, Qt.ItemDataRole.ToolTipRole):
            return ("", config.column_name, config.detected_problems, rules_description(config),
                    "；".join(config.warnings) or "—")[index.column()]
        return None

    def flags(self, index):
        if not index.isValid():
            return Qt.ItemFlag.NoItemFlags
        flags = Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable
        if index.column() == 0 and not self.configs[index.row()].is_generated:
            flags |= Qt.ItemFlag.ItemIsUserCheckable
        return flags

    def setData(self, index, value, role=Qt.ItemDataRole.EditRole):
        if (not index.isValid() or index.column() != 0 or role != Qt.ItemDataRole.CheckStateRole
                or self.configs[index.row()].is_generated):
            return False
        self.update_config(index.row(), enabled=value in (Qt.CheckState.Checked, Qt.CheckState.Checked.value))
        return True


class BatchPreviewModel(QAbstractTableModel):
    HEADERS = ("工作表", "字段", "Excel行号", "原始值", "清洗后值", "执行规则", "是否变化", "保留说明")

    def __init__(self, parent=None):
        super().__init__(parent)
        self.records = ()

    def set_records(self, records):
        self.beginResetModel()
        self.records = tuple(records)
        self.endResetModel()

    def rowCount(self, parent=QModelIndex()):
        return 0 if parent.isValid() else len(self.records)

    def columnCount(self, parent=QModelIndex()):
        return 0 if parent.isValid() else len(self.HEADERS)

    def headerData(self, section, orientation, role=Qt.ItemDataRole.DisplayRole):
        if role == Qt.ItemDataRole.DisplayRole:
            return self.HEADERS[section] if orientation == Qt.Orientation.Horizontal else str(section + 1)
        return None

    def data(self, index, role=Qt.ItemDataRole.DisplayRole):
        if not index.isValid() or role not in (Qt.ItemDataRole.DisplayRole, Qt.ItemDataRole.ToolTipRole):
            return None
        record = self.records[index.row()]
        def display(value):
            if is_empty_value(value):
                return "（空值）"
            # 引号及转义让文本数字、首尾空格、Tab 等在预览中可区分。
            return repr(value) if isinstance(value, str) else str(value)
        return (record.sheet_name, record.column_name, str(record.excel_row),
                display(record.original_value), display(record.cleaned_value), record.rules,
                "是" if record.changed else "否", "；".join(record.notes))[index.column()]

    def flags(self, index):
        return (Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable
                if index.isValid() else Qt.ItemFlag.NoItemFlags)
