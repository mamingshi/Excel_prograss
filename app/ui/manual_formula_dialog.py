"""独立公式工具：调整引用和规则后复制，不执行清洗或导出。"""

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QApplication, QCheckBox, QComboBox, QDialog, QGridLayout, QGroupBox,
    QHBoxLayout, QLabel, QLineEdit, QPlainTextEdit, QPushButton, QVBoxLayout,
)

from app.core.cleaner import ALL_RULE_ORDER, RULE_ORDER, TYPE_RULE_ORDER, CleaningRules, rule_label
from app.core.manual_formula_builder import build_manual_formula
from app.ui.theme import primary_button


def _label(text=""):
    label = QLabel(text)
    label.setTextFormat(Qt.TextFormat.PlainText)
    label.setWordWrap(True)
    return label


class ManualFormulaDialog(QDialog):
    """规则只保存在窗口中，不会改变主窗口配置、工作副本或撤销历史。"""

    def __init__(
        self, parent=None, *, sheet_name: str = "", column_name: str = "",
        cell_reference: str = "A2", selected_rules: tuple[str, ...] = (),
        hidden_strategy: str = "space",
    ) -> None:
        super().__init__(parent)
        self.setWindowTitle("查看 / 复制清洗公式")
        self.resize(760, 680)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(18, 16, 18, 16)
        layout.setSpacing(10)
        context = f"工作表：{sheet_name} · 初始字段：{column_name}" if column_name else (
            f"工作表：{sheet_name} · 请手动填写源单元格" if sheet_name else "独立公式工具 · 无需打开或导出文件"
        )
        self.context_label = _label(context)
        self.context_label.setProperty("role", "sectionTitle")
        layout.addWidget(self.context_label)

        source_row = QHBoxLayout()
        source_row.addWidget(_label("源单元格："))
        self.reference_edit = QLineEdit(cell_reference)
        self.reference_edit.setPlaceholderText("例如 D2 或 $D2")
        self.reference_edit.setToolTip("引用同一工作表的源数据；默认第1行为表头，第2行为首条数据。")
        source_row.addWidget(self.reference_edit, 1)
        source_row.addWidget(_label("在新列的对应行粘贴"))
        layout.addLayout(source_row)

        group = QGroupBox("选择公式规则（仅影响本窗口）")
        rules_layout = QGridLayout(group)
        self.rule_boxes = {}
        for index, rule in enumerate(ALL_RULE_ORDER):
            box = QCheckBox(rule_label(rule))
            box.setChecked(rule in selected_rules)
            if rule in TYPE_RULE_ORDER:
                box.setText(rule_label(rule) + "（暂不支持公式）")
            self.rule_boxes[rule] = box
            rules_layout.addWidget(box, index // 2, index % 2)
        layout.addWidget(group)

        strategy_row = QHBoxLayout()
        strategy_row.addWidget(_label("Tab / 换行 / 回车："))
        self.strategy_combo = QComboBox()
        self.strategy_combo.addItem("替换为空格", "space")
        self.strategy_combo.addItem("直接删除", "delete")
        self.strategy_combo.setCurrentIndex(1 if hidden_strategy == "delete" else 0)
        strategy_row.addWidget(self.strategy_combo, 1)
        layout.addLayout(strategy_row)

        mode_row = QHBoxLayout()
        mode_row.addWidget(_label("公式模式："))
        self.mode_combo = QComboBox()
        self.mode_combo.addItem("简洁版（默认，常用）", "simple")
        self.mode_combo.addItem("严格版（保留中间连续空格）", "strict")
        mode_row.addWidget(self.mode_combo, 1)
        layout.addLayout(mode_row)
        self.mode_note_label = _label()
        self.mode_note_label.setProperty("role", "muted")
        layout.addWidget(self.mode_note_label)

        self.formula_edit = QPlainTextEdit()
        self.formula_edit.setObjectName("formulaEditor")
        self.formula_edit.setReadOnly(True)
        self.formula_edit.setMinimumHeight(130)
        layout.addWidget(self.formula_edit, 1)
        self.status_label = _label()
        self.status_label.setProperty("role", "muted")
        layout.addWidget(self.status_label)
        self.note_label = _label(
            "在源数据旁新增空白列，把公式粘贴到相同行，再向下填充。"
            "引用按填写的单元格生成；调整 Excel 列位置后请核对引用。\n"
            "两种模式均保留数字、文本原有类型及空白显示。CHAR(160) 使用 UNICHAR，需 Excel 2013 或更新版本。"
        )
        self.note_label.setProperty("role", "muted")
        layout.addWidget(self.note_label)
        buttons = QHBoxLayout()
        buttons.addStretch()
        self.copy_button = QPushButton("复制公式")
        primary_button(self.copy_button, "check")
        self.close_button = QPushButton("关闭")
        buttons.addWidget(self.copy_button)
        buttons.addWidget(self.close_button)
        layout.addLayout(buttons)
        self.close_button.clicked.connect(self.accept)
        self.copy_button.clicked.connect(self._copy_formula)
        self.reference_edit.textChanged.connect(self._update_formula)
        self.strategy_combo.currentIndexChanged.connect(self._update_formula)
        self.mode_combo.currentIndexChanged.connect(self._update_formula)
        for box in self.rule_boxes.values():
            box.toggled.connect(self._update_formula)
        self._update_formula()

    def _update_formula(self, *_args) -> None:
        strategy = self.strategy_combo.currentData()
        mode = self.mode_combo.currentData()
        trims_spaces = self.rule_boxes["outer_spaces"].isChecked()
        if mode == "simple":
            note = (
                "简洁版：去除首尾空格，并将中间连续普通空格合并为一个；需要保留时请选择严格版。"
                if trims_spaces else "简洁版：未勾选“去除前后普通空格”，中间连续普通空格保持不变。"
            )
        else:
            note = (
                "严格版：只去除首尾普通空格，保留中间连续普通空格；公式较长。"
                if trims_spaces else "严格版：按已选字符替换规则处理，中间连续普通空格保持不变。"
            )
        self.mode_note_label.setText(note)
        for rule in RULE_ORDER:
            self.rule_boxes[rule].setText(rule_label(rule, strategy))
        try:
            unsupported = [rule_label(rule) for rule in TYPE_RULE_ORDER if self.rule_boxes[rule].isChecked()]
            if unsupported:
                raise ValueError(
                    "以下规则暂不支持独立公式：" + "、".join(unsupported)
                    + "。取消这些选项后可生成字符清洗公式。"
                )
            rules = CleaningRules(
                **{rule: self.rule_boxes[rule].isChecked() for rule in RULE_ORDER},
                hidden_strategy=strategy,
            )
            formula = build_manual_formula(rules, self.reference_edit.text(), mode=mode)
        except ValueError as error:
            # 无效引用或不支持的规则立即清空旧公式，避免复制过期结果。
            self.formula_edit.clear()
            self.copy_button.setEnabled(False)
            self.status_label.setText(str(error))
            return
        self.formula_edit.setPlainText(formula)
        self.copy_button.setEnabled(True)
        self.status_label.setText(
            "公式已更新，可直接复制；未执行清洗，也未修改任何文件。"
            if any(box.isChecked() for box in self.rule_boxes.values())
            else "未勾选清洗规则，当前公式只引用源值；请按需选择规则。"
        )

    def _copy_formula(self) -> None:
        if self.copy_button.isEnabled() and self.formula_edit.toPlainText():
            QApplication.clipboard().setText(self.formula_edit.toPlainText())
            mode_name = "简洁版" if self.mode_combo.currentData() == "simple" else "严格版"
            self.status_label.setText(f"已复制{mode_name}公式。请在新列的对应行粘贴，再向下填充。")
