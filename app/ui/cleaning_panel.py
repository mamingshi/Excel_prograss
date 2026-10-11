"""推荐清洗规则和 Excel 参考公式，不执行清洗。"""

from PySide6.QtCore import QSignalBlocker, Signal, Qt
from PySide6.QtWidgets import (
    QCheckBox,
    QComboBox,
    QGroupBox,
    QLabel,
    QPlainTextEdit,
    QPushButton,
    QVBoxLayout,
)

from app.core.cleaner import CleaningRules, RULE_ORDER, rule_label
from app.core.detector import ColumnAnalysis
from app.core.formula_builder import FORMULA_NOTE, build_formula
from app.ui.theme import primary_button
from app.ui.type_rules_panel import TypeRulesPanel


class CleaningPanel(QGroupBox):
    preview_requested = Signal()
    preview_invalidated = Signal()
    annotation_requested = Signal()
    type_annotation_requested = Signal(str)
    annotation_invalidated = Signal()

    def __init__(self, parent=None) -> None:
        super().__init__("推荐清洗", parent)
        self.setObjectName("cleaningPanel")
        self.analysis: ColumnAnalysis | None = None
        self._column_name: str | None = None
        self._cell_reference: str | None = None
        self._busy = False
        self._is_generated = False
        self.setMinimumWidth(285)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 10, 12, 12)
        layout.setSpacing(6)

        self.safety_label = QLabel("预览不会修改原始数据；应用仅修改工作副本。")
        self.safety_label.setWordWrap(True)
        self.safety_label.setProperty("role", "badge")
        self.safety_label.setProperty("tone", "success")
        layout.addWidget(self.safety_label)
        self.column_label = QLabel("当前列：未选择")
        self.column_label.setTextFormat(Qt.TextFormat.PlainText)
        self.column_label.setWordWrap(True)
        self.column_label.setProperty("role", "sectionTitle")
        layout.addWidget(self.column_label)
        self.state_label = QLabel("请先完成当前列的数据检测。")
        self.state_label.setWordWrap(True)
        self.state_label.setProperty("role", "muted")
        layout.addWidget(self.state_label)

        self.rule_boxes = {}
        for rule in RULE_ORDER:
            box = QCheckBox(rule_label(rule))
            box.toggled.connect(self._rules_changed)
            self.rule_boxes[rule] = box
            layout.addWidget(box)

        layout.addWidget(QLabel("Tab / 换行 / 回车处理策略："))
        self.strategy_combo = QComboBox()
        self.strategy_combo.addItem("替换为空格（推荐）", "space")
        self.strategy_combo.addItem("直接删除", "delete")
        self.strategy_combo.setToolTip("替换为空格可保留词语分隔；直接删除可能拼接相邻文本。")
        self.strategy_combo.currentIndexChanged.connect(self._rules_changed)
        layout.addWidget(self.strategy_combo)

        self.empty_label = QLabel("空值：— 条；仅提示，暂不自动处理。")
        self.duplicate_label = QLabel("重复值：— 条；仅提示，暂不自动删除。")
        self.empty_label.setWordWrap(True)
        self.duplicate_label.setWordWrap(True)
        self.empty_label.setProperty("role", "muted")
        self.duplicate_label.setProperty("role", "muted")
        layout.addWidget(self.empty_label)
        layout.addWidget(self.duplicate_label)
        self.type_panel = TypeRulesPanel(self)
        self.type_panel.rules_changed.connect(self._rules_changed)
        self.type_panel.annotation_requested.connect(self.type_annotation_requested.emit)
        self.type_panel.annotation_context_changed.connect(self.annotation_invalidated.emit)
        layout.addWidget(self.type_panel)
        self.generate_button = QPushButton("生成清洗预览")
        primary_button(self.generate_button, "preview")
        self.generate_button.clicked.connect(self.preview_requested.emit)
        layout.addWidget(self.generate_button)

        annotation_group = QGroupBox("异常标注 / 辅助列")
        annotation_layout = QVBoxLayout(annotation_group)
        annotation_note = QLabel("重复值先标注，供人工判断；不删除、合并或覆盖原有记录。")
        annotation_note.setWordWrap(True)
        annotation_note.setProperty("role", "muted")
        annotation_layout.addWidget(annotation_note)
        self.annotation_button = QPushButton("生成重复检查列")
        self.annotation_button.setIcon(self.generate_button.icon())
        self.annotation_button.clicked.connect(self.annotation_requested.emit)
        annotation_layout.addWidget(self.annotation_button)
        layout.addWidget(annotation_group)

        formula_group = QGroupBox("公式参考")
        formula_group.setObjectName("formulaGroup")
        formula_layout = QVBoxLayout(formula_group)
        self.formula_edit = QPlainTextEdit()
        self.formula_edit.setObjectName("formulaEditor")
        self.formula_edit.setReadOnly(True)
        self.formula_edit.setFixedHeight(75)
        self.formula_edit.setPlaceholderText("完成检测后，参考公式随勾选规则更新。")
        formula_layout.addWidget(self.formula_edit)
        self.formula_note = QLabel(FORMULA_NOTE)
        self.formula_note.setWordWrap(True)
        self.formula_note.setProperty("role", "muted")
        formula_layout.addWidget(self.formula_note)
        layout.addWidget(formula_group)
        layout.addStretch()
        self.set_busy(False)

    def current_rules(self) -> CleaningRules:
        return CleaningRules(
            **{rule: box.isChecked() for rule, box in self.rule_boxes.items()},
            hidden_strategy=self.strategy_combo.currentData(),
            **self.type_panel.rule_values(),
        )

    def set_column(self, column_name: str | None, cell_reference: str | None, *, is_generated: bool = False) -> None:
        self._is_generated = is_generated
        self.annotation_invalidated.emit()
        self._column_name = column_name
        self._cell_reference = cell_reference
        self.analysis = None
        self.type_panel.set_analysis(None)
        self.column_label.setText(f"当前列：{column_name if column_name is not None else '未选择'}")
        with QSignalBlocker(self.strategy_combo):
            self.strategy_combo.setCurrentIndex(0)
        for rule, box in self.rule_boxes.items():
            with QSignalBlocker(box):
                box.setChecked(False)
            box.setText(rule_label(rule))
        self.state_label.setText("辅助列随源字段更新，请选择原始字段进行处理。" if is_generated else "请先完成当前列的数据检测。")
        self.empty_label.setText("空值：— 条；仅提示，暂不自动处理。")
        self.duplicate_label.setText("重复值：— 条；仅提示，暂不自动删除。")
        self.formula_edit.clear()
        self.preview_invalidated.emit()
        self.set_busy(self._busy)

    def set_analysis(self, analysis: ColumnAnalysis) -> None:
        self.analysis = analysis
        self.type_panel.set_analysis(analysis)
        # 隐藏字符的细分建议根据实际原始值生成，不把三种字符混为一条规则。
        suggestions = {
            "nbsp": analysis.counts["CHAR(160)"] > 0,
            "tab": any(isinstance(issue.original_value, str) and "\t" in issue.original_value for issue in analysis.issues),
            "newline": any(isinstance(issue.original_value, str) and "\n" in issue.original_value for issue in analysis.issues),
            "carriage_return": any(isinstance(issue.original_value, str) and "\r" in issue.original_value for issue in analysis.issues),
            # 替换可能在边缘产生普通空格，因此有空白问题时也建议最后去首尾空格。
            "outer_spaces": any(analysis.counts[name] > 0 for name in ("前后空格", "CHAR(160)", "Tab / 换行")),
        }
        for rule, box in self.rule_boxes.items():
            with QSignalBlocker(box):
                box.setChecked(suggestions[rule])
        self.empty_label.setText(f"空值：{analysis.counts['空值']} 条；仅提示，暂不自动处理。")
        self.duplicate_label.setText(f"重复值：{analysis.counts['重复值']} 条；可生成辅助列供人工检查，不删除记录。")
        self.state_label.setText("辅助列随源字段更新，请选择原始字段进行处理。" if self._is_generated else "已根据检测结果勾选建议，可手动调整。")
        self.preview_invalidated.emit()
        self._update_formula()
        self.set_busy(self._busy)

    def set_busy(self, busy: bool) -> None:
        self._busy = busy
        enabled = not busy and self.analysis is not None and not self._is_generated
        self.annotation_button.setEnabled(enabled and self.analysis.counts["重复值"] > 0)
        self.generate_button.setEnabled(enabled)
        self.type_panel.set_busy(not enabled)
        self.strategy_combo.setEnabled(enabled)
        for box in self.rule_boxes.values():
            box.setEnabled(enabled)

    def _update_formula(self) -> None:
        if self.analysis is not None and self._cell_reference is not None:
            self.formula_edit.setPlainText(build_formula(self.current_rules(), self._cell_reference))

    def _rules_changed(self, *_args) -> None:
        strategy = self.strategy_combo.currentData()
        for rule, box in self.rule_boxes.items():
            box.setText(rule_label(rule, strategy))
        self.preview_invalidated.emit()
        self._update_formula()
        if self.analysis is not None:
            self.state_label.setText("规则已调整，请重新生成清洗预览。")
