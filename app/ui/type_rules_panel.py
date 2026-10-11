"""高级类型规则与辅助列入口；保持预览、人工确认、应用的现有流程。"""

from PySide6.QtCore import QSignalBlocker, Signal
from PySide6.QtWidgets import QCheckBox, QGroupBox, QLabel, QPushButton, QVBoxLayout

from app.core.cleaner import TYPE_RULE_ORDER, rule_label
from app.core.detector import ColumnAnalysis


class TypeRulesPanel(QGroupBox):
    rules_changed = Signal()
    annotation_requested = Signal(str)
    annotation_context_changed = Signal()

    def __init__(self, parent=None) -> None:
        super().__init__("高级类型处理", parent)
        self.analysis: ColumnAnalysis | None = None
        self._busy = True
        layout = QVBoxLayout(self)
        self.note_label = QLabel(
            "前导0和精度风险保留原值；日期需明确确认，歧义或无法识别的数据不转换。"
        )
        self.note_label.setWordWrap(True)
        self.note_label.setProperty("role", "muted")
        layout.addWidget(self.note_label)
        self.rule_boxes = {}
        for rule in TYPE_RULE_ORDER:
            box = QCheckBox(rule_label(rule))
            box.toggled.connect(lambda _checked: self.rules_changed.emit())
            self.rule_boxes[rule] = box
            layout.addWidget(box)
        self.date_confirmed_box = QCheckBox("我确认当前列是日期字段")
        self.date_confirmed_box.toggled.connect(self._date_confirmation_changed)
        layout.addWidget(self.date_confirmed_box)
        self.date_state_label = QLabel()
        self.date_state_label.setWordWrap(True)
        self.date_state_label.setProperty("role", "muted")
        layout.addWidget(self.date_state_label)
        self.annotation_buttons = {}
        for kind, text in (
            ("numeric_type", "生成类型检查列"),
            ("leading_zero", "生成前导0风险列"),
            ("date", "生成日期检查列"),
        ):
            button = QPushButton(text)
            button.clicked.connect(lambda _checked=False, selected=kind: self.annotation_requested.emit(selected))
            self.annotation_buttons[kind] = button
            layout.addWidget(button)
        self.set_analysis(None)

    def rule_values(self) -> dict[str, bool]:
        return {**{name: box.isChecked() for name, box in self.rule_boxes.items()},
                "date_confirmed": self.date_confirmed_box.isChecked()}

    def set_analysis(self, analysis: ColumnAnalysis | None) -> None:
        self.analysis = analysis
        for name, box in self.rule_boxes.items():
            with QSignalBlocker(box):
                box.setChecked(bool(analysis and name == "fullwidth" and analysis.counts["全角字符"]))
        with QSignalBlocker(self.date_confirmed_box):
            self.date_confirmed_box.setChecked(False)
        if analysis is None:
            text = "完成检测后可选择处理规则。"
        else:
            profile = analysis.date_profile
            text = (f"日期证据：{profile.recognized_count}/{profile.nonempty_count} 个非空值可明确解析。"
                    if profile else "尚无日期识别统计。")
            text += " 疑似日期列。" if profile and profile.suspected else " 不自动按日期列检测。"
        self.date_state_label.setText(text)
        self.set_busy(self._busy)

    def set_busy(self, busy: bool) -> None:
        self._busy = busy
        enabled = not busy and self.analysis is not None
        for name, box in self.rule_boxes.items():
            box.setEnabled(enabled and (name != "date_normalize" or self.date_confirmed_box.isChecked()))
        self.date_confirmed_box.setEnabled(enabled)
        counts = self.analysis.counts if self.analysis else {}
        self.annotation_buttons["numeric_type"].setEnabled(enabled and counts.get("文本数字混合", 0) > 0)
        self.annotation_buttons["leading_zero"].setEnabled(enabled and counts.get("前导0风险", 0) > 0)
        profile = self.analysis.date_profile if self.analysis else None
        self.annotation_buttons["date"].setEnabled(
            enabled and (bool(profile and profile.suspected) or self.date_confirmed_box.isChecked())
        )

    def _date_confirmation_changed(self, confirmed: bool) -> None:
        if not confirmed:
            with QSignalBlocker(self.rule_boxes["date_normalize"]):
                self.rule_boxes["date_normalize"].setChecked(False)
        self.set_busy(self._busy)
        self.annotation_context_changed.emit()
        self.rules_changed.emit()
