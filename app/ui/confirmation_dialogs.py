"""工作副本修改前的中文确认对话框。"""

from pathlib import Path

from PySide6.QtCore import Qt, QUrl
from PySide6.QtGui import QDesktopServices
from PySide6.QtWidgets import QDialog, QHBoxLayout, QLabel, QMessageBox, QPlainTextEdit, QPushButton, QRadioButton, QVBoxLayout, QWidget

from app.core.cleaner import CleaningPreview, ALL_RULE_ORDER, rule_label
from app.core.batch_cleaner import BatchCleaningPreview, rules_description
from app.core.annotation_builder import AnnotationPreview
from app.models.export_data import ExportMode, ExportPlan, ExportResult
from app.ui.theme import primary_button


def _confirm(parent: QWidget, title: str, text: str, note: str, action: str) -> bool:
    dialog = QMessageBox(parent)
    dialog.setWindowTitle(title)
    dialog.setIcon(QMessageBox.Icon.Question)
    dialog.setTextFormat(Qt.TextFormat.PlainText)
    dialog.setText(text)
    dialog.setInformativeText(note)
    dialog.setMinimumWidth(480)
    # 限制正文宽度，不依赖 Qt 内部布局的列位置，也不拉宽图标或间隔。
    for label in dialog.findChildren(QLabel):
        if label.text() in (text, note):
            label.setMinimumWidth(360)
    cancel = dialog.addButton("取消", QMessageBox.ButtonRole.RejectRole)
    accept = dialog.addButton(action, QMessageBox.ButtonRole.AcceptRole)
    dialog.setDefaultButton(cancel)
    dialog.setEscapeButton(cancel)
    dialog.exec()
    return dialog.clickedButton() is accept


def confirm_apply(parent: QWidget, sheet_name: str, column_name: str, preview: CleaningPreview) -> bool:
    rules = "\n".join(
        "• " + rule_label(rule, preview.rules.hidden_strategy)
        for rule in ALL_RULE_ORDER if getattr(preview.rules, rule)
    )
    return _confirm(
        parent, "确认应用清洗",
        f"即将修改工作副本：\n工作表：{sheet_name}\n字段：{column_name}"
        f"\n影响数据：{preview.changed_count} 行\n\n执行规则：\n{rules}",
        "该操作只修改当前工作副本，不会覆盖原 Excel 文件，并且可以撤销。"
        + ("\n已明确确认该列为日期字段；歧义、无法识别及含时间的值将保留。" if preview.rules.date_normalize else "")
        + ("\n带前导0或存在精度风险的文本数字将保留，不批量转换。" if preview.rules.numeric_conversion else ""),
        "确认应用",
    )


def confirm_batch_apply(parent: QWidget, preview: BatchCleaningPreview) -> bool:
    dialog = QDialog(parent)
    dialog.setWindowTitle("确认应用批量清洗")
    dialog.resize(600, 460)
    layout = QVBoxLayout(dialog)
    summary = QLabel(
        f"即将应用批量清洗\n工作表：{preview.sheet_name}\n字段：{len(preview.configs)}"
        f"\n修改单元格：{preview.changed_cell_count}\n修改数据行：{preview.changed_row_count}"
    )
    summary.setTextFormat(Qt.TextFormat.PlainText)
    layout.addWidget(summary)
    rules = QPlainTextEdit()
    rules.setReadOnly(True)
    rules.setPlainText("\n".join(
        f"{config.column_name}：{rules_description(config)}" for config in preview.configs
    ))
    layout.addWidget(rules, 1)
    note = QLabel("操作只会修改当前内存工作副本，不会覆盖原 Excel 文件，并且支持一次撤销整批操作。")
    note.setTextFormat(Qt.TextFormat.PlainText)
    note.setWordWrap(True)
    layout.addWidget(note)
    buttons = QHBoxLayout()
    buttons.addStretch()
    cancel = QPushButton("取消")
    accept = QPushButton("确认应用")
    cancel.setDefault(True)
    accept.setAutoDefault(False)
    primary_button(accept, "check")
    buttons.addWidget(cancel)
    buttons.addWidget(accept)
    layout.addLayout(buttons)
    cancel.clicked.connect(dialog.reject)
    accept.clicked.connect(dialog.accept)
    cancel.setFocus()
    return dialog.exec() == QDialog.DialogCode.Accepted


def confirm_annotation(parent: QWidget, sheet_name: str, preview: AnnotationPreview) -> bool:
    return _confirm(
        parent, "确认应用辅助列",
        f"工作表：{sheet_name}\n源字段：{preview.source_name}"
        f"\n新增辅助列：{preview.annotation.column_name}\n位置：当前工作表数据区域最右侧"
        f"\n数据行数：{preview.row_count}"
        + (f"\n重复记录：{preview.duplicate_count}" if preview.annotation.kind == "duplicate"
           else f"\n需关注记录：{preview.problem_count}")
        + f"\n空值：{preview.empty_count}",
        "只添加标注列，不删除、合并或修改原有记录。应用仅修改工作副本，可以撤销；"
        + ("导出时写入公式。" if preview.annotation.kind == "duplicate" else "导出时保存检测标签。"),
        "确认应用辅助列",
    )


def confirm_restore(parent: QWidget, sheet_name: str) -> bool:
    return _confirm(
        parent, "确认恢复原始数据", f"工作表：{sheet_name}\n是否恢复此工作表的原始数据？",
        "这会取消当前工作表中的所有清洗修改，并删除后来添加的辅助列，但不会影响磁盘中的原 Excel 文件。恢复操作也可以撤销。",
        "确认恢复",
    )


class ExportConfirmationDialog(QDialog):
    """在同一次确认中选择导出模式，默认新增对照列，不占用主窗口表格空间。"""

    def __init__(self, parent: QWidget, plan: ExportPlan, target: Path) -> None:
        super().__init__(parent)
        self.setWindowTitle("确认导出清洗结果")
        self.setMinimumWidth(540)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 18, 20, 18)
        layout.setSpacing(14)
        self.summary_label = QLabel(
            f"源文件：{plan.snapshot.source_path.name}\n修改工作表：{plan.sheet_count}"
            f"\n修改单元格：{plan.cell_count}\n新增辅助列：{plan.annotation_column_count}\n\n目标文件：\n{target}"
        )
        self.summary_label.setTextFormat(Qt.TextFormat.PlainText)
        self.summary_label.setWordWrap(True)
        self.summary_label.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        layout.addWidget(self.summary_label)
        title = QLabel("选择导出模式")
        title.setProperty("role", "sectionTitle")
        layout.addWidget(title)
        self.formula_radio = QRadioButton("新增对照列（公式） · 推荐")
        self.formula_radio.setChecked(True)
        self.update_radio = QRadioButton("原列更新")
        layout.addWidget(self.formula_radio)
        layout.addWidget(self.update_radio)
        self.mode_note_label = QLabel()
        self.mode_note_label.setWordWrap(True)
        self.mode_note_label.setTextFormat(Qt.TextFormat.PlainText)
        layout.addWidget(self.mode_note_label)
        self.safety_label = QLabel("将导出所有已修改工作表。原文件不会被修改，也不会覆盖已有文件。")
        self.safety_label.setWordWrap(True)
        layout.addWidget(self.safety_label)
        buttons = QHBoxLayout()
        buttons.addStretch()
        self.cancel_button = QPushButton("取消")
        self.accept_button = QPushButton("确认导出")
        self.cancel_button.setDefault(True)
        self.accept_button.setAutoDefault(False)
        primary_button(self.accept_button, "export")
        buttons.addWidget(self.cancel_button)
        buttons.addWidget(self.accept_button)
        layout.addLayout(buttons)
        self.cancel_button.clicked.connect(self.reject)
        self.accept_button.clicked.connect(self.accept)
        self.formula_radio.toggled.connect(lambda: self._update_note(plan))
        self._update_note(plan)
        self.cancel_button.setFocus()

    def selected_mode(self) -> ExportMode:
        return ExportMode.ADD_FORMULA_COLUMNS if self.formula_radio.isChecked() else ExportMode.UPDATE_COLUMNS

    def _update_note(self, plan: ExportPlan) -> None:
        if self.selected_mode() == ExportMode.ADD_FORMULA_COLUMNS:
            self.mode_note_label.setText(
                f"保留原始列，在每个清洗列右侧新增“字段名_清洗后”，共 {plan.column_count} 列。"
                "普通空白清洗列对全部数据行写入公式；高级类型处理列保存预览确认后的值。"
                "复杂公式或结构无法安全调整时会取消导出，可改选“原列更新”。"
            )
        else:
            self.mode_note_label.setText(
                "在导出的副本中替换原列的清洗值，原有列数和顺序保持不变；保存的是清洗后的值。"
            )
        if plan.annotation_column_count:
            self.mode_note_label.setText(
                self.mode_note_label.text()
                + f"\n另在数据区域最右侧追加 {plan.annotation_column_count} 个辅助检查列。"
                "重复检查列写入公式，类型、前导0及日期检查列保存分析标签。"
            )


def confirm_export(parent: QWidget, plan: ExportPlan, target: Path) -> ExportMode | None:
    dialog = ExportConfirmationDialog(parent, plan, target)
    if dialog.exec() != QDialog.DialogCode.Accepted:
        return None
    return dialog.selected_mode()


def show_export_success(parent: QWidget, result: ExportResult) -> None:
    dialog = QMessageBox(parent)
    dialog.setWindowTitle("导出成功")
    dialog.setIcon(QMessageBox.Icon.Information)
    dialog.setTextFormat(Qt.TextFormat.PlainText)
    dialog.setText(
        f"清洗结果已保存\n\n保存位置：\n{result.target_path}"
        f"\n\n导出模式：{result.mode.label}"
        f"\n修改工作表：{result.sheet_count}\n修改单元格：{result.cell_count}"
        + (f"\n新增对照列：{result.added_column_count}" if result.added_column_count else "")
        + (f"\n新增辅助列：{result.annotation_column_count}" if result.annotation_column_count else "")
    )
    note = "源文件保持不变；工作副本和撤销历史仍然保留。"
    if result.mode == ExportMode.ADD_FORMULA_COLUMNS or result.annotation_column_count:
        note += "\n普通空白清洗和重复检查列包含公式；高级类型处理及其他检查列保存已确认的值。"
    dialog.setInformativeText(note)
    folder = dialog.addButton("打开所在文件夹", QMessageBox.ButtonRole.ActionRole)
    close = dialog.addButton("确定", QMessageBox.ButtonRole.AcceptRole)
    dialog.setDefaultButton(close)
    dialog.setEscapeButton(close)
    dialog.exec()
    if dialog.clickedButton() is folder:
        QDesktopServices.openUrl(QUrl.fromLocalFile(str(result.target_path.parent)))
