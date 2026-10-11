"""集中管理界面主题和图标，业务模块不依赖这些样式。"""

from pathlib import Path

from PySide6.QtCore import QSize
from PySide6.QtGui import QFont, QFontDatabase, QIcon
from PySide6.QtWidgets import QApplication, QLabel, QPushButton, QTableView


UI_ROOT = Path(__file__).resolve().parent


def apply_theme(application: QApplication) -> None:
    """应用统一的浅色主题；主窗口单独创建时也可以安全调用。"""
    if application.property("excel_theme_applied"):
        return
    application.setStyle("Fusion")
    families = set(QFontDatabase.families())
    for family in ("Microsoft YaHei UI", "Microsoft YaHei", "Noto Sans CJK SC"):
        if family in families:
            application.setFont(QFont(family, 10))
            break
    stylesheet = (UI_ROOT / "styles" / "light.qss").read_text(encoding="utf-8")
    # 使用资源所在目录，避免依赖用户的启动目录。
    stylesheet = stylesheet.replace("@ASSETS@", (UI_ROOT / "assets").as_posix())
    application.setStyleSheet(stylesheet)
    application.setProperty("excel_theme_applied", True)


def icon(name: str) -> QIcon:
    return QIcon(str(UI_ROOT / "assets" / f"{name}.svg"))


def primary_button(button: QPushButton, icon_name: str) -> None:
    button.setProperty("variant", "primary")
    button.setIcon(icon(icon_name))
    button.setIconSize(QSize(16, 16))
    button.setMinimumWidth(108)


def style_table(table: QTableView) -> None:
    """只设置表格外观，不改变编辑、选择或列宽调整行为。"""
    table.setShowGrid(False)
    table.verticalHeader().setDefaultSectionSize(28)
    table.verticalHeader().setMinimumSectionSize(26)
    table.horizontalHeader().setMinimumHeight(32)


def set_tone(label: QLabel, tone: str) -> None:
    """更新统计徽标的颜色，保留标签中的统计文本。"""
    label.setProperty("tone", tone)
    label.style().unpolish(label)
    label.style().polish(label)
    label.update()
