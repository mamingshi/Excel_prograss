"""Excel 数据清洗助手的启动入口。"""

import sys

from PySide6.QtCore import QLibraryInfo, QTranslator
from PySide6.QtWidgets import QApplication

from app.ui.main_window import MainWindow
from app.ui.theme import apply_theme


def main() -> int:
    application = QApplication(sys.argv)
    application.setApplicationName("Excel 数据清洗助手")
    application.setApplicationVersion("0.8")
    apply_theme(application)
    # 为 Qt 自带对话框中的按钮和详情标签加载中文翻译。
    translator = QTranslator(application)
    if translator.load(
        "qtbase_zh_CN", QLibraryInfo.path(QLibraryInfo.LibraryPath.TranslationsPath)
    ):
        application.installTranslator(translator)
    window = MainWindow()
    window.show()
    return application.exec()


if __name__ == "__main__":
    sys.exit(main())
