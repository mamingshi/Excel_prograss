"""Qt 界面测试使用离屏平台，不弹出测试窗口。"""

import os

import pytest

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")


@pytest.fixture(scope="session")
def qapp():
    """保持一个 QApplication 活到测试结束，避免反复创建 Qt 应用。"""
    from PySide6.QtWidgets import QApplication

    return QApplication.instance() or QApplication([])
