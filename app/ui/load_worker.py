"""在后台线程读取工作簿，保持主界面响应。"""

from pathlib import Path

from PySide6.QtCore import QThread, Signal

from app.core.file_loader import ExcelFileLoader, ExcelLoadError


class ExcelLoadWorker(QThread):
    succeeded = Signal(object)
    failed = Signal(str, str)

    def __init__(self, file_path: str | Path, sheet_name: str | None = None, parent=None):
        super().__init__(parent)
        self._file_path = file_path
        self._sheet_name = sheet_name

    def run(self) -> None:
        try:
            preview = ExcelFileLoader().load(self._file_path, self._sheet_name)
        except ExcelLoadError as error:
            self.failed.emit(str(error), error.detail)
        else:
            self.succeeded.emit(preview)
