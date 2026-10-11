"""后台整表扫描：仅分析数据，通过信号报告进度和完整结果。"""

from PySide6.QtCore import QThread, Signal

from app.core.data_manager import SheetScanSnapshot
from app.core.sheet_analyzer import ScanCancelled, analyze_sheet


class SheetAnalysisWorker(QThread):
    progress = Signal(int, int, str)
    succeeded = Signal(object)
    failed = Signal(str)
    cancelled = Signal()

    def __init__(self, snapshot: SheetScanSnapshot, parent=None) -> None:
        super().__init__(parent)
        self._snapshot = snapshot

    def run(self) -> None:
        try:
            result = analyze_sheet(
                self._snapshot.dataframe, progress=self.progress.emit,
                is_cancelled=self.isInterruptionRequested,
            )
        except ScanCancelled:
            self.cancelled.emit()
        except Exception as error:
            self.failed.emit(str(error))
        else:
            if self.isInterruptionRequested():
                self.cancelled.emit()
            else:
                self.succeeded.emit(result)
