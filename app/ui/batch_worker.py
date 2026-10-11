"""批量预览后台任务：只读快照、协作取消和逐字段进度。"""

from PySide6.QtCore import QThread, Signal

from app.core.batch_cleaner import generate_batch_preview
from app.core.sheet_analyzer import ScanCancelled


class BatchPreviewWorker(QThread):
    progress = Signal(int, int, str)
    succeeded = Signal(object)
    failed = Signal(str)
    cancelled = Signal()

    def __init__(self, snapshot, configs, original_column_count, parent=None):
        super().__init__(parent)
        self.snapshot = snapshot
        self.configs = tuple(configs)
        self.original_column_count = original_column_count

    def run(self):
        try:
            result = generate_batch_preview(
                self.snapshot.dataframe, self.configs, self.snapshot.sheet_name,
                original_column_count=self.original_column_count,
                progress=self.progress.emit, is_cancelled=self.isInterruptionRequested,
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
