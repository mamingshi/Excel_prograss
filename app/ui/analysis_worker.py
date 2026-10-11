"""后台列检测任务，不读取或保存文件。"""

import pandas as pd
from PySide6.QtCore import QThread, Signal

from app.core.detector import analyze_column


class ColumnAnalysisWorker(QThread):
    succeeded = Signal(object)
    failed = Signal(str)

    def __init__(self, column: pd.Series, parent=None) -> None:
        super().__init__(parent)
        self._column = column

    def run(self) -> None:
        try:
            result = analyze_column(self._column)
        except Exception as error:
            self.failed.emit(str(error))
        else:
            self.succeeded.emit(result)
