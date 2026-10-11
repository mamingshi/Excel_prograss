"""后台生成独立的清洗预览，不读取或写入文件。"""

import pandas as pd
from PySide6.QtCore import QThread, Signal

from app.core.cleaner import CleaningRules, generate_cleaning_preview


class CleaningPreviewWorker(QThread):
    succeeded = Signal(object)
    failed = Signal(str)

    def __init__(self, column: pd.Series, rules: CleaningRules, parent=None) -> None:
        super().__init__(parent)
        self._column = column
        self._rules = rules

    def run(self) -> None:
        try:
            result = generate_cleaning_preview(self._column, self._rules)
        except Exception as error:
            self.failed.emit(str(error))
        else:
            self.succeeded.emit(result)
