"""后台生成辅助列预览，不修改工作副本或 Excel 文件。"""

import pandas as pd
from PySide6.QtCore import QThread, Signal

from app.core.annotation_builder import generate_duplicate_preview, generate_annotation_preview


class AnnotationPreviewWorker(QThread):
    succeeded = Signal(object)
    failed = Signal(str)

    def __init__(self, dataframe: pd.DataFrame, column_index: int, parent=None, *,
                 kind: str = "duplicate", date_confirmed: bool = False) -> None:
        super().__init__(parent)
        self._dataframe = dataframe
        self._column_index = column_index
        self._kind = kind
        self._date_confirmed = date_confirmed

    def run(self) -> None:
        try:
            result = (generate_duplicate_preview(self._dataframe, self._column_index)
                      if self._kind == "duplicate" else
                      generate_annotation_preview(self._dataframe, self._column_index, self._kind,
                                                  date_confirmed=self._date_confirmed))
        except Exception as error:
            self.failed.emit(str(error))
        else:
            self.succeeded.emit(result)
