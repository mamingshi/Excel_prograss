"""在后台导出已确认的快照，界面仅接收统计信息或友好错误。"""

from pathlib import Path

from PySide6.QtCore import QThread, Signal

from app.core.exporter import ExcelExportError, export_workbook
from app.models.export_data import ExportPlan


class ExcelExportWorker(QThread):
    succeeded = Signal(object)
    failed = Signal(str, str)

    def __init__(self, plan: ExportPlan, target_path: Path, parent=None) -> None:
        super().__init__(parent)
        self.plan = plan
        self.target_path = target_path

    def run(self) -> None:
        try:
            self.succeeded.emit(export_workbook(self.plan, self.target_path))
        except ExcelExportError as error:
            self.failed.emit(str(error), error.detail)
        except Exception as error:
            self.failed.emit("导出失败，源文件保持不变，请查看错误详情。", str(error))
