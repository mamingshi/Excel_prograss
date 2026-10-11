"""整表质量总览：只显示统计，字段点击交给现有单列诊断流程。"""

from PySide6.QtCore import QModelIndex, QSortFilterProxyModel, Qt, Signal, Slot
from PySide6.QtWidgets import (
    QAbstractItemView, QHBoxLayout, QHeaderView, QLabel, QProgressBar,
    QPushButton, QTableView, QVBoxLayout, QWidget,
)

from app.core.sheet_analyzer import SheetAnalysis
from app.models.sheet_overview_model import SheetOverviewModel
from app.ui.theme import set_tone, style_table


STALE_SCAN_MESSAGE = "当前数据已发生修改，请重新扫描工作表以更新质量总览。"


class SheetOverviewPanel(QWidget):
    column_requested = Signal(int)
    cancel_requested = Signal()
    batch_requested = Signal()

    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self._has_sheet = False
        self._stale = False
        self._cancel_pending = False
        self.result: SheetAnalysis | None = None
        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(8)

        controls = QHBoxLayout()
        self.summary_label = QLabel("当前工作表：未选择")
        self.summary_label.setTextFormat(Qt.TextFormat.PlainText)
        self.summary_label.setWordWrap(True)
        self.summary_label.setProperty("role", "sectionTitle")
        controls.addWidget(self.summary_label, 1)
        self.batch_button = QPushButton("打开批量清洗工作台")
        self.batch_button.clicked.connect(self.batch_requested.emit)
        controls.addWidget(self.batch_button)
        self.cancel_button = QPushButton("取消扫描")
        self.cancel_button.clicked.connect(self.cancel_requested.emit)
        self.cancel_button.hide()
        controls.addWidget(self.cancel_button)
        layout.addLayout(controls)

        self.state_label = QLabel()
        self.state_label.setTextFormat(Qt.TextFormat.PlainText)
        self.state_label.setWordWrap(True)
        self.state_label.setProperty("role", "badge")
        layout.addWidget(self.state_label)
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.setTextVisible(True)
        self.progress_bar.hide()
        layout.addWidget(self.progress_bar)
        hint = QLabel("点击字段查看单列明细和清洗建议 · 质量分默认从低到高，可点击表头排序 · 95–100 优秀 / 85–94 良好 / 70–84 建议清洗 / 0–69 问题较多")
        hint.setProperty("role", "muted")
        hint.setWordWrap(True)
        layout.addWidget(hint)

        self.model = SheetOverviewModel(self)
        self.proxy_model = QSortFilterProxyModel(self)
        self.proxy_model.setSourceModel(self.model)
        self.table = QTableView()
        style_table(self.table)
        self.table.setModel(self.proxy_model)
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.table.setAlternatingRowColors(True)
        self.table.setWordWrap(False)
        self.table.setHorizontalScrollMode(QAbstractItemView.ScrollMode.ScrollPerPixel)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        self.table.horizontalHeader().sectionDoubleClicked.connect(self.table.resizeColumnToContents)
        self.table.setSortingEnabled(True)
        self.table.sortByColumn(8, Qt.SortOrder.AscendingOrder)
        for index, width in enumerate((170, 80, 100, 90, 95, 100, 70, 80, 80, 105, 320, 110, 100, 100, 100)):
            self.table.setColumnWidth(index, width)
        self.table.clicked.connect(self._activate_column)
        self.table.activated.connect(self._activate_column)
        layout.addWidget(self.table, 1)
        self.set_sheet(None, 0)

    def set_sheet(
        self, sheet_name: str | None, row_count: int,
        result: SheetAnalysis | None = None, *, stale: bool = False,
    ) -> None:
        self._has_sheet = sheet_name is not None
        self._stale = stale
        self.result = result
        self.model.set_result(result)
        if sheet_name is None:
            self.summary_label.setText("当前工作表：未选择")
            self.show_notice("请先打开 Excel，再点击“扫描当前工作表”。")
        elif result is None:
            self.summary_label.setText(f"当前工作表：{sheet_name} · 总数据行：{row_count}")
            self.show_notice("尚未扫描。点击上方“扫描当前工作表”，检查全部字段。")
        else:
            prefix = "上次扫描（已过期）" if stale else "扫描完成"
            self.summary_label.setText(
                f"{prefix} · 当前工作表：{sheet_name} · 字段数：{result.column_count}"
                f" · 总数据行：{row_count} · 存在问题字段：{result.problem_column_count}"
                f" · 正常字段：{result.normal_column_count}"
            )
            self.show_notice(
                STALE_SCAN_MESSAGE if stale else "扫描完成，仅检查当前支持的九类问题。点击字段继续查看；扫描不会自动清洗。",
                "warning" if stale else "success",
            )
        self.set_busy(False)

    def set_busy(self, busy: bool, *, scanning: bool = False) -> None:
        self.batch_button.setEnabled(not busy and self._has_sheet)
        self.table.setEnabled(not busy and self.result is not None and not self._stale)
        self.cancel_button.setVisible(scanning)
        self.cancel_button.setEnabled(scanning and not self._cancel_pending)
        self.progress_bar.setVisible(scanning)

    def begin_scan(self, column_count: int) -> None:
        self._cancel_pending = False
        if self.result is not None and self.summary_label.text().startswith("扫描完成"):
            self.summary_label.setText(self.summary_label.text().replace("扫描完成", "上次完整扫描", 1))
        self.progress_bar.setRange(0, max(1, column_count))
        self.progress_bar.setValue(0)
        self.show_notice(f"正在扫描当前工作表…… 0 / {column_count}")
        self.set_busy(True, scanning=True)

    @Slot(int, int, str)
    def show_progress(self, completed: int, total: int, column_name: str) -> None:
        self.progress_bar.setValue(completed)
        if not self._cancel_pending:
            self.show_notice(f"正在扫描：{column_name}  {completed} / {total}")

    def show_cancel_pending(self) -> None:
        self._cancel_pending = True
        self.cancel_button.setEnabled(False)
        self.show_notice("正在取消扫描，将在当前列分析结束后停止……")

    def show_notice(self, text: str, tone: str = "neutral") -> None:
        self.state_label.setText(text)
        set_tone(self.state_label, tone)

    @Slot(QModelIndex)
    def _activate_column(self, index: QModelIndex) -> None:
        if self.table.isEnabled() and index.isValid():
            source_index = self.proxy_model.mapToSource(index)
            self.column_requested.emit(self.model.data(source_index, Qt.ItemDataRole.UserRole))
