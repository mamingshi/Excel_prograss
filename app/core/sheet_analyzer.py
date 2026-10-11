"""V0.6 工作表只读扫描：复用列检测，返回轻量统计，不保存全表问题明细。"""

from collections.abc import Callable
from dataclasses import dataclass

import pandas as pd

from app.core.detector import analyze_column


class ScanCancelled(Exception):
    """协作取消：不把部分列结果当作完整扫描发布。"""


@dataclass(frozen=True)
class ColumnQuality:
    column_index: int
    column_name: str
    total_rows: int
    affected_rows: int
    outer_spaces: int
    nbsp: int
    hidden_chars: int
    empty: int
    duplicates: int
    quality_score: int
    status: str
    summary: str
    numeric_mixture: int = 0
    leading_zero: int = 0
    date_anomalies: int = 0
    fullwidth: int = 0


@dataclass(frozen=True)
class SheetAnalysis:
    row_count: int
    columns: tuple[ColumnQuality, ...]

    @property
    def column_count(self) -> int:
        return len(self.columns)

    @property
    def problem_column_count(self) -> int:
        return sum(column.affected_rows > 0 for column in self.columns)

    @property
    def normal_column_count(self) -> int:
        return self.column_count - self.problem_column_count


def quality_score(affected_rows: int, total_rows: int) -> int:
    # 仅有表头时没有数据问题，记为 100 分；全空的非空长度列仍按空值扣分。
    if total_rows <= 0:
        return 100
    return max(0, min(100, round(100 * (1 - affected_rows / total_rows))))


def quality_grade(score: int) -> str:
    if score >= 95:
        return "优秀"
    if score >= 85:
        return "良好"
    if score >= 70:
        return "建议清洗"
    return "问题较多"


def _summary(counts: dict[str, int]) -> str:
    labels = (("前后空格", "空格"), ("CHAR(160)", "CHAR160 "), ("Tab / 换行", "Tab/换行"), ("空值", "空值"), ("重复值", "重复"))
    labels += (("文本数字混合", "文本数字"), ("前导0风险", "前导0"), ("日期异常", "日期异常"), ("全角字符", "全角"))
    problems = [(label, counts[name]) for name, label in labels if counts[name]]
    problems.sort(key=lambda item: item[1], reverse=True)
    return " / ".join(f"{label}{count}" for label, count in problems[:3]) or "未发现当前支持的问题"


def analyze_sheet(
    dataframe: pd.DataFrame, *,
    progress: Callable[[int, int, str], None] | None = None,
    is_cancelled: Callable[[], bool] | None = None,
) -> SheetAnalysis:
    """按物理列位置扫描；重名字段也有独立结果，调用方负责提供稳定的只读数据。"""
    results = []
    total_columns = dataframe.shape[1]
    for column_index in range(total_columns):
        if is_cancelled is not None and is_cancelled():
            raise ScanCancelled()
        # 与单列入口共用规则；不为 10 万行 × 50 列积累数百万条 IssueRecord。
        analysis = analyze_column(dataframe.iloc[:, column_index], include_details=False)
        counts = analysis.counts
        score = quality_score(analysis.affected_rows, analysis.row_count)
        results.append(ColumnQuality(
            column_index=column_index, column_name=analysis.column_name,
            total_rows=analysis.row_count, affected_rows=analysis.affected_rows,
            outer_spaces=counts["前后空格"], nbsp=counts["CHAR(160)"],
            hidden_chars=counts["Tab / 换行"], empty=counts["空值"],
            duplicates=counts["重复值"], quality_score=score,
            status=quality_grade(score), summary=_summary(counts),
            numeric_mixture=counts["文本数字混合"], leading_zero=counts["前导0风险"],
            date_anomalies=counts["日期异常"], fullwidth=counts["全角字符"],
        ))
        if progress is not None:
            progress(column_index + 1, total_columns, analysis.column_name)
    if is_cancelled is not None and is_cancelled():
        raise ScanCancelled()
    return SheetAnalysis(len(dataframe), tuple(results))
