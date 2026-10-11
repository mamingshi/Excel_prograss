"""工作副本的操作快照，不涉及 Excel 文件读写。"""

from dataclasses import dataclass, field
from datetime import datetime, timezone

import pandas as pd

from app.core.cleaner import CleaningRules
from app.core.annotation_builder import AnnotationColumn


@dataclass(frozen=True)
class CleaningAction:
    sheet_name: str
    column_index: int
    column_name: str
    # 保存完整列以恢复类型和缺失值；位置不依赖可能重复的 DataFrame index。
    before_values: pd.Series
    after_values: pd.Series
    affected_indices: tuple[int, ...]
    rules: CleaningRules
    previous_rules: tuple[CleaningRules, ...] = ()
    timestamp: datetime = field(default_factory=lambda: datetime.now(timezone.utc))
    operation_name: str = "应用清洗"


@dataclass(frozen=True)
class RestoreAction:
    sheet_name: str
    before_values: pd.DataFrame
    previous_operation_count: int
    previous_modified_columns: frozenset[int]
    previous_column_rules: dict[int, tuple[CleaningRules, ...]] = field(default_factory=dict)
    previous_annotations: tuple[AnnotationColumn, ...] = ()
    timestamp: datetime = field(default_factory=lambda: datetime.now(timezone.utc))
    operation_name: str = "恢复原始"


@dataclass(frozen=True)
class AnnotationAction:
    sheet_name: str
    column_index: int
    column_name: str
    source_column_index: int
    previous_operation_count: int
    timestamp: datetime = field(default_factory=lambda: datetime.now(timezone.utc))
    operation_name: str = "添加重复检查列"


HistoryAction = CleaningAction | RestoreAction | AnnotationAction
