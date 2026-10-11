"""批量工作台的逐列配置；以物理位置标识字段，兼容同名表头。"""

from dataclasses import dataclass


@dataclass(frozen=True)
class ColumnCleaningConfig:
    column_index: int
    column_name: str
    enabled: bool = False
    rules: tuple[str, ...] = ()
    warnings: tuple[str, ...] = ()
    detected_problems: str = "尚未扫描"
    date_confirmed: bool = False
    is_generated: bool = False
