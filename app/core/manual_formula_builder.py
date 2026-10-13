"""生成可直接粘贴到 Excel 的清洗公式，不读取或修改工作簿。"""

from typing import Literal

from app.core.cleaner import CleaningRules
from app.core.export_formula_builder import build_export_formula
from app.core.formula_builder import build_formula


FormulaMode = Literal["simple", "strict"]


def build_manual_formula(
    rules: CleaningRules, cell_reference: str = "A2", *, mode: FormulaMode = "simple",
) -> str:
    """默认生成简洁公式；严格版复用导出逻辑，两种模式均保留非文本类型。"""
    if mode not in ("simple", "strict"):
        raise ValueError("公式模式仅支持 simple（简洁版）或 strict（严格版）。")
    if rules.has_type_rules:
        raise ValueError(
            "全角转换、文本数字转换和日期标准化暂不支持与安全清洗完全一致的独立公式。"
            "请取消这些高级规则后生成字符清洗公式，或使用现有清洗预览。"
        )
    reference = build_formula(CleaningRules(), cell_reference.strip())[1:]
    if mode == "simple":
        # 简洁版使用 TRIM；仅勾选去首尾空格时，才会同时合并中间连续空格。
        expression = build_formula(rules, reference)[1:]
        return f'=IF(ISBLANK({reference}),"",IF(ISTEXT({reference}),{expression},{reference}))'
    formula = build_export_formula((rules,), reference)
    # Excel 公式栏使用 UNICHAR，xlsx 内部的 _xlfn. 前缀不需要复制给用户。
    return formula.replace("_xlfn.UNICHAR(", "UNICHAR(")
