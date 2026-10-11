"""仅生成 Excel 学习参考公式，不执行公式或调用 Python 清洗函数。"""

import re

from openpyxl.utils.cell import column_index_from_string

from app.core.cleaner import CleaningRules, RULE_ORDER


FORMULA_NOTE = (
    "公式仅供文本处理参考，不参与清洗。Excel TRIM 会压缩中间连续普通空格，"
    "Python 仅去首尾并保留中间空格；实际结果以清洗预览为准。"
)


def build_formula(rules: CleaningRules, cell_reference: str = "A2") -> str:
    match = re.fullmatch(r"(\$?)([A-Za-z]{1,3})(\$?)([1-9][0-9]*)", cell_reference)
    if (
        match is None
        or column_index_from_string(match[2]) > 16384
        or int(match[4]) > 1048576
    ):
        raise ValueError("请输入有效的 Excel 单元格引用，例如 A2 或 $B$2。")
    expression = cell_reference.upper()
    if rules.has_type_rules:
        return "本次包含高级类型处理，以清洗预览为准；新增对照列模式将保存确认后的值。"
    character_codes = {"nbsp": 160, "tab": 9, "newline": 10, "carriage_return": 13}
    for rule in RULE_ORDER:
        if not getattr(rules, rule):
            continue
        if rule == "outer_spaces":
            expression = f"TRIM({expression})"
        else:
            replacement = " " if rule == "nbsp" or rules.hidden_strategy == "space" else ""
            character = "UNICHAR(160)" if rule == "nbsp" else f"CHAR({character_codes[rule]})"
            expression = f'SUBSTITUTE({expression},{character},"{replacement}")'
    return "=" + expression
