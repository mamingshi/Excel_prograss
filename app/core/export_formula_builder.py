"""生成实际导出的 Excel 清洗公式，保留中间普通空格与非文本类型。"""

from openpyxl.formula import Tokenizer

from app.core.cleaner import CleaningRules, RULE_ORDER
from app.core.formula_builder import build_formula


def _trim_outer_spaces(expression: str) -> str:
    """TRIM 只用于定位边界，MID 从原文本取片段，不压缩中间空格。"""
    trimmed = f"TRIM({expression})"
    start = f"FIND(LEFT({trimmed},1),{expression})"
    last = f"RIGHT({trimmed},1)"
    count = f'(LEN({expression})-LEN(SUBSTITUTE({expression},{last},"")))/LEN({last})'
    end = f"FIND(CHAR(1),SUBSTITUTE({expression},{last},CHAR(1),{count}))"
    # 某些 Excel 版本将 emoji 的 LEN 计为 2，按字符实际长度计算出现次数和边界。
    return f'IF({trimmed}="","",MID({expression},{start},{end}-{start}+LEN({last})))'


def build_export_formula(
    rule_sequence: tuple[CleaningRules, ...], cell_reference: str = "A2"
) -> str:
    """按实际应用顺序构造单个公式；重复字符替换和相邻 trim 可安全合并。"""
    if any(rules.has_type_rules for rules in rule_sequence):
        raise ValueError("高级类型处理使用预览确认的值，不能套用通用文本清洗公式。")
    reference = build_formula(CleaningRules(), cell_reference)[1:]
    expression = reference
    character_codes = {"nbsp": 160, "tab": 9, "newline": 10, "carriage_return": 13}
    replaced = set()
    previous = None
    for rules in rule_sequence:
        for rule in RULE_ORDER:
            if not getattr(rules, rule):
                continue
            if rule == "outer_spaces":
                if previous != rule:
                    expression = _trim_outer_spaces(expression)
                    previous = rule
            elif rule not in replaced:
                replacement = " " if rule == "nbsp" or rules.hidden_strategy == "space" else ""
                character = "_xlfn.UNICHAR(160)" if rule == "nbsp" else f"CHAR({character_codes[rule]})"
                expression = f'SUBSTITUTE({expression},{character},"{replacement}")'
                replaced.add(rule)
                previous = rule
    formula = f'=IF(ISBLANK({reference}),"",IF(ISTEXT({reference}),{expression},{reference}))'
    # UNICHAR 精确匹配 Unicode NBSP（Excel 2013+）；不依赖本机 ANSI 编码。
    # 无需 LET / 动态数组，也不要求 Ctrl+Shift+Enter。
    if len(formula) > 8192:
        raise ValueError("该列清洗步骤过于复杂，公式超过 Excel 长度限制，请选择“原列更新”。")
    depth = maximum = 0
    for token in Tokenizer(formula).items:
        if token.subtype == "OPEN":
            depth += 1
            maximum = max(maximum, depth)
        elif token.subtype == "CLOSE":
            depth -= 1
    if maximum > 64:
        raise ValueError("该列清洗步骤过于复杂，公式超过 Excel 嵌套限制，请选择“原列更新”。")
    return formula


def build_passthrough_formula(cell_reference: str) -> str:
    """原始公式单元格只引用其结果，不再执行文本清洗。"""
    reference = build_formula(CleaningRules(), cell_reference)[1:]
    return f'=IF(ISBLANK({reference}),"",{reference})'
