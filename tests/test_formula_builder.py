"""验证参考公式的规则组合、固定顺序和独立生成。"""

from unittest.mock import patch

import pytest

from app.core.cleaner import CleaningRules
from app.core.formula_builder import FORMULA_NOTE, build_formula


def test_nbsp_and_outer_spaces_match_requested_example():
    assert build_formula(CleaningRules(nbsp=True, outer_spaces=True)) == '=TRIM(SUBSTITUTE(A2,UNICHAR(160)," "))'


def test_all_rules_have_the_same_execution_order_as_python():
    rules = CleaningRules(nbsp=True, tab=True, newline=True, carriage_return=True, outer_spaces=True)
    assert build_formula(rules, "B2") == (
        '=TRIM(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(B2,UNICHAR(160)," "),'
        'CHAR(9)," "),CHAR(10)," "),CHAR(13)," "))'
    )


def test_delete_strategy_is_reflected_only_for_selected_hidden_characters():
    rules = CleaningRules(nbsp=True, tab=True, hidden_strategy="delete")
    assert build_formula(rules) == '=SUBSTITUTE(SUBSTITUTE(A2,UNICHAR(160)," "),CHAR(9),"")'


@pytest.mark.parametrize("reference", ["A2", "AA2", "$B$2", "XFD1048576"])
def test_reference_is_used_without_a_hardcoded_column(reference):
    assert build_formula(CleaningRules(), reference) == "=" + reference


@pytest.mark.parametrize("reference", ["A0", "A1048577", "XFE2", "AAAA2", "A2)", "Sheet1!A2"])
def test_invalid_references_are_rejected(reference):
    with pytest.raises(ValueError, match="单元格引用"):
        build_formula(CleaningRules(), reference)


def test_formula_generation_does_not_call_cleaning_logic():
    with patch("app.core.cleaner.apply_cleaning_rules", side_effect=AssertionError("不应执行清洗")):
        assert build_formula(CleaningRules(outer_spaces=True)) == "=TRIM(A2)"
    assert "压缩中间连续普通空格" in FORMULA_NOTE
