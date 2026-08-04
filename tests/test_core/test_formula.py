"""测试公式生成器能力"""

import pytest

from abacus.core.formula import FormulaGeneratorCapability


class TestFormulaGenerator:
    def test_vlookup(self):
        """生成 VLOOKUP 公式"""
        cap = FormulaGeneratorCapability()
        result = cap.execute(None, formula_type="vlookup",
                           params={"lookup_value": "D2", "table_range": "A:B",
                                   "col_index": 2})
        assert "VLOOKUP" in result["formula"]
        assert "D2" in result["formula"]
    
    def test_sumifs(self):
        """生成 SUMIFS 公式"""
        cap = FormulaGeneratorCapability()
        result = cap.execute(None, formula_type="sumifs",
                           params={"sum_range": "C:C", "criteria_range1": "A:A",
                                   "criteria1": "\"North\"", "criteria_range2": "B:B",
                                   "criteria2": "\">100\""})
        assert "SUMIFS" in result["formula"]
    
    def test_if(self):
        """生成 IF 公式"""
        cap = FormulaGeneratorCapability()
        result = cap.execute(None, formula_type="if",
                           params={"condition": "A1>100", "value_if_true": "\"Yes\"",
                                   "value_if_false": "\"No\""})
        assert "IF" in result["formula"]
    
    def test_today(self):
        """生成 TODAY 公式"""
        cap = FormulaGeneratorCapability()
        result = cap.execute(None, formula_type="today", params={})
        assert result["formula"] == "=TODAY()"
    
    def test_unknown_formula(self):
        """未知公式类型"""
        cap = FormulaGeneratorCapability()
        with pytest.raises(Exception):
            cap.execute(None, formula_type="unknown", params={})
