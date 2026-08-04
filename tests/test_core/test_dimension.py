"""测试少广章（反向计算）能力"""

import pytest

from abacus.core.dimension import FindDimensionCapability, DeriveCapability, CalculateCapability


class TestFindDimension:
    def test_rectangle(self):
        """矩形：已知面积和一边求另一边"""
        cap = FindDimensionCapability()
        result = cap.execute(None, area=100, shape="rectangle", known_side=10)
        assert result["side2"] == 10.0
    
    def test_circle(self):
        """圆形：已知面积求半径"""
        cap = FindDimensionCapability()
        result = cap.execute(None, area=314.159, shape="circle")
        assert abs(result["radius"] - 10) < 0.01
    
    def test_cube(self):
        """立方体：已知体积求边长"""
        cap = FindDimensionCapability()
        result = cap.execute(None, area=1000, shape="cube")
        assert abs(result["side"] - 10) < 0.01


class TestDerive:
    def test_simple_interest(self):
        """单利：已知利息求本金"""
        cap = DeriveCapability()
        result = cap.execute(None, target_value=100, formula="simple_interest",
                           params={"r": 0.05, "t": 2})
        assert result["derived"] == "P"
    
    def test_profit_margin(self):
        """利润率：已知利润率和成本求价格"""
        cap = DeriveCapability()
        result = cap.execute(None, target_value=0.2, formula="profit_margin",
                           params={"cost": 80})
        assert result["derived"] == "price"


class TestCalculate:
    def test_simple_expression(self):
        """简单表达式"""
        cap = CalculateCapability()
        result = cap.execute(None, expression="2 + 3 * 4")
        assert result["result"] == 14
    
    def test_with_variables(self):
        """带变量的表达式"""
        cap = CalculateCapability()
        result = cap.execute(None, expression="x + y", variables={"x": 10, "y": 20})
        assert result["result"] == 30
    
    def test_math_functions(self):
        """数学函数"""
        cap = CalculateCapability()
        result = cap.execute(None, expression="sqrt(16) + abs(-5)")
        assert result["result"] == 9
