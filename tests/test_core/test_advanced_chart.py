"""测试高级图表能力"""

import pytest
from pathlib import Path

from abacus.core.work.advanced_chart import AdvancedChartCapability


class TestAdvancedChart:
    def test_create_combo_chart(self, tmp_path):
        """创建组合图"""
        cap = AdvancedChartCapability()
        output = tmp_path / "chart.xlsx"
        data = {
            "headers": ["Month", "Sales", "Profit"],
            "rows": [
                ["Jan", 100, 20],
                ["Feb", 150, 30],
                ["Mar", 200, 40]
            ]
        }
        result = cap.execute(None, file=str(output), data=data,
                           chart_type="combo", title="Sales Chart")
        assert result["created"] is True
        assert result["chart_type"] == "combo"
        assert output.exists()

    def test_create_waterfall_chart(self, tmp_path):
        """创建瀑布图"""
        cap = AdvancedChartCapability()
        output = tmp_path / "waterfall.xlsx"
        data = {
            "headers": ["Item", "Amount"],
            "rows": [
                ["Revenue", 1000],
                ["Cost", -500],
                ["Profit", 500]
            ]
        }
        result = cap.execute(None, file=str(output), data=data,
                           chart_type="waterfall", title="Waterfall")
        assert result["created"] is True
        assert result["chart_type"] == "waterfall"
        assert output.exists()

    def test_create_gantt_chart(self, tmp_path):
        """创建甘特图"""
        cap = AdvancedChartCapability()
        output = tmp_path / "gantt.xlsx"
        data = {
            "headers": ["Task", "Start", "Duration"],
            "rows": [
                ["Task A", 0, 5],
                ["Task B", 3, 7],
                ["Task C", 5, 4]
            ]
        }
        result = cap.execute(None, file=str(output), data=data,
                           chart_type="gantt", title="Gantt Chart")
        assert result["created"] is True
        assert result["chart_type"] == "gantt"
        assert output.exists()

    def test_create_dual_axis_chart(self, tmp_path):
        """创建双轴图"""
        cap = AdvancedChartCapability()
        output = tmp_path / "dual.xlsx"
        data = {
            "headers": ["Month", "Sales", "Rate"],
            "rows": [
                ["Jan", 100, 0.1],
                ["Feb", 150, 0.15],
                ["Mar", 200, 0.2]
            ]
        }
        result = cap.execute(None, file=str(output), data=data,
                           chart_type="dual_axis", title="Dual Axis")
        assert result["created"] is True
        assert result["chart_type"] == "dual_axis"
        assert output.exists()
