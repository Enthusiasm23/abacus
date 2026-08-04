"""测试 Sheet 扩展能力"""

import pytest
from pathlib import Path
from openpyxl import Workbook

from abacus.core.sheet_ext.style import SheetStyleCapability
from abacus.core.sheet_ext.visibility import SheetVisibilityCapability


@pytest.fixture
def sample_excel(tmp_path):
    """创建测试用 Excel 文件（多个工作表）"""
    file_path = tmp_path / "test.xlsx"
    wb = Workbook()
    ws1 = wb.active
    ws1.title = "Sheet1"
    ws1["A1"] = "Test1"
    ws2 = wb.create_sheet("Sheet2")
    ws2["A1"] = "Test2"
    wb.save(file_path)
    wb.close()
    return file_path


class TestSheetStyle:
    def test_set_tab_color(self, sample_excel):
        """设置标签颜色"""
        cap = SheetStyleCapability()
        result = cap.execute(None, file=str(sample_excel), sheet="Sheet1",
                           action="set", color="FF0000")
        assert result["action"] == "set"
        assert result["color"] == "FF0000"
    
    def test_get_tab_color(self, sample_excel):
        """获取标签颜色"""
        cap = SheetStyleCapability()
        # 先设置颜色
        cap.execute(None, file=str(sample_excel), sheet="Sheet1",
                   action="set", color="FF0000")
        # 获取颜色
        result = cap.execute(None, file=str(sample_excel), sheet="Sheet1",
                           action="get")
        assert result["action"] == "get"
    
    def test_clear_tab_color(self, sample_excel):
        """清除标签颜色"""
        cap = SheetStyleCapability()
        # 先设置颜色
        cap.execute(None, file=str(sample_excel), sheet="Sheet1",
                   action="set", color="FF0000")
        # 清除颜色
        result = cap.execute(None, file=str(sample_excel), sheet="Sheet1",
                           action="clear")
        assert result["action"] == "clear"


class TestSheetVisibility:
    def test_hide_sheet(self, sample_excel):
        """隐藏工作表"""
        cap = SheetVisibilityCapability()
        result = cap.execute(None, file=str(sample_excel), sheet="Sheet1",
                           action="hide")
        assert result["action"] == "hide"
    
    def test_show_sheet(self, sample_excel):
        """显示工作表"""
        cap = SheetVisibilityCapability()
        # 先隐藏
        cap.execute(None, file=str(sample_excel), sheet="Sheet1",
                   action="hide")
        # 显示
        result = cap.execute(None, file=str(sample_excel), sheet="Sheet1",
                           action="show")
        assert result["action"] == "show"
    
    def test_very_hide_sheet(self, sample_excel):
        """非常隐藏工作表"""
        cap = SheetVisibilityCapability()
        result = cap.execute(None, file=str(sample_excel), sheet="Sheet1",
                           action="very_hide")
        assert result["action"] == "very_hide"
    
    def test_get_visibility(self, sample_excel):
        """获取可见性状态"""
        cap = SheetVisibilityCapability()
        result = cap.execute(None, file=str(sample_excel), sheet="Sheet1",
                           action="get")
        assert result["action"] == "get"
        assert result["state"] == "visible"
