"""测试变化检测报告"""

import pytest
import tempfile
import os
from pathlib import Path
from openpyxl import Workbook

from abacus.core.work.diff_report import DiffReportCapability


def _make_excel(rows, columns=None):
    """创建 Excel 文件"""
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    if columns:
        ws.append(columns)
    for row in rows:
        ws.append(row)
    tmp = tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False)
    wb.save(tmp.name)
    tmp.close()
    return tmp.name


class TestDiffReport:
    def test_capability_properties(self):
        cap = DiffReportCapability()
        assert cap.name == "generate_diff_report"
        assert cap.chapter == "work"
        assert "变化" in cap.description

    def test_basic_diff_same_data(self):
        """相同数据的对比"""
        cap = DiffReportCapability()
        data = [["A", 100], ["B", 200], ["C", 300]]
        old_file = _make_excel(data, ["Name", "Sales"])
        new_file = _make_excel(data, ["Name", "Sales"])

        result = cap.execute(None, old_file=old_file, old_sheet="Sheet1",
                           new_file=new_file, new_sheet="Sheet1")
        assert result["success"] is True
        assert result["old_rows"] == 3
        assert result["new_rows"] == 3
        assert result["row_diff"] == 0
        assert result["added_columns"] == []
        assert result["removed_columns"] == []

        os.unlink(old_file)
        os.unlink(new_file)

    def test_diff_with_added_rows(self):
        """新增行的对比"""
        cap = DiffReportCapability()
        old_data = [["A", 100], ["B", 200]]
        new_data = [["A", 100], ["B", 200], ["C", 300]]
        old_file = _make_excel(old_data, ["Name", "Sales"])
        new_file = _make_excel(new_data, ["Name", "Sales"])

        result = cap.execute(None, old_file=old_file, old_sheet="Sheet1",
                           new_file=new_file, new_sheet="Sheet1")
        assert result["row_diff"] == 1

        os.unlink(old_file)
        os.unlink(new_file)

    def test_diff_with_removed_rows(self):
        """删除行的对比"""
        cap = DiffReportCapability()
        old_data = [["A", 100], ["B", 200], ["C", 300]]
        new_data = [["A", 100]]
        old_file = _make_excel(old_data, ["Name", "Sales"])
        new_file = _make_excel(new_data, ["Name", "Sales"])

        result = cap.execute(None, old_file=old_file, old_sheet="Sheet1",
                           new_file=new_file, new_sheet="Sheet1")
        assert result["row_diff"] == -2

        os.unlink(old_file)
        os.unlink(new_file)

    def test_diff_with_added_columns(self):
        """新增列的对比"""
        cap = DiffReportCapability()
        old_data = [["A", 100]]
        new_data = [["A", 100, "Extra"]]
        old_file = _make_excel(old_data, ["Name", "Sales"])
        new_file = _make_excel(new_data, ["Name", "Sales", "Extra"])

        result = cap.execute(None, old_file=old_file, old_sheet="Sheet1",
                           new_file=new_file, new_sheet="Sheet1")
        assert "Extra" in result["added_columns"]

        os.unlink(old_file)
        os.unlink(new_file)

    def test_diff_with_removed_columns(self):
        """删除列的对比"""
        cap = DiffReportCapability()
        old_data = [["A", 100, "Extra"]]
        new_data = [["A", 100]]
        old_file = _make_excel(old_data, ["Name", "Sales", "Extra"])
        new_file = _make_excel(new_data, ["Name", "Sales"])

        result = cap.execute(None, old_file=old_file, old_sheet="Sheet1",
                           new_file=new_file, new_sheet="Sheet1")
        assert "Extra" in result["removed_columns"]

        os.unlink(old_file)
        os.unlink(new_file)

    def test_diff_with_key_columns(self):
        """使用键列的详细对比"""
        cap = DiffReportCapability()
        old_data = [["A", 100], ["B", 200]]
        new_data = [["A", 150], ["B", 200], ["C", 300]]
        old_file = _make_excel(old_data, ["Name", "Sales"])
        new_file = _make_excel(new_data, ["Name", "Sales"])

        result = cap.execute(None, old_file=old_file, old_sheet="Sheet1",
                           new_file=new_file, new_sheet="Sheet1",
                           key_columns=["Name"])
        assert result["success"] is True
        assert len(result["changes"]) > 0

        os.unlink(old_file)
        os.unlink(new_file)

    def test_diff_with_multiple_key_columns(self):
        """使用多个键列的详细对比"""
        cap = DiffReportCapability()
        old_data = [["A", "East", 100], ["B", "West", 200]]
        new_data = [["A", "East", 150], ["B", "East", 200], ["C", "North", 300]]
        old_file = _make_excel(old_data, ["Name", "Region", "Sales"])
        new_file = _make_excel(new_data, ["Name", "Region", "Sales"])

        result = cap.execute(None, old_file=old_file, old_sheet="Sheet1",
                           new_file=new_file, new_sheet="Sheet1",
                           key_columns=["Name", "Region"])
        assert result["success"] is True

        os.unlink(old_file)
        os.unlink(new_file)

    def test_missing_old_file(self):
        cap = DiffReportCapability()
        with pytest.raises(Exception):
            cap.execute(None, new_file="new.xlsx", new_sheet="Sheet1")

    def test_missing_new_file(self):
        cap = DiffReportCapability()
        with pytest.raises(Exception):
            cap.execute(None, old_file="old.xlsx", old_sheet="Sheet1")

    def test_old_file_not_found(self):
        cap = DiffReportCapability()
        with pytest.raises(Exception):
            cap.execute(None, old_file="nonexistent_old.xlsx", old_sheet="Sheet1",
                       new_file="nonexistent_new.xlsx", new_sheet="Sheet1")

    def test_new_file_not_found(self):
        cap = DiffReportCapability()
        old_data = [["A", 100]]
        old_file = _make_excel(old_data, ["Name", "Sales"])
        with pytest.raises(Exception):
            cap.execute(None, old_file=old_file, old_sheet="Sheet1",
                       new_file="nonexistent_new.xlsx", new_sheet="Sheet1")
        os.unlink(old_file)

    def test_invalid_key_column(self):
        """键列不存在"""
        cap = DiffReportCapability()
        old_data = [["A", 100]]
        new_data = [["A", 100]]
        old_file = _make_excel(old_data, ["Name", "Sales"])
        new_file = _make_excel(new_data, ["Name", "Sales"])
        with pytest.raises(Exception):
            cap.execute(None, old_file=old_file, old_sheet="Sheet1",
                       new_file=new_file, new_sheet="Sheet1",
                       key_columns=["Nonexistent"])
        os.unlink(old_file)
        os.unlink(new_file)

    def test_diff_column_count_change(self):
        """列数变化的对比"""
        cap = DiffReportCapability()
        old_data = [["A", 100]]
        new_data = [["A", 100, "Extra"]]
        old_file = _make_excel(old_data, ["Name", "Sales"])
        new_file = _make_excel(new_data, ["Name", "Sales", "Extra"])

        result = cap.execute(None, old_file=old_file, old_sheet="Sheet1",
                           new_file=new_file, new_sheet="Sheet1")
        assert len(result["changes"]) > 0
        assert result["changes"][0]["type"] == "column_count"

        os.unlink(old_file)
        os.unlink(new_file)

    def test_schema(self):
        cap = DiffReportCapability()
        schema = cap.schema
        assert len(schema) > 0
        names = [s.name for s in schema]
        assert "old_file" in names
        assert "old_sheet" in names
        assert "new_file" in names
        assert "new_sheet" in names
        assert "key_columns" in names
