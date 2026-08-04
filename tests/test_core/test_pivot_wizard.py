"""测试透视表向导"""

import pytest
import tempfile
import os
from pathlib import Path
from openpyxl import Workbook

from abacus.core.pivot.wizard import PivotWizardCapability


@pytest.fixture
def sample_excel():
    """创建示例 Excel 文件"""
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(["ID", "Name", "Region", "Sales", "Profit"])
    ws.append([1, "Product A", "East", 1000, 200])
    ws.append([2, "Product B", "West", 1500, 300])
    ws.append([3, "Product A", "East", 800, 160])
    ws.append([4, "Product C", "North", 2000, 400])
    ws.append([5, "Product B", "West", 1200, 240])
    ws.append([6, "Product A", "East", 900, 180])
    tmp = tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False)
    wb.save(tmp.name)
    tmp.close()
    return tmp.name


@pytest.fixture
def output_file():
    """创建临时输出路径（非文件）"""
    tmp = tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False)
    tmp.close()
    os.unlink(tmp.name)
    return tmp.name


class TestPivotWizard:
    def test_capability_properties(self):
        cap = PivotWizardCapability()
        assert cap.name == "pivot_wizard"
        assert cap.chapter == "share"
        assert "透视表向导" in cap.description

    def test_create_pivot_sum(self, sample_excel, output_file):
        cap = PivotWizardCapability()
        result = cap.execute(None, file=sample_excel, sheet="Sheet1",
                           range="A1:E7", rows=["Name"], values=["Sales"],
                           agg_function="sum", output=output_file)
        assert result["file"] == sample_excel
        assert result["rows"] > 0
        assert result["agg_function"] == "sum"
        assert Path(output_file).exists()
        os.unlink(sample_excel)
        os.unlink(output_file)

    def test_create_pivot_mean(self, sample_excel, output_file):
        cap = PivotWizardCapability()
        result = cap.execute(None, file=sample_excel, sheet="Sheet1",
                           range="A1:E7", rows=["Region"], values=["Sales"],
                           agg_function="mean", output=output_file)
        assert result["agg_function"] == "mean"
        os.unlink(sample_excel)
        os.unlink(output_file)

    def test_create_pivot_count(self, sample_excel, output_file):
        cap = PivotWizardCapability()
        result = cap.execute(None, file=sample_excel, sheet="Sheet1",
                           range="A1:E7", rows=["Region"], values=["Sales"],
                           agg_function="count", output=output_file)
        assert result["agg_function"] == "count"
        os.unlink(sample_excel)
        os.unlink(output_file)

    def test_create_pivot_min(self, sample_excel, output_file):
        cap = PivotWizardCapability()
        result = cap.execute(None, file=sample_excel, sheet="Sheet1",
                           range="A1:E7", rows=["Name"], values=["Sales"],
                           agg_function="min", output=output_file)
        assert result["agg_function"] == "min"
        os.unlink(sample_excel)
        os.unlink(output_file)

    def test_create_pivot_max(self, sample_excel, output_file):
        cap = PivotWizardCapability()
        result = cap.execute(None, file=sample_excel, sheet="Sheet1",
                           range="A1:E7", rows=["Name"], values=["Sales"],
                           agg_function="max", output=output_file)
        assert result["agg_function"] == "max"
        os.unlink(sample_excel)
        os.unlink(output_file)

    def test_create_pivot_multiple_values(self, sample_excel, output_file):
        cap = PivotWizardCapability()
        result = cap.execute(None, file=sample_excel, sheet="Sheet1",
                           range="A1:E7", rows=["Name"], values=["Sales", "Profit"],
                           agg_function="sum", output=output_file)
        assert len(result["value_fields"]) == 2
        os.unlink(sample_excel)
        os.unlink(output_file)

    def test_create_pivot_multiple_rows(self, sample_excel, output_file):
        cap = PivotWizardCapability()
        result = cap.execute(None, file=sample_excel, sheet="Sheet1",
                           range="A1:E7", rows=["Name", "Region"], values=["Sales"],
                           agg_function="sum", output=output_file)
        assert len(result["row_fields"]) == 2
        os.unlink(sample_excel)
        os.unlink(output_file)

    def test_create_pivot_output_sheet(self, sample_excel, output_file):
        cap = PivotWizardCapability()
        result = cap.execute(None, file=sample_excel, sheet="Sheet1",
                           range="A1:E7", rows=["Name"], values=["Sales"],
                           agg_function="sum", output=output_file, output_sheet="MyPivot")
        assert result["output_sheet"] == "MyPivot"
        os.unlink(sample_excel)
        os.unlink(output_file)

    def test_create_pivot_default_output(self, sample_excel):
        """测试默认输出到源文件"""
        cap = PivotWizardCapability()
        result = cap.execute(None, file=sample_excel, sheet="Sheet1",
                           range="A1:E7", rows=["Name"], values=["Sales"],
                           agg_function="sum")
        assert result["output"] == sample_excel
        os.unlink(sample_excel)

    def test_missing_file(self):
        cap = PivotWizardCapability()
        with pytest.raises(Exception):
            cap.execute(None, sheet="Sheet1", range="A1:E7",
                       rows=["Name"], values=["Sales"])

    def test_missing_rows(self, sample_excel):
        cap = PivotWizardCapability()
        with pytest.raises(Exception):
            cap.execute(None, file=sample_excel, sheet="Sheet1",
                       range="A1:E7", values=["Sales"])
        os.unlink(sample_excel)

    def test_missing_values(self, sample_excel):
        cap = PivotWizardCapability()
        with pytest.raises(Exception):
            cap.execute(None, file=sample_excel, sheet="Sheet1",
                       range="A1:E7", rows=["Name"])
        os.unlink(sample_excel)

    def test_file_not_found(self):
        cap = PivotWizardCapability()
        with pytest.raises(Exception):
            cap.execute(None, file="nonexistent.xlsx", sheet="Sheet1",
                       range="A1:E7", rows=["Name"], values=["Sales"])

    def test_invalid_field(self, sample_excel):
        cap = PivotWizardCapability()
        with pytest.raises(Exception):
            cap.execute(None, file=sample_excel, sheet="Sheet1",
                       range="A1:E7", rows=["Nonexistent"], values=["Sales"])
        os.unlink(sample_excel)

    def test_schema(self):
        cap = PivotWizardCapability()
        schema = cap.schema
        assert len(schema) > 0
        names = [s.name for s in schema]
        assert "file" in names
        assert "sheet" in names
        assert "rows" in names
        assert "values" in names
