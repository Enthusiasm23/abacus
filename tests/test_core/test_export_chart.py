"""测试导出图表为图片"""

import pytest
import tempfile
import os
from pathlib import Path
from openpyxl import Workbook

from abacus.core.work.export_chart import ExportChartAsImageCapability


@pytest.fixture
def sample_excel():
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(["Month", "Sales"])
    ws.append(["Jan", 100])
    ws.append(["Feb", 150])
    ws.append(["Mar", 200])
    tmp = tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False)
    wb.save(tmp.name)
    tmp.close()
    return tmp.name


class TestExportChartAsImage:
    def test_capability_properties(self):
        cap = ExportChartAsImageCapability()
        assert cap.name == "export_chart_as_image"
        assert cap.chapter == "work"
        assert "图表" in cap.description

    def test_missing_file(self):
        cap = ExportChartAsImageCapability()
        with pytest.raises(Exception):
            cap.execute(None, sheet="Sheet1", chart_index=0, output="output.png")

    def test_missing_sheet(self, sample_excel):
        cap = ExportChartAsImageCapability()
        with pytest.raises(Exception):
            cap.execute(None, file=sample_excel, chart_index=0, output="output.png")
        os.unlink(sample_excel)

    def test_missing_chart_index(self, sample_excel):
        cap = ExportChartAsImageCapability()
        with pytest.raises(Exception):
            cap.execute(None, file=sample_excel, sheet="Sheet1", output="output.png")
        os.unlink(sample_excel)

    def test_missing_output(self, sample_excel):
        cap = ExportChartAsImageCapability()
        with pytest.raises(Exception):
            cap.execute(None, file=sample_excel, sheet="Sheet1", chart_index=0)
        os.unlink(sample_excel)

    def test_file_not_found(self):
        cap = ExportChartAsImageCapability()
        with pytest.raises(Exception):
            cap.execute(None, file="nonexistent.xlsx", sheet="Sheet1",
                       chart_index=0, output="output.png")

    def test_sheet_not_found(self, sample_excel):
        cap = ExportChartAsImageCapability()
        with pytest.raises(Exception):
            cap.execute(None, file=sample_excel, sheet="Nonexistent",
                       chart_index=0, output="output.png")
        os.unlink(sample_excel)

    def test_chart_not_found(self, sample_excel):
        cap = ExportChartAsImageCapability()
        with pytest.raises(Exception):
            cap.execute(None, file=sample_excel, sheet="Sheet1",
                       chart_index=0, output="output.png")
        os.unlink(sample_excel)

    def test_schema(self):
        cap = ExportChartAsImageCapability()
        schema = cap.schema
        assert len(schema) == 4
        names = [s.name for s in schema]
        assert "file" in names
        assert "sheet" in names
        assert "chart_index" in names
        assert "output" in names
