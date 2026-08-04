"""测试格式转换能力"""

import pytest
import pandas as pd
from pathlib import Path

from abacus.core.conversion import ExcelToMarkdownCapability, SplitSheetCapability


@pytest.fixture
def sample_excel(tmp_path):
    """创建测试用 Excel 文件"""
    file_path = tmp_path / "test.xlsx"
    df = pd.DataFrame({
        "Name": ["Alice", "Bob", "Charlie", "David"],
        "Sales": [100, 200, 150, 250],
        "Region": ["North", "South", "North", "South"]
    })
    df.to_excel(file_path, index=False, sheet_name="Data")
    return file_path


class TestExcelToMarkdown:
    def test_convert_to_markdown(self, sample_excel, tmp_path):
        """转换为 Markdown"""
        cap = ExcelToMarkdownCapability()
        output = tmp_path / "output.md"
        result = cap.execute(None, file=str(sample_excel), output=str(output))
        assert result["sheets_converted"] == 1
        assert output.exists()
    
    def test_convert_specific_sheet(self, sample_excel, tmp_path):
        """转换指定工作表"""
        cap = ExcelToMarkdownCapability()
        output = tmp_path / "output.md"
        result = cap.execute(None, file=str(sample_excel), sheet="Data", output=str(output))
        assert result["sheets_converted"] == 1


class TestSplitSheet:
    def test_split_by_row_count(self, sample_excel, tmp_path):
        """按行数拆分"""
        cap = SplitSheetCapability()
        output_dir = tmp_path / "split"
        result = cap.execute(None, file=str(sample_excel), sheet="Data",
                           output_dir=str(output_dir), split_by="row_count",
                           row_count=2, prefix="data")
        assert result["files_created"] >= 1
        assert output_dir.exists()
    
    def test_split_by_column(self, sample_excel, tmp_path):
        """按列值拆分"""
        cap = SplitSheetCapability()
        output_dir = tmp_path / "split"
        result = cap.execute(None, file=str(sample_excel), sheet="Data",
                           output_dir=str(output_dir), split_by="column",
                           split_column="Region", prefix="data")
        assert result["files_created"] == 2  # North, South
        assert output_dir.exists()
