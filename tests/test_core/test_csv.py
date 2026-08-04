"""测试 CSV 处理能力"""

import pytest
import pandas as pd
from pathlib import Path

from abacus.core.csv import CSVMergeCapability, CSVVisualizeCapability


@pytest.fixture
def sample_csvs(tmp_path):
    """创建测试用 CSV 文件"""
    df1 = pd.DataFrame({"Name": ["Alice", "Bob"], "Sales": [100, 200]})
    df2 = pd.DataFrame({"Name": ["Charlie", "David"], "Sales": [150, 250]})
    
    file1 = tmp_path / "test1.csv"
    file2 = tmp_path / "test2.csv"
    
    df1.to_csv(file1, index=False)
    df2.to_csv(file2, index=False)
    
    return [str(file1), str(file2)]


class TestCSVMerge:
    def test_concat_merge(self, sample_csvs, tmp_path):
        """纵向合并"""
        cap = CSVMergeCapability()
        output = tmp_path / "merged.xlsx"
        result = cap.execute(None, files=sample_csvs, output=str(output),
                           merge_type="concat")
        assert result["files_merged"] == 2
        assert result["total_rows"] == 4
    
    def test_merge_with_dedup(self, sample_csvs, tmp_path):
        """合并并去重"""
        cap = CSVMergeCapability()
        output = tmp_path / "merged.xlsx"
        result = cap.execute(None, files=sample_csvs, output=str(output),
                           merge_type="concat", dedup=True)
        assert result["duplicates_removed"] >= 0
    
    def test_merge_to_csv(self, sample_csvs, tmp_path):
        """合并为 CSV"""
        cap = CSVMergeCapability()
        output = tmp_path / "merged.csv"
        result = cap.execute(None, files=sample_csvs, output=str(output),
                           merge_type="concat")
        assert output.exists()


class TestCSVVisualize:
    def test_create_bar_chart(self, tmp_path):
        """创建柱状图"""
        df = pd.DataFrame({"Name": ["Alice", "Bob", "Charlie"], "Sales": [100, 200, 150]})
        file = tmp_path / "test.csv"
        df.to_csv(file, index=False)
        
        cap = CSVVisualizeCapability()
        output = tmp_path / "visual.xlsx"
        result = cap.execute(None, file=str(file), output=str(output),
                           chart_type="bar")
        assert result["sheets"] >= 2
        assert output.exists()
    
    def test_create_dashboard(self, tmp_path):
        """创建仪表板"""
        df = pd.DataFrame({"Name": ["Alice", "Bob"], "Sales": [100, 200], "Profit": [20, 40]})
        file = tmp_path / "test.csv"
        df.to_csv(file, index=False)
        
        cap = CSVVisualizeCapability()
        output = tmp_path / "visual.xlsx"
        result = cap.execute(None, file=str(file), output=str(output),
                           include_dashboard=True, include_stats=True)
        assert result["sheets"] >= 3
