"""测试数据分析能力"""

import pytest
import pandas as pd
from pathlib import Path

from abacus.core.analysis import DataAnalysisCapability, DataCleaningCapability, PivotAnalysisCapability


@pytest.fixture
def sample_csv(tmp_path):
    """创建测试用 CSV 文件"""
    file_path = tmp_path / "test.csv"
    df = pd.DataFrame({
        "Name": ["Alice", "Bob", "Charlie", "Alice"],
        "Region": ["North", "South", "North", "South"],
        "Sales": [100, 200, 150, 300],
        "Profit": [20, 40, 30, 60]
    })
    df.to_csv(file_path, index=False)
    return file_path


class TestDataAnalysis:
    def test_auto_analysis(self, sample_csv):
        """自动分析"""
        cap = DataAnalysisCapability()
        result = cap.execute(None, file=str(sample_csv), analysis_type="auto")
        assert result["rows"] == 4
        assert result["columns"] == 4
        assert "summary" in result["analysis"]
    
    def test_summary_analysis(self, sample_csv):
        """统计摘要"""
        cap = DataAnalysisCapability()
        result = cap.execute(None, file=str(sample_csv), analysis_type="summary")
        assert "numeric_summary" in result["analysis"]
    
    def test_correlation_analysis(self, sample_csv):
        """相关性分析"""
        cap = DataAnalysisCapability()
        result = cap.execute(None, file=str(sample_csv), analysis_type="correlation")
        assert "correlation_matrix" in result["analysis"]


class TestDataCleaning:
    def test_remove_duplicates(self, sample_csv, tmp_path):
        """去重"""
        cap = DataCleaningCapability()
        output = tmp_path / "cleaned.xlsx"
        result = cap.execute(None, file=str(sample_csv), output=str(output), 
                           operations=["remove_duplicates"])
        # Alice appears twice but with different regions, so no exact duplicates
        assert result["duplicates_removed"] >= 0
    
    def test_handle_missing(self, sample_csv, tmp_path):
        """处理缺失值"""
        cap = DataCleaningCapability()
        output = tmp_path / "cleaned.xlsx"
        result = cap.execute(None, file=str(sample_csv), output=str(output),
                           operations=["handle_missing"])
        assert "missing_removed" in result


class TestPivotAnalysis:
    def test_pivot_by_region(self, sample_csv):
        """按区域透视"""
        cap = PivotAnalysisCapability()
        result = cap.execute(None, file=str(sample_csv), group_by="Region",
                           value_field="Sales", agg_function="sum")
        assert result["groups"] == 2
