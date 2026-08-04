"""测试高级数据分析和数据转换能力"""

import pytest
import pandas as pd
import numpy as np
from pathlib import Path

from abacus.core.analysis.advanced import AdvancedAnalysisCapability
from abacus.core.grain.data_transform import DataTransformCapability
from abacus.core.exceptions import DataError, FileNotFoundError, ValidationError


@pytest.fixture
def regression_csv(tmp_path):
    """创建线性回归测试数据"""
    file_path = tmp_path / "regression.csv"
    np.random.seed(42)
    x = np.arange(1, 21)
    y = 2 * x + 5 + np.random.normal(0, 2, 20)
    df = pd.DataFrame({"X": x, "Y": y})
    df.to_csv(file_path, index=False)
    return file_path


@pytest.fixture
def timeseries_csv(tmp_path):
    """创建时间序列测试数据"""
    file_path = tmp_path / "timeseries.csv"
    values = [10, 12, 15, 14, 18, 20, 22, 25, 28, 30, 32, 35]
    df = pd.DataFrame({"Value": values})
    df.to_csv(file_path, index=False)
    return file_path


@pytest.fixture
def pivot_csv(tmp_path):
    """创建透视表测试数据"""
    file_path = tmp_path / "pivot.csv"
    df = pd.DataFrame({
        "Category": ["A", "B", "A", "B", "A", "B"],
        "Region": ["North", "North", "South", "South", "North", "South"],
        "Sales": [100, 200, 150, 250, 300, 350],
        "Profit": [20, 40, 30, 50, 60, 70]
    })
    df.to_csv(file_path, index=False)
    return file_path


@pytest.fixture
def wide_csv(tmp_path):
    """创建宽表测试数据（用于 melt）"""
    file_path = tmp_path / "wide.csv"
    df = pd.DataFrame({
        "ID": [1, 2, 3],
        "Name": ["Alice", "Bob", "Charlie"],
        "Math": [90, 85, 92],
        "English": [88, 92, 85],
        "Science": [95, 80, 88]
    })
    df.to_csv(file_path, index=False)
    return file_path


class TestAdvancedAnalysisCapability:
    """测试高级分析能力"""

    def test_regression_analysis(self, regression_csv):
        """线性回归分析"""
        cap = AdvancedAnalysisCapability()
        result = cap.execute(None, file=str(regression_csv), analysis_type="regression",
                             x_column="X", y_column="Y")
        assert result["type"] == "regression"
        assert "slope" in result
        assert "intercept" in result
        assert "r_squared" in result
        assert result["n"] == 20
        assert result["r_squared"] > 0.8  # 强线性关系
        assert "equation" in result

    def test_regression_missing_columns(self, regression_csv):
        """回归分析 - 列名错误"""
        cap = AdvancedAnalysisCapability()
        with pytest.raises(DataError, match="不存在"):
            cap.execute(None, file=str(regression_csv), analysis_type="regression",
                        x_column="NonExist", y_column="Y")

    def test_timeseries_analysis(self, timeseries_csv):
        """时间序列分析"""
        cap = AdvancedAnalysisCapability()
        result = cap.execute(None, file=str(timeseries_csv), analysis_type="timeseries",
                             y_column="Value")
        assert result["type"] == "timeseries"
        assert "statistics" in result
        assert "mean" in result["statistics"]
        assert "std" in result["statistics"]
        assert "trend" in result
        assert result["trend"]["direction"] == "上升"
        assert result["data_points"] == 12

    def test_timeseries_too_few_points(self, tmp_path):
        """时间序列 - 数据点不足"""
        file_path = tmp_path / "few.csv"
        pd.DataFrame({"V": [1, 2, 3]}).to_csv(file_path, index=False)
        cap = AdvancedAnalysisCapability()
        with pytest.raises(DataError, match="至少需要 4 个数据点"):
            cap.execute(None, file=str(file_path), analysis_type="timeseries", y_column="V")

    def test_forecast(self, timeseries_csv):
        """预测分析"""
        cap = AdvancedAnalysisCapability()
        result = cap.execute(None, file=str(timeseries_csv), analysis_type="forecast",
                             y_column="Value", periods=5)
        assert result["type"] == "forecast"
        assert len(result["forecast"]) == 5
        assert len(result["historical"]) == 12
        assert "trend" in result

    def test_forecast_default_periods(self, timeseries_csv):
        """预测 - 默认期数"""
        cap = AdvancedAnalysisCapability()
        result = cap.execute(None, file=str(timeseries_csv), analysis_type="forecast",
                             y_column="Value")
        assert len(result["forecast"]) == 10

    def test_unknown_analysis_type(self, regression_csv):
        """未知分析类型"""
        cap = AdvancedAnalysisCapability()
        with pytest.raises(DataError, match="不支持的分析类型"):
            cap.execute(None, file=str(regression_csv), analysis_type="unknown")

    def test_file_not_found(self):
        """文件不存在"""
        cap = AdvancedAnalysisCapability()
        with pytest.raises(FileNotFoundError):
            cap.execute(None, file="/nonexistent/file.csv", analysis_type="regression")

    def test_no_file_parameter(self):
        """缺少 file 参数"""
        cap = AdvancedAnalysisCapability()
        with pytest.raises(ValidationError, match="缺少必要参数 file"):
            cap.execute(None, analysis_type="regression")

    def test_excel_file(self, tmp_path):
        """Excel 文件支持"""
        file_path = tmp_path / "test.xlsx"
        df = pd.DataFrame({"X": [1, 2, 3, 4, 5], "Y": [2, 4, 6, 8, 10]})
        df.to_excel(file_path, index=False)
        cap = AdvancedAnalysisCapability()
        result = cap.execute(None, file=str(file_path), analysis_type="regression",
                             x_column="X", y_column="Y")
        assert result["type"] == "regression"
        assert result["r_squared"] > 0.99

    def test_schema(self):
        """Schema 定义"""
        cap = AdvancedAnalysisCapability()
        assert cap.name == "advanced_analysis"
        assert cap.chapter == "triangle"
        assert len(cap.schema) == 6


class TestDataTransformCapability:
    """测试数据转换能力"""

    def test_pivot_transform(self, pivot_csv):
        """透视表转换"""
        cap = DataTransformCapability()
        result = cap.execute(None, file=str(pivot_csv), transform_type="pivot",
                             params={"index": "Category", "values": "Sales", "aggfunc": "sum"})
        assert result["transform_type"] == "pivot"
        assert result["output_rows"] == 2  # A 和 B 两类
        assert result["input_rows"] == 6

    def test_melt_transform(self, wide_csv, tmp_path):
        """逆透视（宽表转长表）"""
        cap = DataTransformCapability()
        output = tmp_path / "melted.xlsx"
        result = cap.execute(None, file=str(wide_csv), transform_type="melt",
                             params={"id_vars": ["ID", "Name"]},
                             output=str(output))
        assert result["transform_type"] == "melt"
        assert result["output_rows"] == 9  # 3 行 × 3 科目
        assert result["output_columns"] == 4  # ID, Name, variable, value

    def test_reshape_transform(self, wide_csv, tmp_path):
        """重塑数据"""
        cap = DataTransformCapability()
        output = tmp_path / "reshaped.csv"
        result = cap.execute(None, file=str(wide_csv), transform_type="reshape",
                             params={"pivot_column": "Name", "value_column": "Math"},
                             output=str(output))
        assert result["transform_type"] == "reshape"

    def test_merge_transform(self, pivot_csv, tmp_path):
        """合并数据"""
        # 创建第二个文件
        other_file = tmp_path / "other.csv"
        pd.DataFrame({
            "Category": ["A", "B"],
            "Rating": [4.5, 3.8]
        }).to_csv(other_file, index=False)

        cap = DataTransformCapability()
        result = cap.execute(None, file=str(pivot_csv), transform_type="merge",
                             params={"other_file": str(other_file), "on": "Category"})
        assert result["transform_type"] == "merge"
        assert result["output_columns"] > 4  # 原始列 + 合并列

    def test_unknown_transform_type(self, pivot_csv):
        """未知转换类型"""
        cap = DataTransformCapability()
        with pytest.raises(DataError, match="Unknown transform type"):
            cap.execute(None, file=str(pivot_csv), transform_type="unknown")

    def test_file_not_found(self):
        """文件不存在"""
        cap = DataTransformCapability()
        with pytest.raises(FileNotFoundError):
            cap.execute(None, file="/nonexistent/file.csv", transform_type="pivot")

    def test_no_file_parameter(self):
        """缺少 file 参数"""
        cap = DataTransformCapability()
        with pytest.raises(DataError, match="file parameter is required"):
            cap.execute(None, transform_type="pivot")

    def test_pivot_missing_params(self, pivot_csv):
        """透视 - 缺少必要参数"""
        cap = DataTransformCapability()
        with pytest.raises(DataError, match="index and values required"):
            cap.execute(None, file=str(pivot_csv), transform_type="pivot",
                        params={"aggfunc": "sum"})

    def test_merge_missing_other_file(self, pivot_csv):
        """合并 - 缺少 other_file"""
        cap = DataTransformCapability()
        with pytest.raises(DataError, match="other_file required"):
            cap.execute(None, file=str(pivot_csv), transform_type="merge",
                        params={"on": "Category"})

    def test_schema(self):
        """Schema 定义"""
        cap = DataTransformCapability()
        assert cap.name == "transform_data"
        assert cap.chapter == "grain"
        assert len(cap.schema) == 5

    def test_excel_input(self, tmp_path):
        """Excel 文件输入"""
        file_path = tmp_path / "test.xlsx"
        df = pd.DataFrame({
            "Cat": ["X", "Y", "X", "Y"],
            "Val": [10, 20, 30, 40]
        })
        df.to_excel(file_path, index=False)
        cap = DataTransformCapability()
        result = cap.execute(None, file=str(file_path), transform_type="pivot",
                             params={"index": "Cat", "values": "Val", "aggfunc": "sum"})
        assert result["output_rows"] == 2
