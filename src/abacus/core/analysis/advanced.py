"""勾股章 - 高级数据分析：用 pandas/numpy 实现"""

import logging
from pathlib import Path
from typing import Any

import numpy as np
import pandas as pd

from ..base import Capability, CapabilitySchema
from ..exceptions import DataError, FileNotFoundError, ValidationError

logger = logging.getLogger(__name__)


class AdvancedAnalysisCapability(Capability):
    """高级数据分析：回归分析、时间序列、预测"""

    @property
    def name(self) -> str:
        return "advanced_analysis"

    @property
    def chapter(self) -> str:
        return "triangle"

    @property
    def description(self) -> str:
        return "高级数据分析（回归分析、时间序列、预测）"

    @property
    def schema(self) -> list[CapabilitySchema]:
        return [
            CapabilitySchema(name="file", type="string", description="文件路径", required=True),
            CapabilitySchema(name="sheet", type="string", description="工作表名称", required=False),
            CapabilitySchema(
                name="analysis_type",
                type="string",
                description="分析类型（regression/timeseries/forecast）",
                required=True,
            ),
            CapabilitySchema(
                name="x_column", type="string", description="自变量列名", required=False
            ),
            CapabilitySchema(
                name="y_column", type="string", description="因变量列名", required=False
            ),
            CapabilitySchema(name="periods", type="number", description="预测期数", required=False),
        ]

    def execute(self, context: Any, **params) -> Any:
        file_path = params.get("file")
        sheet_name = params.get("sheet")
        analysis_type = params.get("analysis_type")
        x_column = params.get("x_column")
        y_column = params.get("y_column")
        periods = params.get("periods", 10)

        if not file_path:
            raise ValidationError("执行失败: 缺少必要参数 file")

        path = Path(file_path)
        if not path.exists():
            raise FileNotFoundError(f"文件操作失败: 文件不存在 {file_path}")

        df = self._load_data(path, sheet_name)

        if analysis_type == "regression":
            return self._regression_analysis(df, x_column, y_column)
        elif analysis_type == "timeseries":
            return self._timeseries_analysis(df, y_column)
        elif analysis_type == "forecast":
            return self._forecast(df, y_column, periods)
        else:
            raise DataError(f"数据操作失败: 不支持的分析类型 {analysis_type}")

    def _load_data(self, path: Path, sheet_name: str = None) -> pd.DataFrame:
        suffix = path.suffix.lower()
        if suffix in [".xlsx", ".xls"]:
            return pd.read_excel(path, sheet_name=sheet_name or 0)
        elif suffix == ".csv":
            return pd.read_csv(path)
        else:
            raise DataError(f"数据操作失败: 不支持的文件格式 {suffix}")

    def _regression_analysis(self, df: pd.DataFrame, x_col: str, y_col: str) -> dict:
        """线性回归分析"""
        if x_col not in df.columns or y_col not in df.columns:
            raise DataError(f"数据操作失败: 列 {x_col} 或 {y_col} 不存在")

        x = df[x_col].values.astype(float)
        y = df[y_col].values.astype(float)

        mask = ~(np.isnan(x) | np.isnan(y))
        x = x[mask]
        y = y[mask]

        if len(x) < 2:
            raise DataError("数据操作失败: 回归分析至少需要 2 个数据点")

        n = len(x)
        x_mean = np.mean(x)
        y_mean = np.mean(y)

        numerator = np.sum((x - x_mean) * (y - y_mean))
        denominator = np.sum((x - x_mean) ** 2)

        if denominator == 0:
            raise DataError("数据操作失败: X 值全部相同，无法进行回归分析")

        slope = numerator / denominator
        intercept = y_mean - slope * x_mean

        y_pred = slope * x + intercept
        ss_res = np.sum((y - y_pred) ** 2)
        ss_tot = np.sum((y - y_mean) ** 2)
        r_squared = 1 - (ss_res / ss_tot) if ss_tot != 0 else 0

        return {
            "type": "regression",
            "slope": float(slope),
            "intercept": float(intercept),
            "r_squared": float(r_squared),
            "n": n,
            "equation": f"y = {slope:.4f}x + {intercept:.4f}",
        }

    def _timeseries_analysis(self, df: pd.DataFrame, value_col: str) -> dict:
        """时间序列分析"""
        if value_col not in df.columns:
            raise DataError(f"数据操作失败: 列 {value_col} 不存在")

        values = df[value_col].dropna().values.astype(float)

        if len(values) < 4:
            raise DataError("数据操作失败: 时间序列分析至少需要 4 个数据点")

        stats = {
            "mean": float(np.mean(values)),
            "std": float(np.std(values)),
            "min": float(np.min(values)),
            "max": float(np.max(values)),
        }

        x = np.arange(len(values))
        slope, intercept = np.polyfit(x, values, 1)

        window = min(3, len(values) // 2)
        if window > 1:
            ma = np.convolve(values, np.ones(window) / window, mode="valid")
            stats["moving_average"] = ma.tolist()

        if len(values) >= 8:
            autocorr = np.corrcoef(values[:-1], values[1:])[0, 1]
            stats["autocorrelation"] = float(autocorr)

        return {
            "type": "timeseries",
            "statistics": stats,
            "trend": {
                "slope": float(slope),
                "direction": "上升" if slope > 0 else "下降" if slope < 0 else "平稳",
            },
            "data_points": len(values),
        }

    def _forecast(self, df: pd.DataFrame, value_col: str, periods: int) -> dict:
        """简单预测（线性外推）"""
        if value_col not in df.columns:
            raise DataError(f"数据操作失败: 列 {value_col} 不存在")

        values = df[value_col].dropna().values.astype(float)

        if len(values) < 2:
            raise DataError("数据操作失败: 预测至少需要 2 个数据点")

        x = np.arange(len(values))
        slope, intercept = np.polyfit(x, values, 1)

        forecast_x = np.arange(len(values), len(values) + periods)
        forecast_y = slope * forecast_x + intercept

        return {
            "type": "forecast",
            "historical": values.tolist(),
            "forecast": forecast_y.tolist(),
            "periods": periods,
            "trend": float(slope),
        }
