"""数据分析 - 智能数据检测和统计分析"""

import logging
from pathlib import Path
from typing import Any

import pandas as pd

from ..base import Capability, CapabilitySchema
from ..exceptions import DataError, FileNotFoundError, ValidationError

logger = logging.getLogger(__name__)


class DataAnalysisCapability(Capability):
    """数据分析 - 智能数据检测和统计分析"""

    @property
    def name(self) -> str:
        return "analyze_data"

    @property
    def chapter(self) -> str:
        return "triangle"

    @property
    def description(self) -> str:
        return "智能数据分析（自动检测数据类型、统计摘要、相关性分析）"

    @property
    def schema(self) -> list[CapabilitySchema]:
        return [
            CapabilitySchema(
                name="file", type="string", description="文件路径（Excel 或 CSV）", required=True
            ),
            CapabilitySchema(
                name="sheet", type="string", description="工作表名称（Excel 文件）", required=False
            ),
            CapabilitySchema(
                name="analysis_type",
                type="string",
                description="分析类型（auto/summary/correlation）",
                required=False,
            ),
        ]

    def execute(self, context: Any, **params) -> Any:
        file_path = params.get("file")
        sheet_name = params.get("sheet")
        analysis_type = params.get("analysis_type", "auto")

        if not file_path:
            raise ValidationError("执行失败: 缺少必要参数 file")

        path = Path(file_path)
        if not path.exists():
            raise FileNotFoundError(f"文件操作失败: 文件不存在 {file_path}")

        df = self._load_data(path, sheet_name)
        data_info = self._detect_data_types(df)

        if analysis_type == "auto":
            result = self._auto_analysis(df, data_info)
        elif analysis_type == "summary":
            result = self._summary_analysis(df)
        elif analysis_type == "correlation":
            result = self._correlation_analysis(df)
        else:
            raise DataError(f"数据操作失败: 不支持的分析类型 {analysis_type}")

        return {
            "file": file_path,
            "rows": len(df),
            "columns": len(df.columns),
            "data_types": data_info,
            "analysis": result,
        }

    def _load_data(self, path: Path, sheet_name: str = None) -> pd.DataFrame:
        suffix = path.suffix.lower()

        if suffix in [".xlsx", ".xls"]:
            if sheet_name:
                return pd.read_excel(path, sheet_name=sheet_name)
            else:
                # 默认读取第一个工作表
                return pd.read_excel(path, sheet_name=0)
        elif suffix == ".csv":
            return pd.read_csv(path)
        else:
            raise DataError(f"数据操作失败: 不支持的文件格式 {suffix}")

    def _detect_data_types(self, df: pd.DataFrame) -> dict:
        info = {
            "numeric_columns": [],
            "categorical_columns": [],
            "datetime_columns": [],
            "text_columns": [],
        }

        for col in df.columns:
            dtype = df[col].dtype

            if pd.api.types.is_numeric_dtype(dtype):
                info["numeric_columns"].append(col)
            elif pd.api.types.is_datetime64_any_dtype(dtype):
                info["datetime_columns"].append(col)
            elif df[col].nunique() < len(df) * 0.5:
                info["categorical_columns"].append(col)
            else:
                info["text_columns"].append(col)

        return info

    def _auto_analysis(self, df: pd.DataFrame, data_info: dict) -> dict:
        result = {}

        result["summary"] = {
            "rows": len(df),
            "columns": len(df.columns),
            "missing_values": int(df.isnull().sum().sum()),
            "duplicate_rows": int(df.duplicated().sum()),
        }

        if data_info["numeric_columns"]:
            result["numeric_stats"] = df[data_info["numeric_columns"]].describe().to_dict()

        if len(data_info["numeric_columns"]) > 1:
            result["correlation"] = df[data_info["numeric_columns"]].corr().to_dict()

        if data_info["categorical_columns"]:
            result["categorical_stats"] = {}
            for col in data_info["categorical_columns"]:
                result["categorical_stats"][col] = {
                    "unique_count": int(df[col].nunique()),
                    "top_values": df[col].value_counts().head(5).to_dict(),
                }

        return result

    def _summary_analysis(self, df: pd.DataFrame) -> dict:
        return {
            "shape": df.shape,
            "dtypes": df.dtypes.astype(str).to_dict(),
            "missing": df.isnull().sum().to_dict(),
            "numeric_summary": df.describe().to_dict(),
        }

    def _correlation_analysis(self, df: pd.DataFrame) -> dict:
        numeric_df = df.select_dtypes(include=["number"])
        if len(numeric_df.columns) < 2:
            return {"error": "需要至少2个数值列进行相关性分析"}

        return {
            "correlation_matrix": numeric_df.corr().to_dict(),
            "strong_correlations": self._find_strong_correlations(numeric_df),
        }

    def _find_strong_correlations(self, df: pd.DataFrame, threshold: float = 0.7) -> list:
        corr_matrix = df.corr()
        strong = []

        for i in range(len(corr_matrix.columns)):
            for j in range(i + 1, len(corr_matrix.columns)):
                if abs(corr_matrix.iloc[i, j]) >= threshold:
                    strong.append(
                        {
                            "col1": corr_matrix.columns[i],
                            "col2": corr_matrix.columns[j],
                            "correlation": round(corr_matrix.iloc[i, j], 4),
                        }
                    )

        return strong
