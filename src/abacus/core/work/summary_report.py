"""商功章 - 摘要报告：自动生成数据摘要"""

import logging
from pathlib import Path
from typing import Any

import pandas as pd

from ..base import Capability, CapabilitySchema
from ..exceptions import DataError, FileNotFoundError

logger = logging.getLogger(__name__)


class SummaryReportCapability(Capability):
    """摘要报告：自动生成数据摘要"""

    @property
    def name(self) -> str:
        return "generate_summary_report"

    @property
    def chapter(self) -> str:
        return "work"

    @property
    def description(self) -> str:
        return "自动生成数据摘要（行列数、类型分布、质量问题）"

    @property
    def schema(self) -> list[CapabilitySchema]:
        return [
            CapabilitySchema(
                name="file", type="string", description="Excel 文件路径", required=True
            ),
            CapabilitySchema(name="sheet", type="string", description="工作表名称", required=True),
        ]

    def execute(self, context: Any, **params) -> Any:
        """生成摘要报告"""
        file_path = params.get("file")
        sheet_name = params.get("sheet")

        if not file_path:
            raise DataError("file parameter is required")
        if not sheet_name:
            raise DataError("sheet parameter is required")

        return self._generate_summary_report(file_path, sheet_name)

    def _generate_summary_report(self, filepath: str, sheet_name: str) -> dict[str, Any]:
        """生成摘要报告"""
        try:
            path = Path(filepath)
            if not path.exists():
                raise FileNotFoundError(f"File not found: {filepath}")

            # 读取数据
            df = pd.read_excel(path, sheet_name=sheet_name)

            # 生成摘要
            summary = {
                "success": True,
                "file": filepath,
                "sheet": sheet_name,
                "total_rows": len(df),
                "total_columns": len(df.columns),
                "columns": (df.columns),
                "dtypes": {col: str(dtype) for col, dtype in df.dtypes.items()},
                "null_counts": df.isnull().sum().to_dict(),
                "null_percentages": (df.isnull().sum() / len(df) * 100).to_dict(),
                "numeric_stats": {},
                "categorical_stats": {},
            }

            # 数值列统计
            numeric_cols = df.select_dtypes(include=["number"]).columns
            for col in numeric_cols:
                summary["numeric_stats"][col] = {
                    "mean": float(df[col].mean()) if not df[col].isnull().all() else None,
                    "std": float(df[col].std()) if not df[col].isnull().all() else None,
                    "min": float(df[col].min()) if not df[col].isnull().all() else None,
                    "max": float(df[col].max()) if not df[col].isnull().all() else None,
                    "median": float(df[col].median()) if not df[col].isnull().all() else None,
                }

            # 分类列统计
            categorical_cols = df.select_dtypes(include=["object"]).columns
            for col in categorical_cols:
                summary["categorical_stats"][col] = {
                    "unique_count": int(df[col].nunique()),
                    "top_values": df[col].value_counts().head(5).to_dict(),
                }

            return summary

        except FileNotFoundError:
            raise
        except DataError:
            raise
        except Exception as e:
            logger.error(f"Failed to generate summary report: {e}")
            raise DataError(str(e))
