"""粟米章 - 类型推断：自动检测并转换数据类型"""

import logging
from pathlib import Path
from typing import Any

import pandas as pd

from ..base import Capability, CapabilitySchema
from ..exceptions import DataError, FileNotFoundError, ValidationError

logger = logging.getLogger(__name__)


class AutoTypeInferCapability(Capability):
    """类型推断：自动检测并转换数据类型"""

    @property
    def name(self) -> str:
        return "auto_type_infer"

    @property
    def chapter(self) -> str:
        return "grain"

    @property
    def description(self) -> str:
        return "自动检测并转换数据类型（文本→数字、日期等）"

    @property
    def schema(self) -> list[CapabilitySchema]:
        return [
            CapabilitySchema(
                name="file", type="string", description="Excel 文件路径", required=True
            ),
            CapabilitySchema(name="sheet", type="string", description="工作表名称", required=True),
            CapabilitySchema(
                name="range",
                type="string",
                description="数据范围（可选，默认全部）",
                required=False,
            ),
            CapabilitySchema(
                name="output", type="string", description="输出文件路径（可选）", required=False
            ),
        ]

    def execute(self, context: Any, **params) -> Any:
        """执行类型推断"""
        file_path = params.get("file")
        sheet_name = params.get("sheet")
        data_range = params.get("range")
        output = params.get("output")

        if not file_path:
            raise ValidationError("执行失败: 缺少必要参数 file")
        if not sheet_name:
            raise ValidationError("执行失败: 缺少必要参数 sheet")

        return self._auto_type_infer(file_path, sheet_name, data_range, output)

    def _auto_type_infer(
        self, filepath: str, sheet_name: str, data_range: str = None, output: str = None
    ) -> dict[str, Any]:
        """自动类型推断"""
        try:
            path = Path(filepath)
            if not path.exists():
                raise FileNotFoundError(f"文件操作失败: 文件不存在 {filepath}")

            # 读取数据
            df = pd.read_excel(path, sheet_name=sheet_name)

            # 推断类型
            inferred_types = {}
            conversions = {}

            for col in df.columns:
                original_type = str(df[col].dtype)
                inferred_type = self._infer_column_type(df[col])

                inferred_types[col] = {
                    "original_type": original_type,
                    "inferred_type": inferred_type,
                    "sample_values": df[col].dropna().head(3).tolist(),
                }

                # 执行转换
                if inferred_type != original_type:
                    try:
                        df[col] = self._convert_column(df[col], inferred_type)
                        conversions[col] = {
                            "from": original_type,
                            "to": inferred_type,
                            "success": True,
                        }
                    except Exception as e:
                        conversions[col] = {
                            "from": original_type,
                            "to": inferred_type,
                            "success": False,
                            "error": str(e),
                        }

            # 保存结果
            if output:
                df.to_excel(output, index=False, sheet_name=sheet_name)

            return {
                "success": True,
                "file": filepath,
                "sheet": sheet_name,
                "total_columns": len(df.columns),
                "inferred_types": inferred_types,
                "conversions": conversions,
                "conversion_count": sum(1 for c in conversions.values() if c.get("success")),
                "output": output,
            }

        except FileNotFoundError:
            raise
        except DataError:
            raise
        except Exception as e:
            logger.error(f"类型推断失败: {e}")
            raise DataError(f"数据操作失败: {e}")

    def _infer_column_type(self, series: pd.Series) -> str:
        """推断列类型"""
        # 去除空值
        non_null = series.dropna()

        if len(non_null) == 0:
            return "object"

        # 尝试转换为整数（必须在 float 之前检查）
        try:
            pd.to_numeric(non_null, downcast="integer")
            return "int64"
        except (ValueError, TypeError):
            pass

        # 尝试转换为数值
        try:
            pd.to_numeric(non_null)
            return "float64"
        except (ValueError, TypeError):
            pass

        # 尝试转换为日期
        try:
            pd.to_datetime(non_null)
            return "datetime64[ns]"
        except (ValueError, TypeError):
            pass

        return "object"

    def _convert_column(self, series: pd.Series, target_type: str) -> pd.Series:
        """转换列类型"""
        if target_type == "float64":
            return pd.to_numeric(series, errors="coerce")
        elif target_type == "int64":
            return pd.to_numeric(series, errors="coerce").astype("Int64")
        elif target_type == "datetime64[ns]":
            return pd.to_datetime(series, errors="coerce")
        else:
            return series
