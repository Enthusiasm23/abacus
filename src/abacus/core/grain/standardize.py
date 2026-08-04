"""粟米章 - 标准化：统一日期、数字、文本格式"""

import logging
from pathlib import Path
from typing import Any

import pandas as pd

from ..base import Capability, CapabilitySchema
from ..exceptions import DataError, FileNotFoundError

logger = logging.getLogger(__name__)


class StandardizeCapability(Capability):
    """标准化：统一日期、数字、文本格式"""

    @property
    def name(self) -> str:
        return "standardize_data"

    @property
    def chapter(self) -> str:
        return "grain"

    @property
    def description(self) -> str:
        return "统一日期、数字、文本格式"

    @property
    def schema(self) -> list[CapabilitySchema]:
        return [
            CapabilitySchema(
                name="file", type="string", description="Excel 文件路径", required=True
            ),
            CapabilitySchema(name="sheet", type="string", description="工作表名称", required=True),
            CapabilitySchema(
                name="date_format",
                type="string",
                description="日期格式（如 %Y-%m-%d）",
                required=False,
            ),
            CapabilitySchema(
                name="number_format",
                type="string",
                description="数字格式（如 %.2f）",
                required=False,
            ),
            CapabilitySchema(
                name="text_case",
                type="string",
                description="文本大小写：lower/upper/title",
                required=False,
            ),
            CapabilitySchema(
                name="output", type="string", description="输出文件路径（可选）", required=False
            ),
        ]

    def execute(self, context: Any, **params) -> Any:
        """执行标准化"""
        file_path = params.get("file")
        sheet_name = params.get("sheet")
        date_format = params.get("date_format")
        number_format = params.get("number_format")
        text_case = params.get("text_case")
        output = params.get("output")

        if not file_path:
            raise DataError("file parameter is required")
        if not sheet_name:
            raise DataError("sheet parameter is required")

        return self._standardize(
            file_path, sheet_name, date_format, number_format, text_case, output
        )

    def _standardize(
        self,
        filepath: str,
        sheet_name: str,
        date_format: str = None,
        number_format: str = None,
        text_case: str = None,
        output: str = None,
    ) -> dict[str, Any]:
        """标准化数据"""
        try:
            path = Path(filepath)
            if not path.exists():
                raise FileNotFoundError(f"File not found: {filepath}")

            # 读取数据
            df = pd.read_excel(path, sheet_name=sheet_name)

            # 记录标准化操作
            operations = []

            # 标准化日期
            if date_format:
                for col in df.columns:
                    if pd.api.types.is_datetime64_any_dtype(df[col]):
                        df[col] = df[col].dt.strftime(date_format)
                        operations.append(
                            {"column": col, "operation": "date_format", "format": date_format}
                        )

            # 标准化数字
            if number_format:
                for col in df.columns:
                    if pd.api.types.is_numeric_dtype(df[col]):
                        try:
                            df[col] = df[col].apply(
                                lambda x: float(format(x, number_format)) if pd.notna(x) else x
                            )
                            operations.append(
                                {
                                    "column": col,
                                    "operation": "number_format",
                                    "format": number_format,
                                }
                            )
                        except Exception:
                            pass

            # 标准化文本
            if text_case:
                for col in df.columns:
                    if pd.api.types.is_object_dtype(df[col]):
                        if text_case == "lower":
                            df[col] = df[col].str.lower()
                        elif text_case == "upper":
                            df[col] = df[col].str.upper()
                        elif text_case == "title":
                            df[col] = df[col].str.title()
                        operations.append(
                            {"column": col, "operation": "text_case", "case": text_case}
                        )

            # 保存结果
            if output:
                df.to_excel(output, index=False, sheet_name=sheet_name)

            return {
                "success": True,
                "file": filepath,
                "sheet": sheet_name,
                "operations": operations,
                "operation_count": len(operations),
                "output": output,
            }

        except FileNotFoundError:
            raise
        except DataError:
            raise
        except Exception as e:
            logger.error(f"Failed to standardize: {e}")
            raise DataError(str(e))
