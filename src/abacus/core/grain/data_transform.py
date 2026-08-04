"""粟米章 - 数据转换：用 pandas 实现高级转换"""

import logging
from pathlib import Path
from typing import Any

import pandas as pd

from ..base import Capability, CapabilitySchema
from ..exceptions import DataError, FileNotFoundError

logger = logging.getLogger(__name__)


class DataTransformCapability(Capability):
    """数据转换：用 pandas 实现高级数据转换"""

    @property
    def name(self) -> str:
        return "transform_data"

    @property
    def chapter(self) -> str:
        return "grain"

    @property
    def description(self) -> str:
        return "高级数据转换（透视、转置、合并、重塑）"

    @property
    def schema(self) -> list[CapabilitySchema]:
        return [
            CapabilitySchema(name="file", type="string", description="文件路径", required=True),
            CapabilitySchema(name="sheet", type="string", description="工作表名称", required=False),
            CapabilitySchema(
                name="transform_type",
                type="string",
                description="转换类型（pivot/melt/merge/reshape）",
                required=True,
            ),
            CapabilitySchema(name="params", type="object", description="转换参数", required=False),
            CapabilitySchema(
                name="output", type="string", description="输出文件路径", required=False
            ),
        ]

    def execute(self, context: Any, **params) -> Any:
        file_path = params.get("file")
        sheet_name = params.get("sheet")
        transform_type = params.get("transform_type")
        transform_params = params.get("params", {})
        output = params.get("output")

        if not file_path:
            raise DataError("file parameter is required")

        path = Path(file_path)
        if not path.exists():
            raise FileNotFoundError(f"File not found: {file_path}")

        df = self._load_data(path, sheet_name)

        if transform_type == "pivot":
            result_df = self._pivot(df, transform_params)
        elif transform_type == "melt":
            result_df = self._melt(df, transform_params)
        elif transform_type == "merge":
            result_df = self._merge(df, transform_params)
        elif transform_type == "reshape":
            result_df = self._reshape(df, transform_params)
        else:
            raise DataError(f"Unknown transform type: {transform_type}")

        if output:
            output_path = Path(output)
            if output_path.suffix.lower() == ".csv":
                result_df.to_csv(output_path, index=False)
            else:
                result_df.to_excel(output_path, index=False)

        return {
            "transform_type": transform_type,
            "input_rows": len(df),
            "output_rows": len(result_df),
            "output_columns": len(result_df.columns),
            "output": output,
        }

    def _load_data(self, path: Path, sheet_name: str = None) -> pd.DataFrame:
        suffix = path.suffix.lower()
        if suffix in [".xlsx", ".xls"]:
            return pd.read_excel(path, sheet_name=sheet_name or 0)
        elif suffix == ".csv":
            return pd.read_csv(path)
        else:
            raise DataError(f"Unsupported format: {suffix}")

    def _pivot(self, df: pd.DataFrame, params: dict) -> pd.DataFrame:
        """透视表转换"""
        index = params.get("index")
        columns = params.get("columns")
        values = params.get("values")
        aggfunc = params.get("aggfunc", "sum")

        if not index or not values:
            raise DataError("index and values required for pivot")

        return pd.pivot_table(
            df, index=index, columns=columns, values=values, aggfunc=aggfunc
        ).reset_index()

    def _melt(self, df: pd.DataFrame, params: dict) -> pd.DataFrame:
        """逆透视（宽表转长表）"""
        id_vars = params.get("id_vars", [])
        value_vars = params.get("value_vars", [])

        return pd.melt(df, id_vars=id_vars, value_vars=value_vars if value_vars else None)

    def _merge(self, df: pd.DataFrame, params: dict) -> pd.DataFrame:
        """合并数据（横向拼接另一文件）"""
        other_file = params.get("other_file")
        other_sheet = params.get("other_sheet")
        on = params.get("on")
        how = params.get("how", "inner")

        if not other_file:
            raise DataError("other_file required for merge")

        other_path = Path(other_file)
        if not other_path.exists():
            raise FileNotFoundError(f"File not found: {other_file}")

        other_df = self._load_data(other_path, other_sheet)

        if on:
            return pd.merge(df, other_df, on=on, how=how)
        else:
            return pd.concat([df, other_df], axis=1)

    def _reshape(self, df: pd.DataFrame, params: dict) -> pd.DataFrame:
        """重塑数据"""
        pivot_col = params.get("pivot_column")
        value_col = params.get("value_column")
        index_col = params.get("index_column")

        if not pivot_col or not value_col:
            raise DataError("pivot_column and value_column required for reshape")

        if index_col:
            return df.pivot(index=index_col, columns=pivot_col, values=value_col).reset_index()
        else:
            return df.pivot(columns=pivot_col, values=value_col).reset_index()
