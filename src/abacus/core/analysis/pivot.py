"""数据透视 - 智能数据分析和汇总"""

import logging
from pathlib import Path
from typing import Any

import pandas as pd

from ..base import Capability, CapabilitySchema
from ..exceptions import DataError, FileNotFoundError, ValidationError

logger = logging.getLogger(__name__)


class PivotAnalysisCapability(Capability):
    """数据透视 - 智能数据分析和汇总"""

    @property
    def name(self) -> str:
        return "pivot_analysis"

    @property
    def chapter(self) -> str:
        return "share"

    @property
    def description(self) -> str:
        return "数据透视分析（分组汇总、交叉分析）"

    @property
    def schema(self) -> list[CapabilitySchema]:
        return [
            CapabilitySchema(name="file", type="string", description="文件路径", required=True),
            CapabilitySchema(name="sheet", type="string", description="工作表名称", required=False),
            CapabilitySchema(name="group_by", type="string", description="分组字段", required=True),
            CapabilitySchema(
                name="value_field", type="string", description="值字段", required=True
            ),
            CapabilitySchema(
                name="agg_function",
                type="string",
                description="聚合函数（sum/mean/count）",
                required=False,
            ),
            CapabilitySchema(
                name="output", type="string", description="输出文件路径", required=False
            ),
        ]

    def execute(self, context: Any, **params) -> Any:
        file_path = params.get("file")
        sheet_name = params.get("sheet")
        group_by = params.get("group_by")
        value_field = params.get("value_field")
        agg_function = params.get("agg_function", "sum")
        output = params.get("output")

        if not file_path:
            raise ValidationError("执行失败: 缺少必要参数 file")
        if not group_by:
            raise ValidationError("执行失败: 缺少必要参数 group_by")
        if not value_field:
            raise ValidationError("执行失败: 缺少必要参数 value_field")

        path = Path(file_path)
        if not path.exists():
            raise FileNotFoundError(f"文件操作失败: 文件不存在 {file_path}")

        suffix = path.suffix.lower()
        if suffix in [".xlsx", ".xls"]:
            if sheet_name:
                df = pd.read_excel(path, sheet_name=sheet_name)
            else:
                df = pd.read_excel(path, sheet_name=0)
        elif suffix == ".csv":
            df = pd.read_csv(path)
        else:
            raise DataError(f"数据操作失败: 不支持的文件格式 {suffix}")

        if group_by not in df.columns:
            raise DataError(f"数据操作失败: 列 '{group_by}' 不存在")
        if value_field not in df.columns:
            raise DataError(f"数据操作失败: 列 '{value_field}' 不存在")

        if agg_function == "sum":
            result = df.groupby(group_by)[value_field].sum()
        elif agg_function == "mean":
            result = df.groupby(group_by)[value_field].mean()
        elif agg_function == "count":
            result = df.groupby(group_by)[value_field].count()
        else:
            raise DataError(f"数据操作失败: 不支持的聚合函数 {agg_function}")

        result_df = result.reset_index()
        result_df.columns = [group_by, f"{agg_function}({value_field})"]

        if output:
            output_path = Path(output)
            if output_path.suffix.lower() == ".csv":
                result_df.to_csv(output_path, index=False)
            else:
                result_df.to_excel(output_path, index=False)

        return {
            "file": file_path,
            "group_by": group_by,
            "value_field": value_field,
            "agg_function": agg_function,
            "groups": len(result_df),
            "result": result_df.to_dict(orient="records"),
        }
