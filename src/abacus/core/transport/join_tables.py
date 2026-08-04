"""均输章 - 关联表：SQL 风格关联（LEFT/RIGHT/INNER/OUTER JOIN）"""

import logging
from pathlib import Path
from typing import Any

import pandas as pd

from ..base import Capability, CapabilitySchema
from ..exceptions import DataError, FileNotFoundError, ValidationError

logger = logging.getLogger(__name__)


class JoinTablesCapability(Capability):
    """关联表：SQL 风格关联"""

    @property
    def name(self) -> str:
        return "join_tables"

    @property
    def chapter(self) -> str:
        return "transport"

    @property
    def description(self) -> str:
        return "SQL 风格关联（LEFT/RIGHT/INNER/OUTER JOIN）"

    @property
    def schema(self) -> list[CapabilitySchema]:
        return [
            CapabilitySchema(
                name="left_file", type="string", description="左表 Excel 文件路径", required=True
            ),
            CapabilitySchema(
                name="left_sheet", type="string", description="左表工作表名称", required=True
            ),
            CapabilitySchema(
                name="right_file", type="string", description="右表 Excel 文件路径", required=True
            ),
            CapabilitySchema(
                name="right_sheet", type="string", description="右表工作表名称", required=True
            ),
            CapabilitySchema(
                name="on", type="array", description="关联键（列名列表）", required=True
            ),
            CapabilitySchema(
                name="how",
                type="string",
                description="关联类型：left/right/inner/outer",
                required=False,
                default="inner",
            ),
            CapabilitySchema(
                name="output", type="string", description="输出文件路径（可选）", required=False
            ),
        ]

    def execute(self, context: Any, **params) -> Any:
        """执行关联"""
        left_file = params.get("left_file")
        left_sheet = params.get("left_sheet")
        right_file = params.get("right_file")
        right_sheet = params.get("right_sheet")
        on = params.get("on", [])
        how = params.get("how", "inner")
        output = params.get("output")

        if not left_file:
            raise ValidationError("执行失败: 缺少必要参数 left_file")
        if not right_file:
            raise ValidationError("执行失败: 缺少必要参数 right_file")
        if not on:
            raise ValidationError("执行失败: 缺少必要参数 on")

        return self._join_tables(left_file, left_sheet, right_file, right_sheet, on, how, output)

    def _join_tables(
        self,
        left_file: str,
        left_sheet: str,
        right_file: str,
        right_sheet: str,
        on: list[str],
        how: str,
        output: str = None,
    ) -> dict[str, Any]:
        """执行关联"""
        try:
            left_path = Path(left_file)
            right_path = Path(right_file)

            if not left_path.exists():
                raise FileNotFoundError(f"文件操作失败: 左表文件不存在 {left_file}")
            if not right_path.exists():
                raise FileNotFoundError(f"文件操作失败: 右表文件不存在 {right_file}")

            # 读取数据
            df_left = pd.read_excel(left_path, sheet_name=left_sheet)
            df_right = pd.read_excel(right_path, sheet_name=right_sheet)

            # 检查关联键
            for key in on:
                if key not in df_left.columns:
                    raise DataError(f"数据操作失败: 左表中未找到关联键 '{key}'")
                if key not in df_right.columns:
                    raise DataError(f"数据操作失败: 右表中未找到关联键 '{key}'")

            # 执行关联
            result = pd.merge(df_left, df_right, on=on, how=how, suffixes=("_left", "_right"))

            # 保存结果
            if output:
                result.to_excel(output, index=False, sheet_name="MergedData")

            return {
                "success": True,
                "left_file": left_file,
                "right_file": right_file,
                "join_type": how,
                "join_keys": on,
                "left_rows": len(df_left),
                "right_rows": len(df_right),
                "result_rows": len(result),
                "result_columns": (result.columns),
                "output": output,
            }

        except FileNotFoundError:
            raise
        except DataError:
            raise
        except Exception as e:
            logger.error(f"表关联失败: {e}")
            raise DataError(f"数据操作失败: {e}")
