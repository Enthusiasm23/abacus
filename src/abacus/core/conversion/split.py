"""工作表拆分 - 将 Excel 工作表按条件拆分为多个文件"""

import logging
from pathlib import Path
from typing import Any

import pandas as pd

from ..base import Capability, CapabilitySchema
from ..exceptions import DataError, FileNotFoundError, ValidationError

logger = logging.getLogger(__name__)


class SplitSheetCapability(Capability):
    """工作表拆分 - 将 Excel 工作表按条件拆分为多个文件"""

    @property
    def name(self) -> str:
        return "split_sheet"

    @property
    def chapter(self) -> str:
        return "work"

    @property
    def description(self) -> str:
        return "将 Excel 工作表按条件拆分为多个文件（按列值、行数、范围）"

    @property
    def schema(self) -> list[CapabilitySchema]:
        return [
            CapabilitySchema(
                name="file", type="string", description="Excel 文件路径", required=True
            ),
            CapabilitySchema(name="sheet", type="string", description="工作表名称", required=True),
            CapabilitySchema(
                name="output_dir", type="string", description="输出目录", required=True
            ),
            CapabilitySchema(
                name="split_by",
                type="string",
                description="拆分方式（column/row_count/range）",
                required=True,
            ),
            CapabilitySchema(
                name="split_column",
                type="string",
                description="拆分列名（split_by=column 时）",
                required=False,
            ),
            CapabilitySchema(
                name="row_count",
                type="number",
                description="每文件行数（split_by=row_count 时）",
                required=False,
            ),
            CapabilitySchema(
                name="prefix", type="string", description="输出文件前缀", required=False
            ),
        ]

    def execute(self, context: Any, **params) -> Any:
        file_path = params.get("file")
        sheet_name = params.get("sheet")
        output_dir = params.get("output_dir")
        split_by = params.get("split_by")
        split_column = params.get("split_column")
        row_count = params.get("row_count")
        prefix = params.get("prefix", "split")

        if not file_path:
            raise ValidationError("执行失败: 缺少必要参数 file")
        if not sheet_name:
            raise ValidationError("执行失败: 缺少必要参数 sheet")
        if not output_dir:
            raise ValidationError("执行失败: 缺少必要参数 output_dir")
        if not split_by:
            raise ValidationError("执行失败: 缺少必要参数 split_by")

        path = Path(file_path)
        if not path.exists():
            raise FileNotFoundError(f"文件操作失败: 文件不存在 {file_path}")

        # 加载数据
        try:
            df = pd.read_excel(path, sheet_name=sheet_name)
        except Exception as e:  # noqa: BLE001
            raise DataError(f"数据操作失败: 读取 Excel 文件失败 {e}")

        # 创建输出目录
        output_path = Path(output_dir)
        output_path.mkdir(parents=True, exist_ok=True)

        # 执行拆分
        if split_by == "column":
            if not split_column:
                raise ValidationError("执行失败: 按列拆分需要 split_column 参数")
            if split_column not in df.columns:
                raise DataError(f"数据操作失败: 列 '{split_column}' 不存在")

            result = self._split_by_column(df, split_column, output_path, prefix)

        elif split_by == "row_count":
            if not row_count or row_count <= 0:
                raise ValidationError("执行失败: row_count 必须为正数")

            result = self._split_by_row_count(df, int(row_count), output_path, prefix)

        elif split_by == "range":
            # 按范围拆分（每100行一个文件）
            result = self._split_by_row_count(df, 100, output_path, prefix)

        else:
            raise DataError(f"数据操作失败: 不支持的拆分方式 {split_by}")

        return {
            "file": file_path,
            "sheet": sheet_name,
            "split_by": split_by,
            "files_created": len(result),
            "output_dir": output_dir,
            "files": result,
        }

    def _split_by_column(
        self, df: pd.DataFrame, column: str, output_path: Path, prefix: str
    ) -> [str]:
        """按列值拆分"""
        files = []

        for value, group in df.groupby(column):
            # 生成文件名
            safe_value = str(value).replace("/", "_").replace("\\", "_")[:50]
            filename = f"{prefix}_{safe_value}.xlsx"
            filepath = output_path / filename

            # 保存
            group.to_excel(filepath, index=False)
            files.append(str(filepath))

        return files

    def _split_by_row_count(
        self, df: pd.DataFrame, row_count: int, output_path: Path, prefix: str
    ) -> [str]:
        """按行数拆分"""
        files = []
        total_rows = len(df)

        for i in range(0, total_rows, row_count):
            chunk = df.iloc[i : i + row_count]

            # 生成文件名
            filename = f"{prefix}_{i // row_count + 1:03d}.xlsx"
            filepath = output_path / filename

            # 保存
            chunk.to_excel(filepath, index=False)
            files.append(str(filepath))

        return files
