"""均输章 - 批量合并：从文件夹批量合并多个 Excel 文件"""

import logging
from pathlib import Path
from typing import Any

import pandas as pd

from ..base import Capability, CapabilitySchema
from ..exceptions import DataError, FileNotFoundError, ValidationError

logger = logging.getLogger(__name__)


class BatchMergeCapability(Capability):
    """批量合并：从文件夹批量合并多个 Excel 文件"""

    @property
    def name(self) -> str:
        return "batch_merge"

    @property
    def chapter(self) -> str:
        return "transport"

    @property
    def description(self) -> str:
        return "从文件夹批量合并多个 Excel 文件"

    @property
    def schema(self) -> list[CapabilitySchema]:
        return [
            CapabilitySchema(name="folder", type="string", description="文件夹路径", required=True),
            CapabilitySchema(
                name="pattern",
                type="string",
                description="文件匹配模式（默认 *.xlsx）",
                required=False,
                default="*.xlsx",
            ),
            CapabilitySchema(
                name="sheet", type="string", description="工作表名称（可选）", required=False
            ),
            CapabilitySchema(
                name="output", type="string", description="输出文件路径", required=True
            ),
        ]

    def execute(self, context: Any, **params) -> Any:
        """执行批量合并"""
        folder = params.get("folder")
        pattern = params.get("pattern", "*.xlsx")
        sheet = params.get("sheet")
        output = params.get("output")

        if not folder:
            raise ValidationError("执行失败: 缺少必要参数 folder")
        if not output:
            raise ValidationError("执行失败: 缺少必要参数 output")

        return self._batch_merge(folder, pattern, sheet, output)

    def _batch_merge(self, folder: str, pattern: str, sheet: str, output: str) -> dict:
        """批量合并"""
        try:
            folder_path = Path(folder)
            if not folder_path.exists():
                raise FileNotFoundError(f"文件操作失败: 文件夹不存在 {folder}")

            # 查找文件
            files = list(folder_path.glob(pattern))
            if not files:
                raise DataError(f"数据操作失败: 在 {folder} 中未找到匹配 '{pattern}' 的文件")

            # 分块处理，每 100 个文件合并一次
            chunk_size = 100
            all_data = []
            file_stats = []

            for i in range(0, len(files), chunk_size):
                chunk_files = files[i : i + chunk_size]
                chunk_data = []

                for file in chunk_files:
                    try:
                        if sheet:
                            df = pd.read_excel(file, sheet_name=sheet)
                            chunk_data.append(df)
                            file_stats.append(
                                {"file": str(file), "rows": len(df), "columns": list(df.columns)}
                            )
                        else:
                            # 读取所有工作表
                            sheets_dict = pd.read_excel(file, sheet_name=None)
                            for sheet_name, df in sheets_dict.items():
                                chunk_data.append(df)
                                file_stats.append(
                                    {
                                        "file": str(file),
                                        "sheet": sheet_name,
                                        "rows": len(df),
                                        "columns": list(df.columns),
                                    }
                                )
                    except Exception as e:
                        logger.warning(f"Failed to read {file}: {e}")
                        file_stats.append({"file": str(file), "error": str(e)})

                if chunk_data:
                    all_data.append(pd.concat(chunk_data, ignore_index=True))

            if not all_data:
                raise DataError("数据操作失败: 无法从任何文件读取数据")

            # 合并所有分块数据
            result = pd.concat(all_data, ignore_index=True)

            # 保存结果
            result.to_excel(output, index=False, sheet_name="MergedData")

            return {
                "success": True,
                "folder": folder,
                "pattern": pattern,
                "file_count": len(files),
                "success_count": sum(1 for s in file_stats if "error" not in s),
                "total_rows": len(result),
                "columns": list(result.columns),
                "output": output,
                "file_stats": file_stats,
            }

        except FileNotFoundError:
            raise
        except DataError:
            raise
        except Exception as e:
            logger.error(f"批量合并失败: {e}")
            raise DataError(f"数据操作失败: {e}")
