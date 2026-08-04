"""商功章 - 变化检测：对比两个版本的数据，检测变化"""

import logging
from pathlib import Path
from typing import Any

import pandas as pd

from ..base import Capability, CapabilitySchema
from ..exceptions import DataError, FileNotFoundError

logger = logging.getLogger(__name__)


class DiffReportCapability(Capability):
    """变化检测：对比两个版本的数据，检测变化"""

    @property
    def name(self) -> str:
        return "generate_diff_report"

    @property
    def chapter(self) -> str:
        return "work"

    @property
    def description(self) -> str:
        return "对比两个版本的数据，检测变化"

    @property
    def schema(self) -> list[CapabilitySchema]:
        return [
            CapabilitySchema(
                name="old_file", type="string", description="旧版本文件路径", required=True
            ),
            CapabilitySchema(
                name="old_sheet", type="string", description="旧版本工作表名称", required=True
            ),
            CapabilitySchema(
                name="new_file", type="string", description="新版本文件路径", required=True
            ),
            CapabilitySchema(
                name="new_sheet", type="string", description="新版本工作表名称", required=True
            ),
            CapabilitySchema(
                name="key_columns",
                type="array",
                description="用于匹配的键列（可选）",
                required=False,
            ),
        ]

    def execute(self, context: Any, **params) -> Any:
        """生成变化检测报告"""
        old_file = params.get("old_file")
        old_sheet = params.get("old_sheet")
        new_file = params.get("new_file")
        new_sheet = params.get("new_sheet")
        key_columns = params.get("key_columns")

        if not old_file:
            raise DataError("old_file parameter is required")
        if not new_file:
            raise DataError("new_file parameter is required")

        return self._generate_diff_report(old_file, old_sheet, new_file, new_sheet, key_columns)

    def _generate_diff_report(
        self,
        old_file: str,
        old_sheet: str,
        new_file: str,
        new_sheet: str,
        key_columns: list[str] = None,
    ) -> dict[str, Any]:
        """生成变化检测报告"""
        try:
            old_path = Path(old_file)
            new_path = Path(new_file)

            if not old_path.exists():
                raise FileNotFoundError(f"Old file not found: {old_file}")
            if not new_path.exists():
                raise FileNotFoundError(f"New file not found: {new_file}")

            # 读取数据
            df_old = pd.read_excel(old_path, sheet_name=old_sheet)
            df_new = pd.read_excel(new_path, sheet_name=new_sheet)

            # 基本统计
            diff = {
                "success": True,
                "old_file": old_file,
                "new_file": new_file,
                "old_rows": len(df_old),
                "new_rows": len(df_new),
                "row_diff": len(df_new) - len(df_old),
                "old_columns": (df_old.columns),
                "new_columns": (df_new.columns),
                "added_columns": list(c for c in df_new.columns if c not in df_old.columns),
                "removed_columns": list(c for c in df_old.columns if c not in df_new.columns),
                "changes": [],
            }

            if key_columns:
                # 基于键的详细对比
                for key in key_columns:
                    if key not in df_old.columns or key not in df_new.columns:
                        raise DataError(f"Key column '{key}' not found in both tables")

                # 合并数据进行对比
                merged = pd.merge(
                    df_old,
                    df_new,
                    on=key_columns,
                    how="outer",
                    suffixes=("_old", "_new"),
                    indicator=True,
                )

                # 统计变化
                added = merged[merged["_merge"] == "right_only"]
                removed = merged[merged["_merge"] == "left_only"]
                modified = merged[merged["_merge"] == "both"]

                diff["changes"].append(
                    {
                        "type": "summary",
                        "added_rows": len(added),
                        "removed_rows": len(removed),
                        "modified_rows": len(modified),
                    }
                )
            else:
                # 简单的行列对比
                if len(df_new.columns) != len(df_old.columns):
                    diff["changes"].append(
                        {
                            "type": "column_count",
                            "old_count": len(df_old.columns),
                            "new_count": len(df_new.columns),
                        }
                    )

            return diff

        except FileNotFoundError:
            raise
        except DataError:
            raise
        except Exception as e:
            logger.error(f"Failed to generate diff report: {e}")
            raise DataError(str(e))
