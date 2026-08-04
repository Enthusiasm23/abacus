"""商功章 - 数据视图：创建和管理不同角色的数据视图"""

import json
import logging
from pathlib import Path
from typing import Any

import pandas as pd

from ..base import Capability, CapabilitySchema
from ..exceptions import DataError, FileNotFoundError

logger = logging.getLogger(__name__)


class DataViewCapability(Capability):
    """数据视图：创建和管理不同角色的数据视图"""

    @property
    def name(self) -> str:
        return "manage_data_view"

    @property
    def chapter(self) -> str:
        return "work"

    @property
    def description(self) -> str:
        return "创建和管理不同角色的数据视图"

    @property
    def schema(self) -> list[CapabilitySchema]:
        return [
            CapabilitySchema(
                name="file", type="string", description="Excel 文件路径", required=True
            ),
            CapabilitySchema(name="sheet", type="string", description="工作表名称", required=True),
            CapabilitySchema(
                name="action", type="string", description="操作：create//get/delete", required=True
            ),
            CapabilitySchema(
                name="view_name", type="string", description="视图名称", required=False
            ),
            CapabilitySchema(
                name="columns", type="array", description="视图包含的列", required=False
            ),
            CapabilitySchema(name="filters", type="object", description="过滤条件", required=False),
        ]

    def execute(self, context: Any, **params) -> Any:
        """管理数据视图"""
        file_path = params.get("file")
        sheet_name = params.get("sheet")
        action = params.get("action")
        view_name = params.get("view_name")
        columns = params.get("columns")
        filters = params.get("filters")

        if not file_path:
            raise DataError("file parameter is required")
        if not sheet_name:
            raise DataError("sheet parameter is required")
        if not action:
            raise DataError("action parameter is required")

        return self._manage_data_view(file_path, sheet_name, action, view_name, columns, filters)

    def _manage_data_view(
        self,
        filepath: str,
        sheet_name: str,
        action: str,
        view_name: str = None,
        columns: list[str] = None,
        filters: dict[str, Any] = None,
    ) -> dict[str, Any]:
        """管理数据视图"""
        try:
            path = Path(filepath)
            if not path.exists():
                raise FileNotFoundError(f"File not found: {filepath}")

            # 视图配置文件
            config_file = path.parent / f"{path.stem}_views.json"

            # 加载现有视图配置
            views = {}
            if config_file.exists():
                with open(config_file, "r", encoding="utf-8") as f:
                    views = json.load(f)

            if action == "create":
                if not view_name:
                    raise DataError("view_name parameter is required for create action")
                if not columns:
                    raise DataError("columns parameter is required for create action")

                views[view_name] = {
                    "sheet": sheet_name,
                    "columns": columns,
                    "filters": filters or {},
                }

                # 保存配置
                with open(config_file, "w", encoding="utf-8") as f:
                    json.dump(views, f, ensure_ascii=False, indent=2)

                return {
                    "success": True,
                    "action": "create",
                    "view_name": view_name,
                    "columns": columns,
                    "filters": filters,
                }

            elif action == "list":
                return {
                    "success": True,
                    "action": "list",
                    "views": (views.keys()),
                    "count": len(views),
                }

            elif action == "get":
                if not view_name:
                    raise DataError("view_name parameter is required for get action")

                if view_name not in views:
                    raise DataError(f"View '{view_name}' not found")

                view_config = views[view_name]

                # 读取数据并应用视图
                df = pd.read_excel(path, sheet_name=view_config["sheet"])

                # 应用列选择
                if view_config["columns"]:
                    available_cols = [c for c in view_config["columns"] if c in df.columns]
                    df = df[available_cols]

                # 应用过滤
                if view_config.get("filters"):
                    for col, value in view_config["filters"].items():
                        if col in df.columns:
                            if isinstance(value):
                                df = df[df[col].isin(value)]
                            else:
                                df = df[df[col] == value]

                return {
                    "success": True,
                    "action": "get",
                    "view_name": view_name,
                    "config": view_config,
                    "rows": len(df),
                    "columns": (df.columns),
                    "data": df.head(100).to_dict(orient="records"),
                }

            elif action == "delete":
                if not view_name:
                    raise DataError("view_name parameter is required for delete action")

                if view_name not in views:
                    raise DataError(f"View '{view_name}' not found")

                del views[view_name]

                # 保存配置
                with open(config_file, "w", encoding="utf-8") as f:
                    json.dump(views, f, ensure_ascii=False, indent=2)

                return {"success": True, "action": "delete", "view_name": view_name}

            else:
                raise DataError(f"Invalid action: {action}. Use 'create', '', 'get', or 'delete'")

        except FileNotFoundError:
            raise
        except DataError:
            raise
        except Exception as e:
            logger.error(f"Failed to manage data view: {e}")
            raise DataError(str(e))
