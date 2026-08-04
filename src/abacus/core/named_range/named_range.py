"""命名范围操作"""

import logging
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.workbook.defined_name import DefinedName

from ..base import Capability, CapabilitySchema
from ..exceptions import DataError, FileNotFoundError

logger = logging.getLogger(__name__)


class NamedRangeCapability(Capability):
    """命名范围管理"""

    @property
    def name(self) -> str:
        return "manage_named_range"

    @property
    def chapter(self) -> str:
        return "field"

    @property
    def description(self) -> str:
        return "管理命名范围（创建/列出/读取/删除）"

    @property
    def schema(self) -> list[CapabilitySchema]:
        return [
            CapabilitySchema(
                name="file", type="string", description="Excel 文件路径", required=True
            ),
            CapabilitySchema(
                name="action",
                type="string",
                description="操作（create//read/delete）",
                required=True,
            ),
            CapabilitySchema(
                name="name", type="string", description="命名范围名称", required=False
            ),
            CapabilitySchema(
                name="refers_to",
                type="string",
                description="引用位置（如 Sheet1!$A$1:$D$10）",
                required=False,
            ),
        ]

    def execute(self, context: Any, **params) -> Any:
        file_path = params.get("file")
        action = params.get("action")
        name = params.get("name")
        refers_to = params.get("refers_to")

        if not file_path:
            raise DataError("file parameter is required")
        if not action:
            raise DataError("action parameter is required")

        try:
            path = Path(file_path)
            if not path.exists():
                raise FileNotFoundError(f"File not found: {file_path}")

            wb = load_workbook(path)

            if action == "create":
                if not name or not refers_to:
                    raise DataError("name and refers_to required for create action")

                dn = DefinedName(name, refers_to)
                wb.defined_names.add(dn)
                result = {"action": "create", "name": name, "refers_to": refers_to}

            elif action == "list":
                names = []
                for dn in wb.defined_names.values():
                    names.append({"name": dn.name, "refers_to": dn.attr_text})
                result = {"action": "list", "named_ranges": names}

            elif action == "read":
                if not name:
                    raise DataError("name required for read action")

                if name in wb.defined_names:
                    dn = wb.defined_names[name]
                    result = {"action": "read", "name": name, "refers_to": dn.attr_text}
                else:
                    raise DataError(f"Named range '{name}' not found")

            elif action == "delete":
                if not name:
                    raise DataError("name required for delete action")

                if name in wb.defined_names:
                    del wb.defined_names[name]
                    result = {"action": "delete", "name": name}
                else:
                    raise DataError(f"Named range '{name}' not found")

            else:
                raise DataError(f"Unknown action: {action}")

            wb.save(file_path)
            wb.close()

            return result

        except Exception as e:
            logger.error(f"Failed to manage named range: {e}")
            raise DataError(str(e))
