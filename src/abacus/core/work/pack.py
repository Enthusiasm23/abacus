"""商功章 - 文件打包：将 Excel 文件打包为 ZIP"""

import logging
import zipfile
from pathlib import Path
from typing import Any

from ..base import Capability, CapabilitySchema
from ..exceptions import DataError, FileNotFoundError

logger = logging.getLogger(__name__)


class PackFileCapability(Capability):
    """文件打包：将 Excel 文件打包为 ZIP"""

    @property
    def name(self) -> str:
        return "pack_file"

    @property
    def chapter(self) -> str:
        return "work"

    @property
    def description(self) -> str:
        return "将 Excel 文件打包为 ZIP（用于调试和分析）"

    @property
    def schema(self) -> list[CapabilitySchema]:
        return [
            CapabilitySchema(
                name="file", type="string", description="Excel 文件路径", required=True
            ),
            CapabilitySchema(
                name="output", type="string", description="输出 ZIP 文件路径", required=False
            ),
        ]

    def execute(self, context: Any, **params) -> Any:
        file_path = params.get("file")
        output = params.get("output")

        if not file_path:
            raise DataError("file parameter is required")

        path = Path(file_path)
        if not path.exists():
            raise FileNotFoundError(f"File not found: {file_path}")

        output_path = output or str(path.with_suffix(".zip"))

        try:
            with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zf:
                zf.write(path, path.name)

            return {"file": file_path, "output": output_path, "packed": True}
        except Exception as e:
            raise DataError(f"Failed to pack file: {e}")
