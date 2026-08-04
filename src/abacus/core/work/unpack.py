"""商功章 - 文件解包：将 ZIP 解包为 Excel 文件"""

import logging
import zipfile
from pathlib import Path
from typing import Any

from ..base import Capability, CapabilitySchema
from ..exceptions import DataError, FileNotFoundError

logger = logging.getLogger(__name__)


class UnpackFileCapability(Capability):
    """文件解包：将 ZIP 解包为 Excel 文件"""

    @property
    def name(self) -> str:
        return "unpack_file"

    @property
    def chapter(self) -> str:
        return "work"

    @property
    def description(self) -> str:
        return "将 ZIP 解包为 Excel 文件（用于调试和分析）"

    @property
    def schema(self) -> list[CapabilitySchema]:
        return [
            CapabilitySchema(name="file", type="string", description="ZIP 文件路径", required=True),
            CapabilitySchema(name="output", type="string", description="输出目录", required=False),
        ]

    def execute(self, context: Any, **params) -> Any:
        file_path = params.get("file")
        output = params.get("output")

        if not file_path:
            raise DataError("file parameter is required")

        path = Path(file_path)
        if not path.exists():
            raise FileNotFoundError(f"File not found: {file_path}")

        output_dir = Path(output) if output else path.parent / path.stem
        output_dir.mkdir(parents=True, exist_ok=True)

        try:
            with zipfile.ZipFile(path, "r") as zf:
                zf.extractall(output_dir)

            return {
                "file": file_path,
                "output": str(output_dir),
                "files": list(str(output_dir / f) for f in zf.namelist()),
                "unpacked": True,
            }
        except Exception as e:
            raise DataError(f"Failed to unpack file: {e}")
