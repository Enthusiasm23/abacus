from pathlib import Path
from typing import Any

import openpyxl

from ..core.exceptions import FileNotFoundError
from .base import ExcelAdapter


class XlsxAdapter(ExcelAdapter):
    """openpyxl 适配器"""

    @property
    def supported_formats(self) -> list[str]:
        return [".xlsx", ".xlsm"]

    def open(self, file_path: Path) -> Any:
        if not file_path.exists():
            raise FileNotFoundError(f"文件操作失败: 文件不存在 {file_path}")
        return openpyxl.load_workbook(file_path)

    def save(self, workbook: Any, file_path: Path) -> None:
        workbook.save(file_path)
