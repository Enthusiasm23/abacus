from abc import ABC, abstractmethod
from pathlib import Path
from typing import Any


class ExcelAdapter(ABC):
    """Excel 适配器基类"""

    @property
    @abstractmethod
    def supported_formats(self) -> list[str]:
        """支持的文件格式"""

    @abstractmethod
    def open(self, file_path: Path) -> Any:
        """打开工作簿"""

    @abstractmethod
    def save(self, workbook: Any, file_path: Path) -> None:
        """保存工作簿"""
