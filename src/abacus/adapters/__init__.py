"""File adapter layer."""

from .base import ExcelAdapter
from .xlsx import XlsxAdapter

__all__ = ["ExcelAdapter", "XlsxAdapter"]
