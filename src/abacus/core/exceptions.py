"""Abacus 异常定义"""

from ..i18n import get_message


class AbacusError(Exception):
    """基础异常"""

    def __init__(self, message: str = None, key: str = None, **kwargs):
        if key:
            message = get_message(key, **kwargs)
        super().__init__(message or "Abacus error")


class ValidationError(AbacusError):
    """验证错误"""

    def __init__(self, message: str = None, key: str = None, **kwargs):
        if key:
            message = get_message(key, **kwargs)
        super().__init__(message or get_message("validation_error", message=""))


class FileNotFoundError(AbacusError):
    """文件不存在"""

    def __init__(self, file: str = None, message: str = None, key: str = None, **kwargs):
        if key:
            message = get_message(key, **kwargs)
        elif file:
            message = get_message("file_not_found", file=file)
        super().__init__(message or "File not found")


class SheetNotFoundError(AbacusError):
    """工作表不存在"""

    def __init__(self, sheet: str = None, message: str = None, key: str = None, **kwargs):
        if key:
            message = get_message(key, **kwargs)
        elif sheet:
            message = get_message("sheet_not_found", sheet=sheet)
        super().__init__(message or "Sheet not found")


class RangeError(AbacusError):
    """范围错误"""

    def __init__(self, range: str = None, message: str = None, key: str = None, **kwargs):
        if key:
            message = get_message(key, **kwargs)
        elif range:
            message = get_message("range_error", range=range)
        super().__init__(message or "Range error")


class DataError(AbacusError):
    """数据错误"""

    def __init__(self, message: str = None, key: str = None, **kwargs):
        if key:
            message = get_message(key, **kwargs)
        super().__init__(message or get_message("data_error", message=message or ""))


class FormulaError(AbacusError):
    """公式错误"""

    def __init__(self, message: str = None, key: str = None, **kwargs):
        if key:
            message = get_message(key, **kwargs)
        super().__init__(message or get_message("formula_error", message=message or ""))
