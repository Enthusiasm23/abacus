"""国际化支持模块"""

from typing import Literal

# 当前语言设置
_current_language: Literal["zh", "en"] = "zh"

# 错误消息映射
ERROR_MESSAGES = {
    "zh": {
        "file_not_found": "文件不存在: {file}",
        "sheet_not_found": "工作表 '{sheet}' 不存在",
        "range_error": "范围错误: {range}",
        "data_error": "数据错误: {message}",
        "formula_error": "公式错误: {message}",
        "validation_error": "验证错误: {message}",
        "invalid_type": "无效的类型: {type}。必须是 {valid_types} 之一",
        "param_required": "参数 {param} 是必需的",
        "conversion_failed": "转换失败: {message}",
        "read_error": "读取失败: {message}",
        "write_error": "写入失败: {message}",
    },
    "en": {
        "file_not_found": "File not found: {file}",
        "sheet_not_found": "Sheet '{sheet}' not found",
        "range_error": "Range error: {range}",
        "data_error": "Data error: {message}",
        "formula_error": "Formula error: {message}",
        "validation_error": "Validation error: {message}",
        "invalid_type": "Invalid type: {type}. Must be one of {valid_types}",
        "param_required": "Parameter {param} is required",
        "conversion_failed": "Conversion failed: {message}",
        "read_error": "Read error: {message}",
        "write_error": "Write error: {message}",
    },
}


def set_language(lang: Literal["zh", "en"]) -> None:
    """设置当前语言"""
    global _current_language
    _current_language = lang


def get_language() -> Literal["zh", "en"]:
    """获取当前语言"""
    return _current_language


def get_message(key: str, **kwargs) -> str:
    """获取翻译消息

    Args:
        key: 消息键名
        **kwargs: 模板参数

    Returns:
        翻译后的消息字符串
    """
    template = ERROR_MESSAGES.get(_current_language, {}).get(key)
    if template is None:
        template = ERROR_MESSAGES["zh"].get(key, key)
    try:
        return template.format(**kwargs)
    except KeyError:
        return template
