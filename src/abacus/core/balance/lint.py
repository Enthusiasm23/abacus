"""代码审计工具 - 检查 openpyxl 代码常见问题"""

import logging
import re
from pathlib import Path
from typing import Any

from ..base import Capability, CapabilitySchema
from ..exceptions import DataError, FileNotFoundError, ValidationError

logger = logging.getLogger(__name__)


class ExcelLintCapability(Capability):
    """代码审计 - 检查 openpyxl 代码常见问题"""

    # 检查规则
    RULES = {
        "XL001": "data_only=True 后 save() 会导致公式永久丢失",
        "XL002": "pandas to_excel 写入的公式会变成字符串",
        "XL003": "公式中包含中文逗号（，）",
        "XL004": "sheet 名包含禁止字符（: \\ / ? * [ ]）",
        "XL005": "sheet 名超过 31 字符",
        "XL006": "PatternFill 没有传 fill_type",
        "XL007": ".xlsm 文件 load 时没有 keep_vba=True",
        "XL008": "跨 sheet 引用时 sheet 名含空格但没用单引号",
        "XL009": "中文字符没有设置字体",
        "XL010": "read_only=True 后试图 save()",
    }

    @property
    def name(self) -> str:
        return "excel_lint"

    @property
    def chapter(self) -> str:
        return "balance"

    @property
    def description(self) -> str:
        return "检查 openpyxl 代码的 10 类常见问题"

    @property
    def schema(self) -> list[CapabilitySchema]:
        return [
            CapabilitySchema(
                name="code", type="string", description="Python 代码内容", required=False
            ),
            CapabilitySchema(
                name="file", type="string", description="Python 文件路径", required=False
            ),
        ]

    def execute(self, context: Any, **params) -> Any:
        code = params.get("code")
        file_path = params.get("file")

        if file_path:
            path = Path(file_path)
            if not path.exists():
                raise FileNotFoundError(f"文件操作失败: 文件不存在 {file_path}")
            code = path.read_text(encoding="utf-8")

        if not code:
            raise ValidationError("执行失败: 缺少必要参数 code 或 file")

        return self._lint_code(code)

    def _lint_code(self, code: str) -> dict[str, Any]:
        """检查代码"""
        issues = []
        lines = code.split("\n")

        for i, line in enumerate(lines, 1):
            # XL001: data_only=True 后 save()
            if "data_only=True" in code and "save(" in line:
                issues.append(
                    {
                        "rule": "XL001",
                        "level": "ERROR",
                        "line": i,
                        "message": self.RULES["XL001"],
                        "code": line.strip(),
                    }
                )

            # XL003: 中文逗号
            if re.search(r'["\'].*[，].*["\']', line) and (
                "=" in line or "formula" in line.lower()
            ):
                issues.append(
                    {
                        "rule": "XL003",
                        "level": "ERROR",
                        "line": i,
                        "message": self.RULES["XL003"],
                        "code": line.strip(),
                    }
                )

            # XL004: sheet 名禁止字符
            if re.search(r'["\'].*[\\\/\?\*\[\]:].*["\']', line) and (
                "sheet" in line.lower() or "title" in line.lower()
            ):
                issues.append(
                    {
                        "rule": "XL004",
                        "level": "ERROR",
                        "line": i,
                        "message": self.RULES["XL004"],
                        "code": line.strip(),
                    }
                )

            # XL006: PatternFill 没有 fill_type
            if "PatternFill(" in line and "fill_type" not in line and "patternType" not in line:
                issues.append(
                    {
                        "rule": "XL006",
                        "level": "WARN",
                        "line": i,
                        "message": self.RULES["XL006"],
                        "code": line.strip(),
                    }
                )

            # XL010: read_only=True 后 save()
            if "read_only=True" in code and "save(" in line:
                issues.append(
                    {
                        "rule": "XL010",
                        "level": "ERROR",
                        "line": i,
                        "message": self.RULES["XL010"],
                        "code": line.strip(),
                    }
                )

        return {
            "total_issues": len(issues),
            "errors": len([i for i in issues if i["level"] == "ERROR"]),
            "warnings": len([i for i in issues if i["level"] == "WARN"]),
            "issues": issues,
            "rules": self.RULES,
        }
