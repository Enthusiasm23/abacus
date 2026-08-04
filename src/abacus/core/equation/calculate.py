"""方程章 - 计算：深度实现计算"""

import math
from typing import Any

from ..base import Capability, CapabilitySchema
from ..exceptions import DataError, ValidationError


class CalculateCapability(Capability):
    @property
    def name(self) -> str:
        return "calculate"

    @property
    def chapter(self) -> str:
        return "equation"

    @property
    def description(self) -> str:
        return "深度执行计算（数学表达式、函数）"

    @property
    def schema(self) -> list[CapabilitySchema]:
        return [
            CapabilitySchema(
                name="expression", type="string", description="数学表达式", required=True
            ),
            CapabilitySchema(
                name="variables", type="object", description="变量值（可选）", required=False
            ),
        ]

    def execute(self, context: Any, **params) -> Any:
        expression = params.get("expression")
        variables = params.get("variables", {})

        if not expression:
            raise ValidationError("执行失败: 缺少必要参数 expression")

        # 安全的数学函数
        safe_dict = {
            "abs": abs,
            "round": round,
            "min": min,
            "max": max,
            "sqrt": math.sqrt,
            "pow": pow,
            "log": math.log,
            "log10": math.log10,
            "sin": math.sin,
            "cos": math.cos,
            "tan": math.tan,
            "pi": math.pi,
            "e": math.e,
        }
        safe_dict.update(variables)

        try:
            result = eval(expression, {"__builtins__": {}}, safe_dict)
            return {"expression": expression, "variables": variables, "result": result}
        except Exception as e:
            raise DataError(f"数据操作失败: 表达式计算失败 {e}")
