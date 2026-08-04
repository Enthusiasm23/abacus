"""方程章 - 解方程：深度实现方程求解"""

import re
from typing import Any

from ..base import Capability, CapabilitySchema
from ..exceptions import DataError, ValidationError


class SolveEquationCapability(Capability):
    @property
    def name(self) -> str:
        return "solve_equation"

    @property
    def chapter(self) -> str:
        return "equation"

    @property
    def description(self) -> str:
        return "深度解方程（一元一次、一元二次）"

    @property
    def schema(self) -> list[CapabilitySchema]:
        return [
            CapabilitySchema(
                name="equation",
                type="string",
                description="方程（如 2x+3=7, x^2-4=0）",
                required=True,
            ),
        ]

    def execute(self, context: Any, **params) -> Any:
        equation = params.get("equation")

        if not equation:
            raise ValidationError("执行失败: 缺少必要参数 equation")

        # 简单解析：ax+b=c 形式
        match = re.match(r"([+-]?\d*)x\s*([+-]\s*\d+)\s*=\s*(\d+)", equation.replace(" ", ""))
        if match:
            a_str, b_str, c_str = match.groups()
            a = (
                int(a_str)
                if a_str and a_str != "+" and a_str != "-"
                else (1 if a_str != "-" else -1)
            )
            b = int(b_str.replace(" ", ""))
            c = int(c_str)

            x = (c - b) / a
            return {"equation": equation, "type": "linear", "solution": x}

        # x^2 + bx + c = 0 形式
        match = re.match(r"x\^2\s*([+-]\s*\d+)x\s*([+-]\s*\d+)\s*=\s*0", equation.replace(" ", ""))
        if match:
            import math

            b_str, c_str = match.groups()
            b = int(b_str.replace(" ", ""))
            c = int(c_str.replace(" ", ""))

            discriminant = b**2 - 4 * c
            if discriminant < 0:
                return {"equation": equation, "type": "quadratic", "solution": "no real solution"}
            elif discriminant == 0:
                x = -b / 2
                return {"equation": equation, "type": "quadratic", "solution": x}
            else:
                x1 = (-b + math.sqrt(discriminant)) / 2
                x2 = (-b - math.sqrt(discriminant)) / 2
                return {"equation": equation, "type": "quadratic", "solutions": [x1, x2]}

        # x^2 + c = 0 形式（无 bx 项）
        match = re.match(r"x\^2\s*([+-]\s*\d+)\s*=\s*0", equation.replace(" ", ""))
        if match:
            import math

            c_str = match.groups()[0]
            c = int(c_str.replace(" ", ""))

            if c > 0:
                return {"equation": equation, "type": "quadratic", "solution": "no real solution"}
            elif c == 0:
                return {"equation": equation, "type": "quadratic", "solution": 0}
            else:
                x1 = math.sqrt(-c)
                x2 = -math.sqrt(-c)
                return {"equation": equation, "type": "quadratic", "solutions": [x1, x2]}

        raise DataError(f"数据操作失败: 不支持的方程格式 {equation}")
