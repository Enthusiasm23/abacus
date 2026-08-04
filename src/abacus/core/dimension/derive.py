"""少广章 - 推导：深度实现反向推导"""

from typing import Any

from ..base import Capability, CapabilitySchema
from ..exceptions import DataError, ValidationError


class DeriveCapability(Capability):
    @property
    def name(self) -> str:
        return "derive"

    @property
    def chapter(self) -> str:
        return "dimension"

    @property
    def description(self) -> str:
        return "深度反向推导（已知结果求参数）"

    @property
    def schema(self) -> list[CapabilitySchema]:
        return [
            CapabilitySchema(
                name="target_value", type="number", description="目标值", required=True
            ),
            CapabilitySchema(
                name="formula",
                type="string",
                description="公式类型（simple_interest/compound_interest/profit_margin）",
                required=True,
            ),
            CapabilitySchema(name="params", type="object", description="已知参数", required=True),
        ]

    def execute(self, context: Any, **params) -> Any:
        target_value = params.get("target_value")
        formula = params.get("formula")
        known_params = params.get("params", {})

        if target_value is None:
            raise ValidationError("执行失败: 缺少必要参数 target_value")
        if not formula:
            raise ValidationError("执行失败: 缺少必要参数 formula")

        if formula == "simple_interest":
            # I = P * r * t, 求 P/r/t
            if "P" in known_params and "r" in known_params:
                t = target_value / (known_params["P"] * known_params["r"])
                return {"derived": "t", "value": t}
            elif "P" in known_params and "t" in known_params:
                r = target_value / (known_params["P"] * known_params["t"])
                return {"derived": "r", "value": r}
            elif "r" in known_params and "t" in known_params:
                P = target_value / (known_params["r"] * known_params["t"])
                return {"derived": "P", "value": P}
            else:
                raise DataError("数据操作失败: 至少需要 2 个已知参数")

        elif formula == "profit_margin":
            # margin = (price - cost) / price, 求 price/cost
            if "cost" in known_params:
                price = known_params["cost"] / (1 - target_value)
                return {"derived": "price", "value": price}
            elif "price" in known_params:
                cost = known_params["price"] * (1 - target_value)
                return {"derived": "cost", "value": cost}
            else:
                raise DataError("数据操作失败: 需要 cost 或 price 参数")

        else:
            raise DataError(f"数据操作失败: 不支持的公式类型 {formula}")
