"""少广章 - 求边：深度实现反向计算"""

import math
from typing import Any

from ..base import Capability, CapabilitySchema
from ..exceptions import DataError, ValidationError


class FindDimensionCapability(Capability):
    @property
    def name(self) -> str:
        return "find_dimension"

    @property
    def chapter(self) -> str:
        return "dimension"

    @property
    def description(self) -> str:
        return "深度反向计算（已知面积求边长、已知体积求边长）"

    @property
    def schema(self) -> list[CapabilitySchema]:
        return [
            CapabilitySchema(name="area", type="number", description="面积或体积", required=True),
            CapabilitySchema(
                name="shape",
                type="string",
                description="形状（rectangle/circle/triangle/cube）",
                required=True,
            ),
            CapabilitySchema(
                name="known_side", type="number", description="已知边长（矩形时）", required=False
            ),
        ]

    def execute(self, context: Any, **params) -> Any:
        area = params.get("area")
        shape = params.get("shape")
        known_side = params.get("known_side")

        if area is None:
            raise ValidationError("执行失败: 缺少必要参数 area")
        if not shape:
            raise ValidationError("执行失败: 缺少必要参数 shape")

        if shape == "rectangle":
            if not known_side:
                raise ValidationError("执行失败: 矩形计算需要 known_side 参数")
            other_side = area / known_side
            return {"shape": shape, "area": area, "side1": known_side, "side2": other_side}

        elif shape == "circle":
            radius = math.sqrt(area / math.pi)
            return {"shape": shape, "area": area, "radius": radius, "diameter": radius * 2}

        elif shape == "triangle":
            # 假设等边三角形
            side = math.sqrt(4 * area / math.sqrt(3))
            return {"shape": shape, "area": area, "side": side}

        elif shape == "cube":
            side = area ** (1 / 3)
            return {"shape": shape, "volume": area, "side": side}

        else:
            raise DataError(f"数据操作失败: 不支持的形状类型 {shape}")
