from abc import ABC, abstractmethod
from typing import Any

from pydantic import BaseModel


class CapabilitySchema(BaseModel):
    """能力参数 Schema"""

    name: str
    type: str
    description: str
    required: bool = False
    default: Any = None


class Capability(ABC):
    """能力基类"""

    @property
    @abstractmethod
    def name(self) -> str:
        """能力名称"""
        pass

    @property
    @abstractmethod
    def chapter(self) -> str:
        """所属九章"""
        pass

    @property
    @abstractmethod
    def description(self) -> str:
        """能力描述"""
        pass

    @property
    def schema(self) -> list[CapabilitySchema]:
        """参数 Schema"""
        return []

    @abstractmethod
    def execute(self, context: Any, **params) -> Any:
        """执行能力"""
        pass
