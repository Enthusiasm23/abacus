from typing import TYPE_CHECKING, Optional

if TYPE_CHECKING:
    from .base import Capability


class CapabilityRegistry:
    """能力注册表"""

    def __init__(self) -> None:
        self._capabilities: dict[str, "Capability"] = {}

    def register(self, capability: "Capability") -> None:
        """注册能力"""
        self._capabilities[capability.name] = capability

    def get(self, name: str) -> Optional["Capability"]:
        """获取能力"""
        return self._capabilities.get(name)

    def list_all(self) -> list["Capability"]:
        """列出所有能力"""
        return list(self._capabilities.values())

    def list_by_chapter(self, chapter: str) -> list["Capability"]:
        """按九章列出能力"""
        return [c for c in self._capabilities.values() if c.chapter == chapter]
