from abacus.core.base import Capability
from abacus.core.registry import CapabilityRegistry


class MockCapability(Capability):
    @property
    def name(self) -> str:
        return "mock"

    @property
    def chapter(self) -> str:
        return "test"

    @property
    def description(self) -> str:
        return "Mock capability"

    def execute(self, context, **params):
        return {"result": "ok"}


class TestCapabilityRegistry:
    def test_register_and_get(self):
        registry = CapabilityRegistry()
        cap = MockCapability()
        registry.register(cap)

        assert registry.get("mock") == cap
        assert registry.get("nonexistent") is None

    def test_list_all(self):
        registry = CapabilityRegistry()
        cap = MockCapability()
        registry.register(cap)

        assert len(registry.list_all()) == 1

    def test_list_by_chapter(self):
        registry = CapabilityRegistry()
        cap = MockCapability()
        registry.register(cap)

        result = registry.list_by_chapter("test")
        assert len(result) == 1

        result = registry.list_by_chapter("other")
        assert len(result) == 0
