from abacus.core.base import Capability, CapabilitySchema


class TestCapability:
    def test_capability_has_name(self):
        """能力必须有名称"""
        class MyCapability(Capability):
            @property
            def name(self) -> str:
                return "my_cap"
            @property
            def chapter(self) -> str:
                return "test"
            @property
            def description(self) -> str:
                return "Test capability"
            def execute(self, context, **params):
                return {}

        cap = MyCapability()
        assert cap.name == "my_cap"

    def test_capability_has_schema(self):
        """能力可以定义参数 Schema"""
        class MyCapability(Capability):
            @property
            def name(self) -> str:
                return "my_cap"
            @property
            def chapter(self) -> str:
                return "test"
            @property
            def description(self) -> str:
                return "Test capability"
            @property
            def schema(self):
                return [
                    CapabilitySchema(name="file", type="string", description="File path", required=True)
                ]
            def execute(self, context, **params):
                return {}

        cap = MyCapability()
        assert len(cap.schema) == 1
        assert cap.schema[0].name == "file"
