from abacus.core.work.advanced_filter import AdvancedFilterCapability


class TestAdvancedFilter:
    def test_capability_properties(self):
        cap = AdvancedFilterCapability()
        assert cap.name == "advanced_filter"
        assert cap.chapter == "work"
        assert "高级筛选" in cap.description
