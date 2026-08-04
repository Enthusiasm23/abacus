from abacus.core.work.batch_execute import BatchExecuteCapability


class TestBatchExecute:
    def test_capability_properties(self):
        cap = BatchExecuteCapability()
        assert cap.name == "batch_execute"
        assert cap.chapter == "work"