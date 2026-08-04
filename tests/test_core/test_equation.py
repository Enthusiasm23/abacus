from abacus.core.equation.create_formula import CreateFormulaCapability


class TestCreateFormula:
    def test_capability_properties(self):
        cap = CreateFormulaCapability()
        assert cap.name == "create_formula"
        assert cap.chapter == "equation"