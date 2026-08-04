"""测试 validation_engine.py - 数据验证规则引擎"""

import pytest
from pathlib import Path
from openpyxl import Workbook

from abacus.core.balance.validation_engine import (
    ValidationEngineCapability,
    ValidationRule,
    TypeRule,
    RangeRule,
    NotEmptyRule,
    RegexRule,
    CustomRule,
    AndRule,
    OrRule,
    ValidationChain,
    RuleResult,
)
from abacus.core.exceptions import DataError, FileNotFoundError


@pytest.fixture
def validation_excel(tmp_path):
    """创建测试用 Excel 文件"""
    file_path = tmp_path / "validation_test.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "Data"
    ws.append(["ID", "Name", "Amount", "Status", "Email"])
    ws.append([1, "Alice", 100, "active", "alice@test.com"])
    ws.append([2, "Bob", 200, "inactive", "bob@test.com"])
    ws.append([3, "Charlie", -50, "active", "invalid-email"])
    ws.append([4, "David", 150, "active", "david@test.com"])
    wb.save(file_path)
    wb.close()
    return file_path


class TestRuleResult:
    """测试 RuleResult 数据类"""

    def test_rule_result_success(self):
        result = RuleResult(valid=True, cell="A1", rule_name="test")
        assert result.valid is True
        assert result.cell == "A1"
        assert result.errors == []

    def test_rule_result_failure(self):
        result = RuleResult(valid=False, cell="A1", rule_name="test", errors=["error msg"])
        assert result.valid is False
        assert result.errors == ["error msg"]


class TestTypeRule:
    """测试 TypeRule"""

    def test_type_rule_int_pass(self):
        rule = TypeRule("int")
        result = rule.validate(42)
        assert result.valid is True

    def test_type_rule_int_fail(self):
        rule = TypeRule("int")
        result = rule.validate("hello")
        assert result.valid is False
        assert "类型" in result.errors[0] or "type" in result.errors[0].lower()

    def test_type_rule_float_pass(self):
        rule = TypeRule("float")
        result = rule.validate(3.14)
        assert result.valid is True

    def test_type_rule_float_from_int(self):
        rule = TypeRule("float")
        result = rule.validate(42)
        assert result.valid is True

    def test_type_rule_str_pass(self):
        rule = TypeRule("str")
        result = rule.validate("hello")
        assert result.valid is True

    def test_type_rule_date_pass(self):
        from datetime import date
        rule = TypeRule("date")
        result = rule.validate(date(2024, 1, 1))
        assert result.valid is True

    def test_type_rule_name(self):
        rule = TypeRule("int")
        assert rule.name == "type"


class TestRangeRule:
    """测试 RangeRule"""

    def test_range_inclusive(self):
        rule = RangeRule(min_val=0, max_val=100)
        result = rule.validate(50)
        assert result.valid is True

    def test_range_below_min(self):
        rule = RangeRule(min_val=0, max_val=100)
        result = rule.validate(-1)
        assert result.valid is False
        assert len(result.errors) > 0

    def test_range_above_max(self):
        rule = RangeRule(min_val=0, max_val=100)
        result = rule.validate(101)
        assert result.valid is False

    def test_range_no_min(self):
        rule = RangeRule(max_val=100)
        result = rule.validate(-999)
        assert result.valid is True

    def test_range_no_max(self):
        rule = RangeRule(min_val=0)
        result = rule.validate(999999)
        assert result.valid is True

    def test_range_name(self):
        rule = RangeRule(min_val=0, max_val=1)
        assert rule.name == "range"


class TestNotEmptyRule:
    """测试 NotEmptyRule"""

    def test_not_empty_pass(self):
        rule = NotEmptyRule()
        result = rule.validate("hello")
        assert result.valid is True

    def test_not_empty_fail_none(self):
        rule = NotEmptyRule()
        result = rule.validate(None)
        assert result.valid is False

    def test_not_empty_fail_empty_string(self):
        rule = NotEmptyRule()
        result = rule.validate("")
        assert result.valid is False

    def test_not_empty_name(self):
        rule = NotEmptyRule()
        assert rule.name == "not_empty"


class TestRegexRule:
    """测试 RegexRule"""

    def test_regex_match(self):
        rule = RegexRule(pattern=r"^\d+$")
        result = rule.validate("12345")
        assert result.valid is True

    def test_regex_no_match(self):
        rule = RegexRule(pattern=r"^\d+$")
        result = rule.validate("abc")
        assert result.valid is False

    def test_regex_email(self):
        rule = RegexRule(pattern=r"^[\w.-]+@[\w.-]+\.\w+$")
        result = rule.validate("test@example.com")
        assert result.valid is True

    def test_regex_name(self):
        rule = RegexRule(pattern=r".*")
        assert rule.name == "regex"


class TestCustomRule:
    """测试 CustomRule"""

    def test_custom_rule_pass(self):
        rule = CustomRule(
            func=lambda v: v > 0,
            name="positive_check"
        )
        result = rule.validate(42)
        assert result.valid is True

    def test_custom_rule_fail(self):
        rule = CustomRule(
            func=lambda v: v > 0,
            name="positive_check"
        )
        result = rule.validate(-1)
        assert result.valid is False

    def test_custom_rule_with_message(self):
        rule = CustomRule(
            func=lambda v: v > 0,
            name="positive_check",
            error_message="值必须为正数"
        )
        result = rule.validate(-1)
        assert result.valid is False
        assert "正数" in result.errors[0]


class TestAndRule:
    """测试 AndRule 组合"""

    def test_and_both_pass(self):
        rule = AndRule([
            TypeRule("int"),
            RangeRule(min_val=0, max_val=100),
        ])
        result = rule.validate(50)
        assert result.valid is True

    def test_and_first_fails(self):
        rule = AndRule([
            TypeRule("int"),
            RangeRule(min_val=0, max_val=100),
        ])
        result = rule.validate("hello")
        assert result.valid is False

    def test_and_second_fails(self):
        rule = AndRule([
            TypeRule("int"),
            RangeRule(min_val=0, max_val=100),
        ])
        result = rule.validate(200)
        assert result.valid is False

    def test_and_collects_all_errors(self):
        rule = AndRule([
            TypeRule("int"),
            RangeRule(min_val=10, max_val=20),
        ])
        result = rule.validate("abc")
        assert result.valid is False
        assert len(result.errors) >= 1


class TestOrRule:
    """测试 OrRule 组合"""

    def test_or_first_passes(self):
        rule = OrRule([
            TypeRule("int"),
            TypeRule("str"),
        ])
        result = rule.validate(42)
        assert result.valid is True

    def test_or_second_passes(self):
        rule = OrRule([
            TypeRule("int"),
            TypeRule("str"),
        ])
        result = rule.validate("hello")
        assert result.valid is True

    def test_or_both_fail(self):
        rule = OrRule([
            TypeRule("int"),
            RangeRule(min_val=100, max_val=200),
        ])
        result = rule.validate("hello")
        assert result.valid is False

    def test_or_name(self):
        rule = OrRule([TypeRule("int")])
        assert rule.name == "or"


class TestAndRuleName:
    def test_and_name(self):
        rule = AndRule([TypeRule("int")])
        assert rule.name == "and"


class TestValidationChain:
    """测试规则链"""

    def test_chain_sequential(self):
        chain = ValidationChain([
            NotEmptyRule(),
            TypeRule("int"),
            RangeRule(min_val=0, max_val=100),
        ])
        result = chain.validate(50)
        assert result.valid is True

    def test_chain_stops_on_first_failure(self):
        chain = ValidationChain([
            NotEmptyRule(),
            TypeRule("int"),
            RangeRule(min_val=0, max_val=100),
        ])
        result = chain.validate(None)
        assert result.valid is False

    def test_chain_collects_errors(self):
        chain = ValidationChain([
            NotEmptyRule(),
            TypeRule("int"),
        ])
        result = chain.validate(None)
        assert result.valid is False
        assert len(result.errors) >= 1


class TestValidationEngineCapability:
    """测试 ValidationEngineCapability"""

    def test_capability_properties(self):
        cap = ValidationEngineCapability()
        assert cap.name == "validation_engine"
        assert cap.chapter == "balance"
        assert "验证" in cap.description or "validation" in cap.description.lower()

    def test_schema_has_required_params(self):
        cap = ValidationEngineCapability()
        names = [s.name for s in cap.schema]
        assert "file" in names
        assert "sheet" in names
        assert "range" in names
        assert "rules" in names

    def test_missing_file_raises_error(self):
        cap = ValidationEngineCapability()
        with pytest.raises(DataError, match="file"):
            cap.execute(None, sheet="Data", range="A1:A5", rules=[])

    def test_file_not_found_raises_error(self, tmp_path):
        cap = ValidationEngineCapability()
        with pytest.raises(FileNotFoundError):
            cap.execute(
                None,
                file=str(tmp_path / "nonexistent.xlsx"),
                sheet="Data",
                range="A1:A5",
                rules=[],
            )

    def test_sheet_not_found_raises_error(self, validation_excel):
        cap = ValidationEngineCapability()
        with pytest.raises(DataError, match="不存在"):
            cap.execute(
                None,
                file=str(validation_excel),
                sheet="NoSuchSheet",
                range="A1:A5",
                rules=[],
            )

    def test_validate_no_rules(self, validation_excel):
        """无规则时所有单元格通过"""
        cap = ValidationEngineCapability()
        result = cap.execute(
            None,
            file=str(validation_excel),
            sheet="Data",
            range="A1:E5",
            rules=[],
        )
        assert result["valid"] is True
        assert result["total_cells"] == 25

    def test_validate_type_rule(self, validation_excel):
        """类型规则验证"""
        cap = ValidationEngineCapability()
        result = cap.execute(
            None,
            file=str(validation_excel),
            sheet="Data",
            range="A2:A5",
            rules=[{"type": "type", "params": {"expected_type": "int"}}],
        )
        assert result["valid"] is True
        assert result["passed_cells"] == 4

    def test_validate_range_rule(self, validation_excel):
        """范围规则验证"""
        cap = ValidationEngineCapability()
        result = cap.execute(
            None,
            file=str(validation_excel),
            sheet="Data",
            range="C2:C5",
            rules=[{"type": "range", "params": {"min_val": 0, "max_val": 1000}}],
        )
        assert result["valid"] is False
        assert result["failed_cells"] == 1  # Charlie's -50

    def test_validate_not_empty_rule(self, validation_excel):
        """非空规则验证"""
        cap = ValidationEngineCapability()
        result = cap.execute(
            None,
            file=str(validation_excel),
            sheet="Data",
            range="A2:E5",
            rules=[{"type": "not_empty"}],
        )
        assert result["valid"] is True

    def test_validate_regex_rule(self, validation_excel):
        """正则规则验证"""
        cap = ValidationEngineCapability()
        result = cap.execute(
            None,
            file=str(validation_excel),
            sheet="Data",
            range="E2:E5",
            rules=[{"type": "regex", "params": {"pattern": r"^[\w.-]+@[\w.-]+\.\w+$"}}],
        )
        assert result["valid"] is False
        assert result["failed_cells"] == 1  # invalid-email

    def test_validate_and_combination(self, validation_excel):
        """AND 组合规则"""
        cap = ValidationEngineCapability()
        result = cap.execute(
            None,
            file=str(validation_excel),
            sheet="Data",
            range="C2:C5",
            rules=[{
                "type": "and",
                "rules": [
                    {"type": "type", "params": {"expected_type": "int"}},
                    {"type": "range", "params": {"min_val": 0, "max_val": 1000}},
                ],
            }],
        )
        assert result["valid"] is False

    def test_validate_or_combination(self, validation_excel):
        """OR 组合规则"""
        cap = ValidationEngineCapability()
        result = cap.execute(
            None,
            file=str(validation_excel),
            sheet="Data",
            range="A2:A5",
            rules=[{
                "type": "or",
                "rules": [
                    {"type": "type", "params": {"expected_type": "int"}},
                    {"type": "type", "params": {"expected_type": "str"}},
                ],
            }],
        )
        assert result["valid"] is True

    def test_validate_chain(self, validation_excel):
        """规则链验证"""
        cap = ValidationEngineCapability()
        result = cap.execute(
            None,
            file=str(validation_excel),
            sheet="Data",
            range="C2:C5",
            rules=[
                {"type": "not_empty"},
                {"type": "type", "params": {"expected_type": "int"}},
                {"type": "range", "params": {"min_val": 0, "max_val": 1000}},
            ],
        )
        assert result["valid"] is False

    def test_validate_custom_rule(self, validation_excel):
        """自定义规则 - 通过 Python 表达式"""
        cap = ValidationEngineCapability()
        result = cap.execute(
            None,
            file=str(validation_excel),
            sheet="Data",
            range="C2:C5",
            rules=[{"type": "custom", "params": {"expression": "value > 0"}}],
        )
        assert result["valid"] is False

    def test_result_structure(self, validation_excel):
        """验证返回结构"""
        cap = ValidationEngineCapability()
        result = cap.execute(
            None,
            file=str(validation_excel),
            sheet="Data",
            range="A2:A3",
            rules=[{"type": "not_empty"}],
        )
        assert "valid" in result
        assert "total_cells" in result
        assert "passed_cells" in result
        assert "failed_cells" in result
        assert "results" in result
        assert isinstance(result["results"], list)

    def test_cell_results_contain_details(self, validation_excel):
        """单元格结果包含详细信息"""
        cap = ValidationEngineCapability()
        result = cap.execute(
            None,
            file=str(validation_excel),
            sheet="Data",
            range="C2:C3",
            rules=[{"type": "range", "params": {"min_val": 0, "max_val": 100}}],
        )
        for cell_result in result["results"]:
            assert "cell" in cell_result
            assert "valid" in cell_result
            assert "errors" in cell_result
