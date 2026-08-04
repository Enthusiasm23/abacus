import pytest
from abacus.core.cell_utils import (
    parse_cell_reference,
    parse_range,
    validate_cell_reference,
)
from abacus.core.exceptions import RangeError


class TestParseCellReference:
    def test_simple_cell(self):
        row, col = parse_cell_reference("A1")
        assert row == 1
        assert col == 1
    
    def test_multi_letter_column(self):
        row, col = parse_cell_reference("AA1")
        assert row == 1
        assert col == 27
    
    def test_large_row(self):
        row, col = parse_cell_reference("A100")
        assert row == 100
        assert col == 1
    
    def test_invalid_format(self):
        with pytest.raises(RangeError):
            parse_cell_reference("1A")


class TestParseRange:
    def test_single_cell(self):
        start_row, start_col, end_row, end_col = parse_range("A1")
        assert start_row == 1
        assert start_col == 1
        assert end_row is None
        assert end_col is None
    
    def test_standard_range(self):
        start_row, start_col, end_row, end_col = parse_range("A1:D10")
        assert start_row == 1
        assert start_col == 1
        assert end_row == 10
        assert end_col == 4
    
    def test_column_range(self):
        start_row, start_col, end_row, end_col = parse_range("A:D")
        assert start_row == 1
        assert start_col == 1
        assert end_row is None
        assert end_col == 4


class TestValidateCellReference:
    def test_valid(self):
        assert validate_cell_reference("A1") is True
        assert validate_cell_reference("AA100") is True
    
    def test_invalid(self):
        assert validate_cell_reference("") is False
        assert validate_cell_reference("1A") is False
        assert validate_cell_reference("A") is False
