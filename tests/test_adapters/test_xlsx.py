import pytest
from pathlib import Path
from abacus.adapters.xlsx import XlsxAdapter
from abacus.core.exceptions import FileNotFoundError as AbacusFileNotFoundError


class TestXlsxAdapter:
    def test_supported_formats(self):
        adapter = XlsxAdapter()
        assert ".xlsx" in adapter.supported_formats

    def test_open_nonexistent_file(self):
        adapter = XlsxAdapter()
        with pytest.raises(AbacusFileNotFoundError):
            adapter.open(Path("nonexistent.xlsx"))
