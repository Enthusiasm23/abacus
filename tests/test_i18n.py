"""i18n 模块测试"""

import pytest
from abacus.i18n import set_language, get_language, get_message


class TestI18n:
    def test_default_language(self):
        set_language("zh")
        assert get_language() == "zh"

    def test_set_language_en(self):
        set_language("en")
        assert get_language() == "en"
        set_language("zh")

    def test_get_message_zh(self):
        set_language("zh")
        msg = get_message("file_not_found", file="test.xlsx")
        assert "test.xlsx" in msg
        assert "文件不存在" in msg

    def test_get_message_en(self):
        set_language("en")
        msg = get_message("file_not_found", file="test.xlsx")
        assert "test.xlsx" in msg
        assert "File not found" in msg
        set_language("zh")

    def test_get_message_with_params(self):
        set_language("zh")
        msg = get_message("sheet_not_found", sheet="Sheet1")
        assert "Sheet1" in msg

    def test_get_message_missing_key(self):
        msg = get_message("nonexistent_key")
        assert msg == "nonexistent_key"

    def test_exceptions_use_i18n(self):
        from abacus.core.exceptions import FileNotFoundError, SheetNotFoundError

        set_language("zh")
        with pytest.raises(FileNotFoundError, match="文件不存在"):
            raise FileNotFoundError(file="test.xlsx")

        with pytest.raises(SheetNotFoundError, match="工作表"):
            raise SheetNotFoundError(sheet="Sheet1")

    def test_exceptions_en(self):
        from abacus.core.exceptions import FileNotFoundError, SheetNotFoundError

        set_language("en")
        with pytest.raises(FileNotFoundError, match="File not found"):
            raise FileNotFoundError(file="test.xlsx")

        with pytest.raises(SheetNotFoundError, match="Sheet"):
            raise SheetNotFoundError(sheet="Sheet1")
        set_language("zh")
