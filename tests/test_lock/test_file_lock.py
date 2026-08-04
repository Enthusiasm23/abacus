from pathlib import Path
from abacus.lock.file_lock import FileLock


class TestFileLock:
    def test_lock_path(self):
        lock = FileLock(Path("test.xlsx"))
        assert lock.lock_path == Path("test.xlsx.lock")

    def test_context_manager(self, tmp_path):
        test_file = tmp_path / "test.xlsx"
        test_file.touch()
        lock = FileLock(test_file)
        with lock.acquire():
            pass
