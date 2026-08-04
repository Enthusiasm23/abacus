from contextlib import contextmanager
from pathlib import Path

import portalocker


class FileLock:
    """Excel 文件锁（Windows 兼容）"""

    def __init__(self, file_path: Path, timeout: float = 30.0):
        self.file_path = file_path
        self.lock_path = file_path.with_suffix(file_path.suffix + ".lock")
        self.timeout = timeout

    @contextmanager
    def acquire(self):
        lock = portalocker.Lock(
            str(self.lock_path), timeout=self.timeout, flags=portalocker.LOCK_EX
        )
        try:
            lock.acquire()
            yield
        finally:
            lock.release()
