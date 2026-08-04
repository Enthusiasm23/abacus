"""日志配置模块 - Abacus 统一日志管理"""

import logging
import os
import sys


def setup_logging(level: str = "INFO") -> None:
    """配置 Abacus 日志

    优先级：环境变量 ABACUS_LOG_LEVEL > 参数 level > 默认 INFO

    Args:
        level: 日志级别（DEBUG, INFO, WARNING, ERROR, CRITICAL）
    """
    env_level = os.environ.get("ABACUS_LOG_LEVEL", "").upper()
    if env_level and env_level in ("DEBUG", "INFO", "WARNING", "ERROR", "CRITICAL"):
        level = env_level

    log_level = getattr(logging, level.upper(), logging.INFO)

    fmt = "%(asctime)s [%(levelname)s] %(name)s: %(message)s"
    datefmt = "%Y-%m-%d %H:%M:%S"

    logging.basicConfig(
        level=log_level,
        format=fmt,
        datefmt=datefmt,
        handlers=[logging.StreamHandler(sys.stderr)],
        force=True,
    )

    logging.getLogger("abacus").setLevel(log_level)
