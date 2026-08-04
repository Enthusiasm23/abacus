"""Abacus - 本地 Excel 自动化框架

用户友好导入方式:
    from abacus import MeasureRangeCapability
    from abacus import CapabilityRegistry
    from abacus import Capability, CapabilitySchema
"""

__version__ = "1.0.0"

# 从 core 导入所有能力到顶层
from abacus.core import *

# 导入核心类
from abacus.core.base import Capability, CapabilitySchema
from abacus.core.registry import CapabilityRegistry
