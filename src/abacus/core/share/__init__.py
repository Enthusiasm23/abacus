"""衰分章：分组汇总"""

from .distribute import DistributeCapability
from .group_by import GroupByCapability
from .subtotal import SubtotalCapability
from .summarize import SummarizeCapability

__all__ = ["GroupByCapability", "DistributeCapability", "SummarizeCapability", "SubtotalCapability"]
