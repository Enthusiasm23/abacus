"""报表生成"""

from .advanced import AdvancedReportCapability
from .basic import BasicReportCapability
from .template import TemplateReportCapability

__all__ = ["BasicReportCapability", "AdvancedReportCapability", "TemplateReportCapability"]
