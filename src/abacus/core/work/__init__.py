"""商功章：批量操作"""

from ..table import TableCapability
from .advanced_chart import AdvancedChartCapability
from .auto_filter import AutoFilterCapability
from .batch_execute import BatchExecuteCapability
from .batch_transform import BatchTransformCapability
from .batch_validate import BatchValidateCapability
from .chart import (
    CreateChartCapability,
    DeleteChartCapability,
    ListChartsCapability,
    UpdateChartCapability,
)
from .comment import CommentCapability
from .data_view import DataViewCapability
from .diff_report import DiffReportCapability
from .export_chart import ExportChartAsImageCapability
from .format import FormatRangeCapability
from .freeze import FreezePaneCapability
from .group_rows import GroupRowsCapability
from .hide_show import HideShowCapability
from .image import InsertImageCapability
from .pack import PackFileCapability
from .pivot import CreatePivotCapability
from .print_area import PrintAreaCapability
from .protection import (
    ProtectSheetCapability,
    ProtectWorkbookCapability,
    SetArrayFormulaCapability,
    UnprotectSheetCapability,
)
from .summary_report import SummaryReportCapability
from .unpack import UnpackFileCapability
from .zoom import ZoomCapability

__all__ = [
    "BatchExecuteCapability",
    "BatchTransformCapability",
    "BatchValidateCapability",
    "CreatePivotCapability",
    "FormatRangeCapability",
    "CreateChartCapability",
    "UpdateChartCapability",
    "ListChartsCapability",
    "DeleteChartCapability",
    "TableCapability",
    "CommentCapability",
    "FreezePaneCapability",
    "AutoFilterCapability",
    "HideShowCapability",
    "AdvancedChartCapability",
    "ProtectWorkbookCapability",
    "ProtectSheetCapability",
    "UnprotectSheetCapability",
    "SetArrayFormulaCapability",
    "InsertImageCapability",
    "GroupRowsCapability",
    "ExportChartAsImageCapability",
    "PackFileCapability",
    "UnpackFileCapability",
    "PrintAreaCapability",
    "ZoomCapability",
    "SummaryReportCapability",
    "DiffReportCapability",
    "DataViewCapability",
]
