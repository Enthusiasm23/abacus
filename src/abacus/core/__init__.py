"""Core capability layer."""

# 数据分析模块
from .analysis import (
    AdvancedAnalysisCapability,
    DataAnalysisCapability,
    DataCleaningCapability,
    PivotAnalysisCapability,
)

# 盈不足章
from .balance import (
    DataValidationCapability,
    ExcelLintCapability,
    FileAnalyzeCapability,
    FileValidateCapability,
    QualityCheckCapability,
    ValidateFormulaCapability,
    ValidateRangeCapability,
    ValidateTypeCapability,
)
from .base import Capability, CapabilitySchema
from .cell_utils import *

# 格式转换
from .conversion import ExcelToMarkdownCapability, SplitSheetCapability

# CSV 处理
from .csv import CSVMergeCapability, CSVVisualizeCapability

# 少广章
from .dimension import (
    AutoSumCapability,
    CalculateCapability,
    DeriveCapability,
    FindDimensionCapability,
    SolveEquationCapability,
)

# 方程章
from .equation import CreateFormulaCapability, DiagnoseFormulaCapability, FormulaRecalcCapability
from .exceptions import *

# 方田章
from .field import (
    DetectColumnsCapability,
    GetSampleDataCapability,
    GetSummaryCapability,
    ListSheetsCapability,
    MeasureCellsCapability,
    MeasureRangeCapability,
    MeasureStructureCapability,
    PeekPreviewCapability,
    SearchContentCapability,
)

# 金融建模
from .finance import VarianceCapability

# 公式生成器
from .formula import FormulaGeneratorCapability

# 粟米章
from .grain import (
    AutoTypeInferCapability,
    ConvertFormatCapability,
    ConvertTypeCapability,
    ConvertUnitCapability,
    DataTransformCapability,
    FuzzyMatchCapability,
    StandardizeCapability,
    TextToColumnsCapability,
    TransposeCapability,
)
from .named_range import NamedRangeCapability

# 透视表向导
from .pivot import PivotWizardCapability

# 范围扩展
from .range import (
    CellLockCapability,
    ClearRangeCapability,
    ColumnRowSizeCapability,
    CopyRangeCapability,
    FindReplaceCapability,
    HyperlinkCapability,
)
from .registry import CapabilityRegistry

# 报表生成
from .report import AdvancedReportCapability, BasicReportCapability, TemplateReportCapability

# 衰分章
from .share import DistributeCapability, GroupByCapability, SubtotalCapability, SummarizeCapability

# 工作表扩展
from .sheet_ext import SheetStyleCapability, SheetVisibilityCapability

# 样式工具
from .style import StyleCapability

# 均输章
from .transport import (
    BatchMergeCapability,
    ExportDataCapability,
    ImportDataCapability,
    JoinTablesCapability,
    MigrateCapability,
)

# 勾股章
from .triangle import (
    AnalyzeCorrelationCapability,
    AnalyzeStatsCapability,
    AnalyzeTrendCapability,
    VisualizeCapability,
)

# 商功章
from .work import (
    AdvancedChartCapability,
    AutoFilterCapability,
    BatchExecuteCapability,
    BatchTransformCapability,
    BatchValidateCapability,
    CommentCapability,
    CreateChartCapability,
    CreatePivotCapability,
    DataViewCapability,
    DeleteChartCapability,
    DiffReportCapability,
    ExportChartAsImageCapability,
    FormatRangeCapability,
    FreezePaneCapability,
    GroupRowsCapability,
    HideShowCapability,
    InsertImageCapability,
    ListChartsCapability,
    PackFileCapability,
    PrintAreaCapability,
    SummaryReportCapability,
    TableCapability,
    UnpackFileCapability,
    UpdateChartCapability,
    ZoomCapability,
)

# Excel 工作流
from .workflow import FormattingWorkflowCapability, SpreadsheetWorkflowCapability

__all__ = [
    "Capability",
    "CapabilitySchema",
    "CapabilityRegistry",
    # 方田章
    "MeasureRangeCapability",
    "MeasureCellsCapability",
    "MeasureStructureCapability",
    "ListSheetsCapability",
    "PeekPreviewCapability",
    "DetectColumnsCapability",
    "SearchContentCapability",
    "GetSummaryCapability",
    "GetSampleDataCapability",
    # 命名范围
    "NamedRangeCapability",
    # 粟米章
    "ConvertFormatCapability",
    "ConvertUnitCapability",
    "ConvertTypeCapability",
    "DataTransformCapability",
    "TransposeCapability",
    "TextToColumnsCapability",
    "FuzzyMatchCapability",
    "AutoTypeInferCapability",
    "StandardizeCapability",
    # 衰分章
    "GroupByCapability",
    "DistributeCapability",
    "SummarizeCapability",
    "SubtotalCapability",
    # 少广章
    "FindDimensionCapability",
    "DeriveCapability",
    "CalculateCapability",
    "SolveEquationCapability",
    "AutoSumCapability",
    # 商功章
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
    # 均输章
    "ImportDataCapability",
    "ExportDataCapability",
    "MigrateCapability",
    "JoinTablesCapability",
    "BatchMergeCapability",
    # 格式转换
    "ExcelToMarkdownCapability",
    "SplitSheetCapability",
    # CSV 处理
    "CSVMergeCapability",
    "CSVVisualizeCapability",
    # 盈不足章
    "ValidateRangeCapability",
    "ValidateTypeCapability",
    "ValidateFormulaCapability",
    "DataValidationCapability",
    "FileValidateCapability",
    "QualityCheckCapability",
    "ExcelLintCapability",
    "FileAnalyzeCapability",
    # 方程章
    "CreateFormulaCapability",
    "DiagnoseFormulaCapability",
    "FormulaRecalcCapability",
    # 勾股章
    "AnalyzeStatsCapability",
    "AnalyzeTrendCapability",
    "AnalyzeCorrelationCapability",
    "VisualizeCapability",
    # 范围扩展
    "ClearRangeCapability",
    "CopyRangeCapability",
    "FindReplaceCapability",
    "HyperlinkCapability",
    "CellLockCapability",
    "ColumnRowSizeCapability",
    # 工作表扩展
    "SheetStyleCapability",
    "SheetVisibilityCapability",
    # 样式工具
    "StyleCapability",
    # 公式生成器
    "FormulaGeneratorCapability",
    # 数据分析模块
    "DataAnalysisCapability",
    "DataCleaningCapability",
    "PivotAnalysisCapability",
    "AdvancedAnalysisCapability",
    # 金融建模
    "VarianceCapability",
    # 报表生成
    "BasicReportCapability",
    "AdvancedReportCapability",
    "TemplateReportCapability",
    # Excel 工作流
    "SpreadsheetWorkflowCapability",
    "FormattingWorkflowCapability",
    # 透视表向导
    "PivotWizardCapability",
]
