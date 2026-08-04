"""MCP Server (FastMCP) - 完整实现"""

from .logging import setup_logging

setup_logging()

from fastmcp import FastMCP

from .core.analysis import (
    AdvancedAnalysisCapability,
    DataAnalysisCapability,
    DataCleaningCapability,
    PivotAnalysisCapability,
)
from .core.balance import (
    DataValidationCapability,
    ExcelLintCapability,
    FileAnalyzeCapability,
    FileValidateCapability,
    ValidateFormulaCapability,
    ValidateRangeCapability,
    ValidateTypeCapability,
    ValidationEngineCapability,
)
from .core.conversion import ExcelToMarkdownCapability, SplitSheetCapability
from .core.csv import CSVMergeCapability, CSVVisualizeCapability
from .core.dimension import (
    AutoSumCapability,
    CalculateCapability,
    DeriveCapability,
    FindDimensionCapability,
    SolveEquationCapability,
)
from .core.equation import (
    CreateFormulaCapability,
    DiagnoseFormulaCapability,
    FormulaRecalcCapability,
)
from .core.field import MeasureCellsCapability, MeasureRangeCapability, MeasureStructureCapability
from .core.finance import VarianceCapability
from .core.formula import FormulaGeneratorCapability
from .core.grain import (
    ConvertFormatCapability,
    ConvertTypeCapability,
    ConvertUnitCapability,
    DataTransformCapability,
    TextToColumnsCapability,
    TransposeCapability,
)
from .core.named_range import NamedRangeCapability
from .core.pivot import PivotWizardCapability
from .core.range import (
    CellLockCapability,
    ClearRangeCapability,
    ColumnRowSizeCapability,
    CopyRangeCapability,
    FindReplaceCapability,
    HyperlinkCapability,
)
from .core.registry import CapabilityRegistry
from .core.report import AdvancedReportCapability, BasicReportCapability, TemplateReportCapability
from .core.share import (
    DistributeCapability,
    GroupByCapability,
    SubtotalCapability,
    SummarizeCapability,
)
from .core.sheet_ext import SheetStyleCapability, SheetVisibilityCapability
from .core.style import StyleCapability
from .core.transport import ExportDataCapability, ImportDataCapability, MigrateCapability
from .core.triangle import (
    AnalyzeCorrelationCapability,
    AnalyzeStatsCapability,
    AnalyzeTrendCapability,
    VisualizeCapability,
)
from .core.work import (
    BatchExecuteCapability,
    BatchTransformCapability,
    BatchValidateCapability,
    CreatePivotCapability,
    DataViewCapability,
    DiffReportCapability,
    SummaryReportCapability,
    TableCapability,
)
from .core.work.advanced_chart import AdvancedChartCapability
from .core.work.advanced_filter import AdvancedFilterCapability
from .core.work.auto_filter import AutoFilterCapability
from .core.work.chart import (
    CreateChartCapability,
    DeleteChartCapability,
    ListChartsCapability,
    UpdateChartCapability,
)
from .core.work.comment import CommentCapability
from .core.work.export_chart import ExportChartAsImageCapability
from .core.work.format import FormatRangeCapability
from .core.work.freeze import FreezePaneCapability
from .core.work.group_rows import GroupRowsCapability
from .core.work.hide_show import HideShowCapability
from .core.work.image import InsertImageCapability
from .core.work.mapping_template import CreateMappingTemplateCapability
from .core.work.pack import PackFileCapability
from .core.work.print_area import PrintAreaCapability
from .core.work.protection import (
    ProtectSheetCapability,
    ProtectWorkbookCapability,
    SetArrayFormulaCapability,
    UnprotectSheetCapability,
)
from .core.work.unpack import UnpackFileCapability
from .core.work.zoom import ZoomCapability
from .core.workflow import FormattingWorkflowCapability, SpreadsheetWorkflowCapability

mcp = FastMCP("abacus")

# 初始化注册表
registry = CapabilityRegistry()


def _get_capability(name: str):
    """获取能力实例，返回 (capability, error_dict)"""
    cap = registry.get(name)
    if cap is None:
        return None, {"error": f"Unknown capability: {name}"}
    return cap, None


# 方田章
registry.register(MeasureRangeCapability())
registry.register(MeasureCellsCapability())
registry.register(MeasureStructureCapability())
registry.register(NamedRangeCapability())

# 粟米章
registry.register(ConvertFormatCapability())
registry.register(ConvertUnitCapability())
registry.register(ConvertTypeCapability())

# 衰分章
registry.register(GroupByCapability())
registry.register(DistributeCapability())
registry.register(SummarizeCapability())

# 少广章
registry.register(FindDimensionCapability())
registry.register(DeriveCapability())
registry.register(CalculateCapability())
registry.register(SolveEquationCapability())
registry.register(AutoSumCapability())

# 商功章
registry.register(BatchExecuteCapability())
registry.register(BatchTransformCapability())
registry.register(BatchValidateCapability())
registry.register(CreatePivotCapability())
registry.register(FormatRangeCapability())
registry.register(CreateChartCapability())
registry.register(UpdateChartCapability())
registry.register(ListChartsCapability())
registry.register(DeleteChartCapability())
registry.register(AdvancedChartCapability())
registry.register(TableCapability())
registry.register(CommentCapability())
registry.register(FreezePaneCapability())
registry.register(AutoFilterCapability())
registry.register(HideShowCapability())
registry.register(AdvancedFilterCapability())
registry.register(CreateMappingTemplateCapability())

# 均输章
registry.register(ImportDataCapability())
registry.register(ExportDataCapability())
registry.register(MigrateCapability())

# 盈不足章
registry.register(ValidateRangeCapability())
registry.register(ValidateTypeCapability())
registry.register(ValidateFormulaCapability())
registry.register(DataValidationCapability())
registry.register(FileValidateCapability())
registry.register(ExcelLintCapability())
registry.register(FileAnalyzeCapability())
registry.register(ValidationEngineCapability())

# 方程章
registry.register(CreateFormulaCapability())
registry.register(DiagnoseFormulaCapability())
registry.register(FormulaRecalcCapability())

# 勾股章
registry.register(AnalyzeStatsCapability())
registry.register(AnalyzeTrendCapability())
registry.register(AnalyzeCorrelationCapability())
registry.register(VisualizeCapability())

# 范围扩展
registry.register(ClearRangeCapability())
registry.register(CopyRangeCapability())
registry.register(FindReplaceCapability())
registry.register(HyperlinkCapability())
registry.register(CellLockCapability())
registry.register(ColumnRowSizeCapability())

# 工作表扩展
registry.register(SheetStyleCapability())
registry.register(SheetVisibilityCapability())

# 审计工具
registry.register(ExcelLintCapability())
registry.register(FileAnalyzeCapability())

# 样式工具
registry.register(StyleCapability())

# 公式生成器
registry.register(FormulaGeneratorCapability())

# 数据分析模块
registry.register(DataAnalysisCapability())
registry.register(DataCleaningCapability())
registry.register(PivotAnalysisCapability())
registry.register(AdvancedAnalysisCapability())

# 数据转换模块
registry.register(DataTransformCapability())

# 金融建模
registry.register(VarianceCapability())

# 报表生成
registry.register(BasicReportCapability())
registry.register(AdvancedReportCapability())
registry.register(TemplateReportCapability())

# CSV 处理
registry.register(CSVMergeCapability())
registry.register(CSVVisualizeCapability())

# 格式转换
registry.register(ExcelToMarkdownCapability())
registry.register(SplitSheetCapability())

# Excel 工作流
registry.register(SpreadsheetWorkflowCapability())
registry.register(FormattingWorkflowCapability())

# 透视表向导
registry.register(PivotWizardCapability())

# 保护工具
registry.register(ProtectWorkbookCapability())
registry.register(ProtectSheetCapability())
registry.register(UnprotectSheetCapability())
registry.register(SetArrayFormulaCapability())

# P1 新增工具
registry.register(InsertImageCapability())
registry.register(GroupRowsCapability())
registry.register(ExportChartAsImageCapability())
registry.register(TransposeCapability())
registry.register(TextToColumnsCapability())
registry.register(SubtotalCapability())

# P3 新增工具
registry.register(PackFileCapability())
registry.register(UnpackFileCapability())
registry.register(PrintAreaCapability())
registry.register(ZoomCapability())

# P2 新增工具 - 智能数据匹配和关联
from .core.balance import QualityCheckCapability
from .core.grain import (
    AutoTypeInferCapability,
    FuzzyMatchCapability,
    StandardizeCapability,
    TransformPipelineCapability,
)
from .core.transport import BatchMergeCapability, JoinTablesCapability

registry.register(FuzzyMatchCapability())
registry.register(AutoTypeInferCapability())
registry.register(StandardizeCapability())
registry.register(TransformPipelineCapability())
registry.register(QualityCheckCapability())
registry.register(JoinTablesCapability())
registry.register(BatchMergeCapability())
registry.register(SummaryReportCapability())
registry.register(DiffReportCapability())
registry.register(DataViewCapability())


# 方田章工具
@mcp.tool()
def measure_range(file: str, sheet: str, range: str) -> dict:
    """[方田章] 读取指定范围数据

    使用场景：
    - 需要读取 Excel 文件中的数据
    - 需要查看特定工作表的特定范围
    - 数据分析、报表生成的第一步

    Args:
        file: Excel 文件路径（必填）
            示例："/path/to/data.xlsx"
        sheet: 工作表名称（必填）
            示例："Sheet1"
        range: 数据范围，使用 A1 表示法（必填）
            示例："A1:D10"

    Returns:
        包含以下字段的字典：
        - range: 实际读取的范围
        - sheet: 工作表名称
        - data: 二维数组数据
        - rows: 行数
        - columns: 列数

    示例：
        measure_range(file="data.xlsx", sheet="Sales", range="A1:C10")
    """
    cap, err = _get_capability("measure_range")
    if err:
        return err
    return cap.execute(None, file=file, sheet=sheet, range=range)


@mcp.tool()
def measure_cells(file: str, sheet: str, range: str) -> dict:
    """[方田章] 读取单元格详细信息（值、公式、样式）

    使用场景：
    - 需要查看单元格的公式内容
    - 需要分析单元格的样式设置
    - 调试 Excel 文件时检查单元格属性

    Args:
        file: Excel 文件路径（必填）
            示例："/path/to/data.xlsx"
        sheet: 工作表名称（必填）
            示例："Sheet1"
        range: 数据范围，使用 A1 表示法（必填）
            示例："A1:D10"

    Returns:
        包含以下字段的字典：
        - cells: 单元格列表，每个包含 value、formula、style 等
        - range: 实际读取的范围

    示例：
        measure_cells(file="data.xlsx", sheet="Sheet1", range="A1:C5")
    """
    cap, err = _get_capability("measure_cells")
    if err:
        return err
    return cap.execute(None, file=file, sheet=sheet, range=range)


@mcp.tool()
def measure_structure(file: str, sheet: str | None = None) -> dict:
    """[方田章] 读取工作表结构（行数、列数、合并单元格等）

    使用场景：
    - 了解 Excel 文件的整体结构
    - 检查工作表的合并单元格情况
    - 确定数据范围和维度

    Args:
        file: Excel 文件路径（必填）
            示例："/path/to/data.xlsx"
        sheet: 工作表名称（可选，默认返回所有工作表信息）
            示例："Sheet1"

    Returns:
        包含以下字段的字典：
        - sheets: 工作表列表
        - dimensions: 数据维度
        - merged_cells: 合并单元格列表

    示例：
        measure_structure(file="data.xlsx", sheet="Sheet1")
    """
    cap, err = _get_capability("measure_structure")
    if err:
        return err
    return cap.execute(None, file=file, sheet=sheet)


@mcp.tool()
def list_sheets(file: str) -> dict:
    """[方田章] 返回 Excel 文件中所有工作表名称列表

    使用场景：
    - 快速获取 Excel 文件包含哪些工作表
    - 获取工作表名称后进行精准读取
    - 批量操作前确认工作表列表

    Args:
        file: Excel 文件路径（必填）
            示例："/path/to/data.xlsx"

    Returns:
        包含以下字段的字典：
        - file: 文件路径
        - sheets: 工作表名称列表
        - count: 工作表数量

    示例：
        list_sheets(file="data.xlsx")

        返回：{"file": "data.xlsx", "sheets": list["Sheet1", "Sales", "Config"], "count": 3}
    """
    cap, err = _get_capability("list_sheets")
    if err:
        return err
    return cap.execute(None, file=file)


@mcp.tool()
def peek_preview(file: str, rows: int = 5, sheet: str | None = None) -> dict:
    """[方田章] 快速预览每个工作表的前几行数据

    使用场景：
    - 快速了解数据长什么样
    - 查看表头和前几行数据
    - 判断数据格式和内容

    Args:
        file: Excel 文件路径（必填）
            示例："/path/to/data.xlsx"
        rows: 预览行数（可选，默认 5）
        sheet: 工作表名称（可选，默认预览所有）

    Returns:
        包含以下字段的字典：
        - preview: 每个工作表的预览数据

    示例：
        peek_preview(file="data.xlsx", rows=3)
    """
    cap, err = _get_capability("peek_preview")
    if err:
        return err
    return cap.execute(None, file=file, rows=rows, sheet=sheet)


@mcp.tool()
def detect_columns(file: str, sheet: str, sample_rows: int = 100) -> dict:
    """[方田章] 检测列名和数据类型

    使用场景：
    - 了解有哪些列可用
    - 检测每列的数据类型
    - 为后续操作选择合适的列

    Args:
        file: Excel 文件路径（必填）
            示例："/path/to/data.xlsx"
        sheet: 工作表名称（必填）
            示例："Sheet1"
        sample_rows: 采样行数（可选，默认 100）

    Returns:
        包含以下字段的字典：
        - columns: 列名列表
        - column_details: 每列的详细信息（类型、采样数等）

    示例：
        detect_columns(file="data.xlsx", sheet="Sheet1")
    """
    cap, err = _get_capability("detect_columns")
    if err:
        return err
    return cap.execute(None, file=file, sheet=sheet, sample_rows=sample_rows)


@mcp.tool()
def search_content(
    file: str, keyword: str, sheet: str | None = None, max_results: int = 50
) -> dict:
    """[方田章] 在 Excel 文件中搜索关键词

    使用场景：
    - 快速定位包含特定内容的单元格
    - 查找特定文本、数字或公式
    - 数据检索和验证

    Args:
        file: Excel 文件路径（必填）
            示例："/path/to/data.xlsx"
        keyword: 搜索关键词（必填）
            示例："销售"
        sheet: 工作表名称（可选，默认搜索所有）
        max_results: 最大结果数（可选，默认 50）

    Returns:
        包含以下字段的字典：
        - results: 搜索结果列表（sheet, cell, value）
        - total_found: 找到的结果数

    示例：
        search_content(file="data.xlsx", keyword="销售")
    """
    cap, err = _get_capability("search_content")
    if err:
        return err
    return cap.execute(None, file=file, keyword=keyword, sheet=sheet, max_results=max_results)


@mcp.tool()
def get_summary(file: str) -> dict:
    """[方田章] 获取 Excel 文件摘要信息

    使用场景：
    - 快速了解文件规模
    - 获取每个工作表的行列数
    - 评估数据量大小

    Args:
        file: Excel 文件路径（必填）
            示例："/path/to/data.xlsx"

    Returns:
        包含以下字段的字典：
        - sheet_count: 工作表数量
        - total_rows: 总行数
        - sheets: 每个工作表的摘要

    示例：
        get_summary(file="data.xlsx")
    """
    cap, err = _get_capability("get_summary")
    if err:
        return err
    return cap.execute(None, file=file)


@mcp.tool()
def get_sample_data(file: str, sheet: str, rows: int = 10) -> dict:
    """[方田章] 获取指定工作表的样本数据

    使用场景：
    - 精准预览某个工作表的数据
    - 获取前 N 行数据用于分析
    - 数据探索和验证

    Args:
        file: Excel 文件路径（必填）
            示例："/path/to/data.xlsx"
        sheet: 工作表名称（必填）
            示例："Sheet1"
        rows: 样本行数（可选，默认 10）

    Returns:
        包含以下字段的字典：
        - columns: 列名列表
        - data: 数据行列表

    示例：
        get_sample_data(file="data.xlsx", sheet="Sheet1", rows=5)
    """
    cap, err = _get_capability("get_sample_data")
    if err:
        return err
    return cap.execute(None, file=file, sheet=sheet, rows=rows)


# 粟米章工具
@mcp.tool()
def convert_format(file: str, sheet: str, range: str, format_type: str) -> dict:
    """[粟米章] 转换数据格式（日期、数字、文本等）

    使用场景：
    - 需要将文本格式的数字转换为数值类型
    - 需要将日期字符串转换为 Excel 日期格式
    - 统一数据格式以便后续计算

    Args:
        file: Excel 文件路径（必填）
            示例："/path/to/data.xlsx"
        sheet: 工作表名称（必填）
            示例："Sheet1"
        range: 数据范围，使用 A1 表示法（必填）
            示例："A1:D10"
        format_type: 目标格式（必填）
            可选值：date、number、text、percentage、currency
            示例："number"

    Returns:
        包含以下字段的字典：
        - success: 是否成功
        - converted_count: 转换单元格数量

    示例：
        convert_format(file="data.xlsx", sheet="Sheet1", range="B2:B100", format_type="number")
    """
    cap, err = _get_capability("convert_format")
    if err:
        return err
    return cap.execute(None, file=file, sheet=sheet, range=range, format_type=format_type)


# 商功章工具
@mcp.tool()
def manage_comment(
    file: str,
    sheet: str,
    action: str,
    cell: str | None = None,
    text: str | None = None,
    author: str = "Abacus",
) -> dict:
    """[商功章] 批注管理（添加、删除、获取批注）

    使用场景：
    - 需要为单元格添加批注
    - 删除不需要的批注
    - 获取批注内容
    - 列出所有批注

    Args:
        file: Excel 文件路径（必填）
            示例："/path/to/data.xlsx"
        sheet: 工作表名称（必填）
            示例："Sheet1"
        action: 操作（必填）
            可选值：add（添加）、delete（删除）、get（获取）、（列出）
            示例："add"
        cell: 单元格位置（add/delete/get 时必填）
            示例："A1"
        text: 批注内容（add 时必填）
            示例："这是一条批注"
        author: 批注作者（可选，默认 "Abacus"）
            示例："张三"

    Returns:
        包含以下字段的字典：
        - action: 执行的操作
        - cell: 单元格位置
        - text: 批注内容（add/get 时）
        - author: 作者（add/get 时）
        - comments: 批注列表（ 时）

    示例：
        manage_comment(file="data.xlsx", sheet="Sheet1", action="add", cell="A1", text="重要数据")
        manage_comment(file="data.xlsx", sheet="Sheet1", action="")
    """
    cap, err = _get_capability("manage_comment")
    if err:
        return err
    return cap.execute(
        None, file=file, sheet=sheet, action=action, cell=cell, text=text, author=author
    )


@mcp.tool()
def freeze_panes(
    file: str,
    sheet: str,
    rows: int | None = None,
    columns: int | None = None,
    cell: str | None = None,
) -> dict:
    """[商功章] 冻结窗格（冻结行、列或行列）

    使用场景：
    - 冻结首行以便滚动时始终可见
    - 冻结首列以便横向滚动时始终可见
    - 同时冻结行和列
    - 解除冻结

    Args:
        file: Excel 文件路径（必填）
            示例："/path/to/data.xlsx"
        sheet: 工作表名称（必填）
            示例："Sheet1"
        rows: 冻结行数（可选）
            示例：1
        columns: 冻结列数（可选）
            示例：1
        cell: 冻结位置（可选，如 B2 表示冻结第一行和第一列）
            示例："B2"

    Returns:
        包含以下字段的字典：
        - action: 执行的操作（freeze/unfreeze）
        - cell: 冻结位置（使用 cell 参数时）
        - rows: 冻结行数（使用 rows 参数时）
        - columns: 冻结列数（使用 columns 参数时）

    示例：
        freeze_panes(file="data.xlsx", sheet="Sheet1", rows=1)
        freeze_panes(file="data.xlsx", sheet="Sheet1", cell="B2")
        freeze_panes(file="data.xlsx", sheet="Sheet1")  # 解除冻结
    """
    cap, err = _get_capability("freeze_panes")
    if err:
        return err
    return cap.execute(None, file=file, sheet=sheet, rows=rows, columns=columns, cell=cell)


@mcp.tool()
def set_auto_filter(
    file: str,
    sheet: str,
    action: str,
    range: str | None = None,
    column: str | None = None,
    criteria: str | None = None,
) -> dict:
    """[商功章] 设置自动筛选（添加、删除、查询筛选）

    使用场景：
    - 为数据区域添加自动筛选下拉箭头
    - 删除已有的自动筛选
    - 查询当前自动筛选状态

    Args:
        file: Excel 文件路径（必填）
            示例："/path/to/data.xlsx"
        sheet: 工作表名称（必填）
            示例："Sheet1"
        action: 操作（必填）
            可选值：set（设置）、remove（删除）、get（查询）
            示例："set"
        range: 筛选范围（set 时必填）
            示例："A1:D100"
        column: 筛选列（可选）
            示例："B"
        criteria: 筛选条件（可选）
            示例："value"

    Returns:
        包含以下字段的字典：
        - action: 执行的操作
        - range: 筛选范围（set/get 时）
        - has_filter: 是否有筛选（get 时）
        - filters: 筛选条件列表（get 时）

    示例：
        set_auto_filter(file="data.xlsx", sheet="Sheet1", action="set", range="A1:D100")
        set_auto_filter(file="data.xlsx", sheet="Sheet1", action="get")
        set_auto_filter(file="data.xlsx", sheet="Sheet1", action="remove")
    """
    cap, err = _get_capability("set_auto_filter")
    if err:
        return err
    return cap.execute(
        None, file=file, sheet=sheet, action=action, range=range, column=column, criteria=criteria
    )


@mcp.tool()
def advanced_filter(
    file: str,
    sheet: str,
    conditions: dict = None,
    return_type: str = "data",
) -> dict:
    """[商功章] 高级筛选（支持复杂条件的数据筛选）

    使用场景：
    - 需要执行复杂的多条件筛选
    - 支持 AND/OR/NOT 逻辑组合
    - 数值范围、文本匹配、日期范围筛选

    Args:
        file: Excel 文件路径（必填）
            示例："/path/to/data.xlsx"
        sheet: 工作表名称（必填）
            示例："Sheet1"
        conditions: 筛选条件（必填），JSON 对象
            示例：{"type": "condition", "field": "Sales", "operator": ">", "value": 1000}
        range: 数据范围（可选，默认整个工作表）
            示例："A1:D100"
        return_type: 返回类型（可选，默认 "data"）
            可选值：data（数据）、rows（行号）
            示例："data"

    Returns:
        包含以下字段的字典：
        - headers: 列名列表
        - rows: 筛选结果数据（return_type 为 data 时）
        - row_numbers: 匹配的行号（return_type 为 rows 时）
        - total_matched: 匹配的总行数

    示例：
        advanced_filter(file="data.xlsx", sheet="Sheet1", conditions={"type": "condition", "field": "Sales", "operator": ">", "value": 1000})
    """
    cap, err = _get_capability("advanced_filter")
    if err:
        return err
    return cap.execute(
        None, file=file, sheet=sheet, range=range, conditions=conditions, return_type=return_type
    )


@mcp.tool()
def create_mapping_template(
    output: str | None = None,
    source_count: int = 4,
    quiet: bool = False,
) -> dict:
    """[商功章] 创建数据映射模板

    使用场景：
    - 创建数据仓库的映射模板
    - 定义目标表与源表的映射关系
    - 标准化数据开发流程

    Args:
        output: 输出文件路径（可选，默认带时间戳）
            示例："/path/to/template.xlsx"
        source_count: 源表数量（可选，默认 4）
            示例：4
        quiet: 静默模式（可选，默认 False）
            示例：False

    Returns:
        包含以下字段的字典：
        - output: 输出文件路径
        - source_count: 源表数量
        - sheets: 创建的工作表列表

    示例：
        create_mapping_template()
        create_mapping_template(output="my_template.xlsx", source_count=3)
    """
    cap, err = _get_capability("create_mapping_template")
    if err:
        return err
    return cap.execute(None, output=output, source_count=source_count, quiet=quiet)


@mcp.tool()
def manage_row_column_visibility(
    file: str,
    sheet: str,
    action: str,
    dimension: str,
    index: int,
) -> dict:
    """[商功章] 管理行列可见性（隐藏/显示行和列）

    使用场景：
    - 隐藏不需要显示的行或列
    - 显示已隐藏的行或列
    - 数据展示时隐藏辅助列

    Args:
        file: Excel 文件路径（必填）
            示例："/path/to/data.xlsx"
        sheet: 工作表名称（必填）
            示例："Sheet1"
        action: 操作（必填）
            可选值：hide（隐藏）、show（显示）
            示例："hide"
        dimension: 维度（必填）
            可选值：row（行）、column（列）
            示例："row"
        index: 行号或列号（必填）
            示例：3

    Returns:
        包含以下字段的字典：
        - action: 执行的操作
        - dimension: 维度
        - index: 行号或列号
        - applied: 是否已应用

    示例：
        manage_row_column_visibility(file="data.xlsx", sheet="Sheet1", action="hide", dimension="row", index=3)
        manage_row_column_visibility(file="data.xlsx", sheet="Sheet1", action="show", dimension="column", index=2)
    """
    cap, err = _get_capability("manage_visibility")
    if err:
        return err
    return cap.execute(
        None, file=file, sheet=sheet, action=action, dimension=dimension, index=index
    )


# 格式化工具
@mcp.tool()
def format_range(
    file: str,
    sheet: str,
    range: str,
    font: dict | None = None,
    fill: dict | None = None,
    border: dict | None = None,
    alignment: dict | None = None,
    number_format: str | None = None,
    conditional: dict | None = None,
) -> dict:
    """[商功章] 格式化单元格（字体、颜色、边框、条件格式等）

    使用场景：
    - 需要美化 Excel 报表
    - 设置单元格的字体、颜色、边框
    - 添加条件格式突出显示数据

    Args:
        file: Excel 文件路径（必填）
            示例："/path/to/data.xlsx"
        sheet: 工作表名称（必填）
            示例："Sheet1"
        range: 数据范围，使用 A1 表示法（必填）
            示例："A1:D10"
        font: 字体设置（可选）
            示例：{"name": "Arial", "size": 12, "bold": true, "color": "000000"}
        fill: 填充设置（可选）
            示例：{"color": "FFFF00", "pattern_type": "solid"}
        border: 边框设置（可选）
            示例：{"style": "thin", "color": "000000"}
        alignment: 对齐设置（可选）
            示例：{"horizontal": "center", "vertical": "center", "wrap_text": true}
        number_format: 数字格式（可选）
            示例："#,##0.00"
        conditional: 条件格式设置（可选）
            示例：{"type": "cell", "operator": "greaterThan", "value": 100}

    Returns:
        包含以下字段的字典：
        - success: 是否成功
        - formatted_range: 已格式化的范围

    示例：
        format_range(file="data.xlsx", sheet="Sheet1", range="A1:D10", font={"bold": true, "size": 14})
    """
    cap, err = _get_capability("format_range")
    if err:
        return err
    return cap.execute(
        None,
        file=file,
        sheet=sheet,
        range=range,
        font=font,
        fill=fill,
        border=border,
        alignment=alignment,
        number_format=number_format,
        conditional=conditional,
    )


@mcp.tool()
def batch_execute(file: str, operations: list) -> dict:
    """[商功章] 批量执行多个操作

    使用场景：
    - 需要一次性执行多个 Excel 操作
    - 减少重复调用，提高效率
    - 复杂的数据处理流程

    Args:
        file: Excel 文件路径（必填）
            示例："/path/to/data.xlsx"
        operations: 操作列表（必填），每个操作包含 type, sheet, range/cell 等
            示例：[
                {"type": "write", "sheet": "Sheet1", "cell": "A1", "value": "Hello"},
                {"type": "format", "sheet": "Sheet1", "range": "A1:D1", "font": {"bold": true}}
            ]

    Returns:
        包含以下字段的字典：
        - success: 是否成功
        - results: 每个操作的结果列表

    示例：
        batch_execute(file="data.xlsx", operations=[{"type": "write", "sheet": "Sheet1", "cell": "A1", "value": "Test"}])
    """
    cap, err = _get_capability("batch_execute")
    if err:
        return err
    return cap.execute(None, file=file, operations=operations)


# 粟米章工具
@mcp.tool()
def convert_unit(file: str, sheet: str, range: str, from_unit: str, to_unit: str) -> dict:
    """[粟米章] 转换单位

    使用场景：
    - 需要将数据从一种单位转换为另一种单位
    - 处理包含单位的数据列
    - 统一数据单位以便分析

    Args:
        file: Excel 文件路径（必填）
            示例："/path/to/data.xlsx"
        sheet: 工作表名称（必填）
            示例："Sheet1"
        range: 数据范围，使用 A1 表示法（必填）
            示例："B2:B100"
        from_unit: 源单位（必填）
            示例："kg"
        to_unit: 目标单位（必填）
            示例："g"

    Returns:
        包含以下字段的字典：
        - success: 是否成功
        - converted_count: 转换单元格数量

    示例：
        convert_unit(file="data.xlsx", sheet="Sheet1", range="B2:B100", from_unit="kg", to_unit="g")
    """
    cap, err = _get_capability("convert_unit")
    if err:
        return err
    return cap.execute(
        None, file=file, sheet=sheet, range=range, from_unit=from_unit, to_unit=to_unit
    )


@mcp.tool()
def convert_type(file: str, sheet: str, range: str, target_type: str) -> dict:
    """[粟米章] 转换数据类型

    使用场景：
    - 需要将文本转换为数值以便计算
    - 需要将数值转换为文本格式
    - 统一数据类型以便后续处理

    Args:
        file: Excel 文件路径（必填）
            示例："/path/to/data.xlsx"
        sheet: 工作表名称（必填）
            示例："Sheet1"
        range: 数据范围，使用 A1 表示法（必填）
            示例："A1:D10"
        target_type: 目标类型（必填）
            可选值：int、float、str、date
            示例："float"

    Returns:
        包含以下字段的字典：
        - success: 是否成功
        - converted_count: 转换单元格数量

    示例：
        convert_type(file="data.xlsx", sheet="Sheet1", range="B2:B100", target_type="float")
    """
    cap, err = _get_capability("convert_type")
    if err:
        return err
    return cap.execute(None, file=file, sheet=sheet, range=range, target_type=target_type)


# 衰分章工具
@mcp.tool()
def group_by(file: str, sheet: str, range: str, group_columns: list[str]) -> dict:
    """[衰分章] 按字段分组

    使用场景：
    - 需要对数据进行分组统计
    - 按类别汇总数据
    - 数据分析中的分组操作

    Args:
        file: Excel 文件路径（必填）
            示例："/path/to/data.xlsx"
        sheet: 工作表名称（必填）
            示例："Sheet1"
        range: 数据范围，使用 A1 表示法（必填）
            示例："A1:D100"
        group_columns: 分组列名列表（必填）
            示例：["Category", "Region"]

    Returns:
        包含以下字段的字典：
        - groups: 分组结果
        - group_count: 分组数量

    示例：
        group_by(file="data.xlsx", sheet="Sheet1", range="A1:D100", group_columns=["Category"])
    """
    cap, err = _get_capability("group_by")
    if err:
        return err
    return cap.execute(None, file=file, sheet=sheet, range=range, group_columns=group_columns)


@mcp.tool()
def distribute(file: str, sheet: str, range: str, total: float, method: str = "equal") -> dict:
    """[衰分章] 按比例分配

    使用场景：
    - 需要将总数按比例分配到各组
    - 预算分配、成本分摊
    - 资源分配场景

    Args:
        file: Excel 文件路径（必填）
            示例："/path/to/data.xlsx"
        sheet: 工作表名称（必填）
            示例："Sheet1"
        range: 数据范围，使用 A1 表示法（必填）
            示例："A1:B10"
        total: 待分配的总数（必填）
            示例：10000
        method: 分配方法（可选，默认 "equal"）
            可选值：equal（等额分配）、proportional（按比例分配）
            示例："proportional"

    Returns:
        包含以下字段的字典：
        - allocations: 分配结果列表
        - total_allocated: 已分配总数

    示例：
        distribute(file="data.xlsx", sheet="Sheet1", range="A1:B10", total=10000, method="proportional")
    """
    cap, err = _get_capability("distribute")
    if err:
        return err
    return cap.execute(None, file=file, sheet=sheet, range=range, total=total, method=method)


@mcp.tool()
def summarize(file: str, sheet: str, range: str, group_by: str, agg_config: dict[str, str]) -> dict:
    """[衰分章] 分组汇总

    使用场景：
    - 需要对数据进行分组汇总
    - 计算各组的总和、平均值、计数等
    - 数据分析中的聚合操作

    Args:
        file: Excel 文件路径（必填）
            示例："/path/to/data.xlsx"
        sheet: 工作表名称（必填）
            示例："Sheet1"
        range: 数据范围，使用 A1 表示法（必填）
            示例："A1:D100"
        group_by: 分组列名（必填）
            示例："Category"
        agg_config: 聚合配置（必填）
            示例：{"Sales": "sum", "Profit": "mean"}

    Returns:
        包含以下字段的字典：
        - summary: 分组汇总结果
        - groups: 分组数量

    示例：
        summarize(file="data.xlsx", sheet="Sheet1", range="A1:D100", group_by="Category", agg_config={"Sales": "sum"})
    """
    cap, err = _get_capability("summarize")
    if err:
        return err
    return cap.execute(
        None, file=file, sheet=sheet, range=range, group_by=group_by, agg_config=agg_config
    )


# 少广章工具
@mcp.tool()
def find_dimension(area: float, shape: str = "rectangle", known_side: float = None) -> dict:
    """[少广章] 已知面积求边长

    使用场景：
    - 已知面积反推边长
    - 几何计算场景
    - 工程设计中的尺寸计算

    Args:
        area: 面积值（必填）
            示例：100
        shape: 形状（可选，默认 "rectangle"）
            可选值：rectangle（长方形）、circle（圆形）
            示例："rectangle"
        known_side: 已知边长（矩形时必填）
            示例：10

    Returns:
        包含以下字段的字典：
        - shape: 形状
        - area: 面积
        - 计算出的尺寸

    示例：
        find_dimension(area=100, shape="rectangle", known_side=10)
    """
    cap, err = _get_capability("find_dimension")
    if err:
        return err
    return cap.execute(None, area=area, shape=shape, known_side=known_side)


@mcp.tool()
def derive(file: str, sheet: str, cell: str, target_value: float, formula: str) -> dict:
    """[少广章] 反向推导

    使用场景：
    - 已知结果反推参数
    - 目标值求解
    - 敏感性分析

    Args:
        file: Excel 文件路径（必填）
            示例："/path/to/data.xlsx"
        sheet: 工作表名称（必填）
            示例："Sheet1"
        cell: 单元格位置（必填）
            示例："A1"
        target_value: 目标值（必填）
            示例：1000
        formula: 公式（必填）
            示例："A2*B2"

    Returns:
        包含以下字段的字典：
        - result: 推导结果
        - iterations: 迭代次数

    示例：
        derive(file="data.xlsx", sheet="Sheet1", cell="A1", target_value=1000, formula="A2*B2")
    """
    cap, err = _get_capability("derive")
    if err:
        return err
    return cap.execute(
        None, file=file, sheet=sheet, cell=cell, target_value=target_value, formula=formula
    )


# 商功章工具
@mcp.tool()
def batch_transform(file: str, operations: list) -> dict:
    """[商功章] 批量转换

    使用场景：
    - 需要一次性执行多个转换操作
    - 批量数据类型转换
    - 批量格式转换

    Args:
        file: Excel 文件路径（必填）
            示例："/path/to/data.xlsx"
        operations: 操作列表（必填），每个操作包含 type、range、target 等
            示例：[
                {"type": "convert_type", "range": "A1:A10", "target": "float"},
                {"type": "convert_format", "range": "B1:B10", "target": "number"}
            ]

    Returns:
        包含以下字段的字典：
        - success: 是否成功
        - results: 每个操作的结果列表

    示例：
        batch_transform(file="data.xlsx", operations=[{"type": "convert_type", "range": "A1:A10", "target": "float"}])
    """
    cap, err = _get_capability("batch_transform")
    if err:
        return err
    return cap.execute(None, file=file, operations=operations)


@mcp.tool()
def batch_validate(file: str, operations: list) -> dict:
    """[商功章] 批量验证

    使用场景：
    - 需要一次性执行多个验证操作
    - 批量数据验证
    - 数据质量检查

    Args:
        file: Excel 文件路径（必填）
            示例："/path/to/data.xlsx"
        operations: 操作列表（必填），每个操作包含 type、range、expected 等
            示例：[
                {"type": "validate_range", "range": "A1:A10", "min": 0, "max": 100},
                {"type": "validate_type", "range": "B1:B10", "expected": "float"}
            ]

    Returns:
        包含以下字段的字典：
        - success: 是否成功
        - results: 每个操作的结果列表
        - errors: 验证错误列表

    示例：
        batch_validate(file="data.xlsx", operations=[{"type": "validate_range", "range": "A1:A10", "min": 0, "max": 100}])
    """
    cap, err = _get_capability("batch_validate")
    if err:
        return err
    return cap.execute(None, file=file, operations=operations)


@mcp.tool()
def create_pivot(
    file: str,
    sheet: str,
    range: str,
    row_fields: list[str],
    value_field: str,
    agg_function: str = "sum",
    output_sheet: str | None = None,
) -> dict:
    """[商功章] 创建数据透视表

    使用场景：
    - 需要对数据进行多维度分析
    - 创建汇总报表
    - 数据分组统计

    Args:
        file: Excel 文件路径（必填）
            示例："/path/to/data.xlsx"
        sheet: 源数据工作表（必填）
            示例："Sheet1"
        range: 源数据范围，使用 A1 表示法（必填）
            示例："A1:D100"
        row_fields: 行字段列表（必填）
            示例：["Category", "Region"]
        value_field: 值字段（必填）
            示例："Sales"
        agg_function: 聚合函数（可选，默认 "sum"）
            可选值：sum、avg、count、min、max
            示例："sum"
        output_sheet: 输出工作表名称（可选）
            示例："PivotResult"

    Returns:
        包含以下字段的字典：
        - success: 是否成功
        - output_sheet: 输出工作表名称
        - rows: 结果行数

    示例：
        create_pivot(file="data.xlsx", sheet="Sheet1", range="A1:D100", row_fields=["Category"], value_field="Sales")
    """
    cap, err = _get_capability("create_pivot")
    if err:
        return err
    return cap.execute(
        None,
        file=file,
        sheet=sheet,
        range=range,
        row_fields=row_fields,
        value_field=value_field,
        agg_function=agg_function,
        output_sheet=output_sheet,
    )


# 均输章工具
@mcp.tool()
def import_data(file: str, source: str, source_type: str = "csv", sheet: str = "Sheet1") -> dict:
    """[均输章] 导入数据

    使用场景：
    - 需要将 CSV 文件导入 Excel
    - 需要将 JSON 数据导入 Excel
    - 数据迁移和整合

    Args:
        file: Excel 文件路径（必填）
            示例："/path/to/output.xlsx"
        source: 源文件路径（必填）
            示例："/path/to/data.csv"
        source_type: 源文件类型（可选，默认 "csv"）
            可选值：csv、json
            示例："csv"
        sheet: 目标工作表（可选，默认 "Sheet1"）
            示例："ImportedData"

    Returns:
        包含以下字段的字典：
        - success: 是否成功
        - rows_imported: 导入行数
        - columns_imported: 导入列数

    示例：
        import_data(file="output.xlsx", source="data.csv", source_type="csv", sheet="Sales")
    """
    cap, err = _get_capability("import_data")
    if err:
        return err
    return cap.execute(None, file=file, source=source, source_type=source_type, sheet=sheet)


@mcp.tool()
def export_data(file: str, sheet: str, range: str, output: str, format: str = "csv") -> dict:
    """[均输章] 导出数据

    使用场景：
    - 需要将 Excel 数据导出为 CSV
    - 需要将 Excel 数据导出为 JSON
    - 数据共享和分发

    Args:
        file: Excel 文件路径（必填）
            示例："/path/to/data.xlsx"
        sheet: 工作表名称（必填）
            示例："Sheet1"
        range: 数据范围，使用 A1 表示法（必填）
            示例："A1:D100"
        output: 输出文件路径（必填）
            示例："/path/to/output.csv"
        format: 输出格式（可选，默认 "csv"）
            可选值：csv、json
            示例："csv"

    Returns:
        包含以下字段的字典：
        - success: 是否成功
        - rows_exported: 导出行数
        - columns_exported: 导出列数

    示例：
        export_data(file="data.xlsx", sheet="Sheet1", range="A1:D100", output="output.csv")
    """
    cap, err = _get_capability("export_data")
    if err:
        return err
    return cap.execute(None, file=file, sheet=sheet, range=range, output=output, format=format)


@mcp.tool()
def migrate(source: str, target: str, sheets: list[list[str]] = None) -> dict:
    """[均输章] 数据迁移

    使用场景：
    - 需要将数据从一个 Excel 文件迁移到另一个
    - 需要复制特定工作表
    - 数据备份和归档

    Args:
        source: 源文件路径（必填）
            示例："/path/to/source.xlsx"
        target: 目标文件路径（必填）
            示例："/path/to/target.xlsx"
        sheets: 工作表列表（可选，默认迁移所有工作表）
            示例：["Sheet1", "Sheet2"]

    Returns:
        包含以下字段的字典：
        - success: 是否成功
        - migrated_sheets: 已迁移的工作表列表

    示例：
        migrate(source="source.xlsx", target="target.xlsx", sheets=["Sheet1", "Sheet2"])
    """
    cap, err = _get_capability("migrate")
    if err:
        return err
    return cap.execute(None, source=source, target=target, sheets=sheets)


# 盈不足章工具
@mcp.tool()
def set_data_validation(
    file: str,
    sheet: str,
    range: str,
    validation_type: str,
    operator: str | None = None,
    formula1: str | None = None,
    formula2: str | None = None,
    error_message: str | None = None,
) -> dict:
    """[盈不足章] 设置单元格数据验证规则

    使用场景：
    - 需要创建下拉列表
    - 限制数值输入范围
    - 限制日期输入范围
    - 限制文本长度

    Args:
        file: Excel 文件路径（必填）
            示例："/path/to/data.xlsx"
        sheet: 工作表名称（必填）
            示例："Sheet1"
        range: 数据范围（必填）
            示例："A1:A10"
        validation_type: 验证类型（必填）
            可选值：（下拉列表）、number（数值范围）、date（日期范围）、text_length（文本长度）
            示例：""
        operator: 运算符（可选）
            可选值：between、notBetween、equal、notEqual 等
            示例："between"
        formula1: 验证公式1（可选）
            示例："选项1,选项2,选项3"（ 类型时为逗号分隔的选项）
        formula2: 验证公式2（可选，between 时需要）
            示例："100"
        error_message: 错误提示消息（可选）
            示例："请输入有效值"

    Returns:
        包含以下字段的字典：
        - success: 是否成功
        - sheet: 工作表名称
        - range: 数据范围
        - validation_type: 验证类型
        - applied: 是否已应用

    示例：
        set_data_validation(file="data.xlsx", sheet="Sheet1", range="A1:A10", validation_type="", formula1="是,否")
        set_data_validation(file="data.xlsx", sheet="Sheet1", range="B1:B10", validation_type="number", operator="between", formula1="0", formula2="100")
    """
    cap, err = _get_capability("set_data_validation")
    if err:
        return err
    return cap.execute(
        None,
        file=file,
        sheet=sheet,
        range=range,
        validation_type=validation_type,
        operator=operator,
        formula1=formula1,
        formula2=formula2,
        error_message=error_message,
    )


@mcp.tool()
def validate_range(
    file: str, sheet: str, range: str, min_value: float | None = None, max_value: list[float] = None
) -> dict:
    """[盈不足章] 验证数据范围

    使用场景：
    - 需要检查数据是否在指定范围内
    - 数据质量检查
    - 异常值检测

    Args:
        file: Excel 文件路径（必填）
            示例："/path/to/data.xlsx"
        sheet: 工作表名称（必填）
            示例："Sheet1"
        range: 数据范围，使用 A1 表示法（必填）
            示例："A1:A100"
        min_value: 最小值（可选）
            示例：0
        max_value: 最大值（可选）
            示例：100

    Returns:
        包含以下字段的字典：
        - valid: 是否全部通过验证
        - invalid_count: 不符合条件的单元格数量
        - invalid_cells: 不符合条件的单元格列表

    示例：
        validate_range(file="data.xlsx", sheet="Sheet1", range="A1:A100", min_value=0, max_value=100)
    """
    cap, err = _get_capability("validate_range")
    if err:
        return err
    return cap.execute(
        None, file=file, sheet=sheet, range=range, min_value=min_value, max_value=max_value
    )


@mcp.tool()
def validate_type(file: str, sheet: str, range: str, expected_type: str) -> dict:
    """[盈不足章] 验证数据类型

    使用场景：
    - 需要检查数据类型是否符合预期
    - 数据质量检查
    - 数据清洗前的类型验证

    Args:
        file: Excel 文件路径（必填）
            示例："/path/to/data.xlsx"
        sheet: 工作表名称（必填）
            示例："Sheet1"
        range: 数据范围，使用 A1 表示法（必填）
            示例："A1:A100"
        expected_type: 期望类型（必填）
            可选值：int、float、str、date
            示例："float"

    Returns:
        包含以下字段的字典：
        - valid: 是否全部通过验证
        - invalid_count: 不符合条件的单元格数量
        - invalid_cells: 不符合条件的单元格列表

    示例：
        validate_type(file="data.xlsx", sheet="Sheet1", range="A1:A100", expected_type="float")
    """
    cap, err = _get_capability("validate_type")
    if err:
        return err
    return cap.execute(None, file=file, sheet=sheet, range=range, expected_type=expected_type)


@mcp.tool()
def validate_formula(file: str, sheet: str | None = None, cell: str | None = None) -> dict:
    """[盈不足章] 验证公式正确性

    使用场景：
    - 需要检查公式是否正确
    - 调试 Excel 公式
    - 公式错误检测

    Args:
        file: Excel 文件路径（必填）
            示例："/path/to/data.xlsx"
        sheet: 工作表名称（必填）
            示例："Sheet1"
        cell: 单元格位置（必填）
            示例："E1"

    Returns:
        包含以下字段的字典：
        - valid: 公式是否正确
        - formula: 公式内容
        - error: 错误信息（如有）

    示例：
        validate_formula(file="data.xlsx", sheet="Sheet1", cell="E1")
    """
    cap, err = _get_capability("validate_formula")
    if err:
        return err
    return cap.execute(None, file=file, sheet=sheet, cell=cell)


@mcp.tool()
def validate_file(file: str) -> dict:
    """[盈不足章] 验证 Excel 文件结构（ZIP 格式、XML 结构、公式错误）

    使用场景：
    - 需要检查 Excel 文件质量
    - 发现文件中的潜在问题
    - 文件健康检查

    Args:
        file: Excel 文件路径（必填）
            示例："/path/to/data.xlsx"

    Returns:
        包含以下字段的字典：
        - file: 文件路径
        - valid: 是否有效
        - checks: 通过的检查列表
        - errors: 错误列表
        - warnings: 警告列表

    示例：
        validate_file(file="data.xlsx")
    """
    cap, err = _get_capability("validate_file")
    if err:
        return err
    return cap.execute(None, file=file)


@mcp.tool()
def validation_engine(
    file: str,
    sheet: str,
    range: str,
    rules: list,
) -> dict:
    """[盈不足章] 数据验证规则引擎（自定义规则、AND/OR 组合、规则链）

    使用场景：
    - 需要执行复杂的多规则验证
    - 支持 AND/OR 逻辑组合
    - 按顺序执行规则链
    - 自定义 Python 表达式验证

    Args:
        file: Excel 文件路径（必填）
            示例："/path/to/data.xlsx"
        sheet: 工作表名称（必填）
            示例："Sheet1"
        range: 数据范围，使用 A1 表示法（必填）
            示例："A1:D100"
        rules: 验证规则列表（必填），每个规则包含 type 和可选 params
            规则类型：
            - type: 类型验证 {"type": "type", "params": {"expected_type": "int"}}
            - range: 范围验证 {"type": "range", "params": {"min_val": 0, "max_val": 100}}
            - not_empty: 非空验证 {"type": "not_empty"}
            - regex: 正则验证 {"type": "regex", "params": {"pattern": "^\\d+$"}}
            - custom: 自定义表达式 {"type": "custom", "params": {"expression": "value > 0"}}
            - and: AND 组合 {"type": "and", "rules": [...]}
            - or: OR 组合 {"type": "or", "rules": [...]}

    Returns:
        包含以下字段的字典：
        - valid: 是否全部通过验证
        - total_cells: 总单元格数
        - passed_cells: 通过的单元格数
        - failed_cells: 失败的单元格数
        - results: 每个单元格的验证结果

    示例：
        validation_engine(file="data.xlsx", sheet="Sheet1", range="A1:A10",
                         rules=[{"type": "type", "params": {"expected_type": "int"}},
                                {"type": "range", "params": {"min_val": 0, "max_val": 100}}])
    """
    cap, err = _get_capability("validation_engine")
    if err:
        return err
    return cap.execute(None, file=file, sheet=sheet, range=range, rules=rules)


# 方程章工具
@mcp.tool()
def diagnose_formula(file: str, sheet: str | None = None, cell: str | None = None) -> dict:
    """[方程章] 诊断公式错误（分析 #REF!, #N/A, #VALUE!, #NAME?, #DIV/0! 等错误）

    使用场景：
    - 需要检查 Excel 文件中的公式错误
    - 诊断公式语法问题
    - 查找包含错误的单元格
    - 公式调试和修复

    Args:
        file: Excel 文件路径（必填）
            示例："/path/to/data.xlsx"
        sheet: 工作表名称（可选，默认检查所有工作表）
            示例："Sheet1"
        cell: 单元格位置（可选，默认检查所有公式单元格）
            示例："A1"

    Returns:
        包含以下字段的字典：
        - file: 文件路径
        - formulas_checked: 检查的公式数量
        - errors_found: 发现的错误数量
        - errors: 错误列表，每个包含 sheet、cell、formula、error、description 等

    示例：
        diagnose_formula(file="data.xlsx")
        diagnose_formula(file="data.xlsx", sheet="Sheet1")
        diagnose_formula(file="data.xlsx", sheet="Sheet1", cell="A1")
    """
    cap, err = _get_capability("diagnose_formula")
    if err:
        return err
    return cap.execute(None, file=file, sheet=sheet, cell=cell)


@mcp.tool()
def create_formula(file: str, sheet: str, cell: str, formula: str) -> dict:
    """[方程章] 在指定单元格创建公式

    使用场景：
    - 需要在单元格中创建公式
    - 自动化公式设置
    - 批量公式创建

    Args:
        file: Excel 文件路径（必填）
            示例："/path/to/data.xlsx"
        sheet: 工作表名称（必填）
            示例："Sheet1"
        cell: 单元格位置（必填）
            示例："E1"
        formula: 公式内容（必填）
            示例："SUM(A1:D1)"

    Returns:
        包含以下字段的字典：
        - success: 是否成功
        - cell: 单元格位置
        - formula: 公式内容

    示例：
        create_formula(file="data.xlsx", sheet="Sheet1", cell="E1", formula="SUM(A1:D1)")
    """
    cap, err = _get_capability("create_formula")
    if err:
        return err
    return cap.execute(None, file=file, sheet=sheet, cell=cell, formula=formula)


@mcp.tool()
def recalc_formulas(file: str, output: str | None = None) -> dict:
    """[方程章] 公式重算：使用 LibreOffice 重算 Excel 公式（扫描所有错误）

    使用场景：
    - 需要重算 Excel 文件中的公式
    - 检查公式错误
    - 批量公式验证

    Args:
        file: Excel 文件路径（必填）
            示例："/path/to/data.xlsx"
        output: 输出文件路径（可选）
            示例："/path/to/recalc.xlsx"

    Returns:
        包含以下字段的字典：
        - file: 文件路径
        - output: 输出路径
        - errors_found: 错误数量
        - errors: 错误列表
        - recalculated: 是否重算成功

    示例：
        recalc_formulas(file="data.xlsx")
        recalc_formulas(file="data.xlsx", output="recalc.xlsx")
    """
    cap, err = _get_capability("recalc_formulas")
    if err:
        return err
    return cap.execute(None, file=file, output=output)


@mcp.tool()
def solve_equation(equation: str) -> dict:
    """[少广章] 解方程

    使用场景：
    - 需要解数学方程
    - 求解未知数
    - 数学计算场景

    Args:
        equation: 方程表达式（必填）
            示例："2x + 3 = 7"

    Returns:
        包含以下字段的字典：
        - solution: 解
        - equation: 方程
        - type: 方程类型

    示例：
        solve_equation(equation="2x + 3 = 7")
    """
    cap, err = _get_capability("solve_equation")
    if err:
        return err
    return cap.execute(None, equation=equation)


@mcp.tool()
def calculate(expression: str, variables: dict[str, float] | None = None) -> dict:
    """[少广章] 执行计算

    使用场景：
    - 需要执行数学计算
    - 计算表达式求值
    - 数学运算场景

    Args:
        expression: 计算表达式（必填）
            示例："2 + 3 * 4"
        variables: 变量值（可选）
            示例：{"x": 10, "y": 20}

    Returns:
        包含以下字段的字典：
        - result: 计算结果
        - expression: 表达式

    示例：
        calculate(expression="2 + 3 * 4")
        calculate(expression="x + y", variables={"x": 10, "y": 20})
    """
    cap, err = _get_capability("calculate")
    if err:
        return err
    return cap.execute(None, expression=expression, variables=variables or {})


# 勾股章工具
@mcp.tool()
def analyze_stats(file: str, sheet: str, range: str) -> dict:
    """[勾股章] 统计分析

    使用场景：
    - 需要对数据进行统计分析
    - 计算均值、中位数、标准差等
    - 数据分布分析

    Args:
        file: Excel 文件路径（必填）
            示例："/path/to/data.xlsx"
        sheet: 工作表名称（必填）
            示例："Sheet1"
        range: 数据范围，使用 A1 表示法（必填）
            示例："A1:A100"

    Returns:
        包含以下字段的字典：
        - mean: 均值
        - median: 中位数
        - std: 标准差
        - min: 最小值
        - max: 最大值
        - count: 计数

    示例：
        analyze_stats(file="data.xlsx", sheet="Sheet1", range="A1:A100")
    """
    cap, err = _get_capability("analyze_stats")
    if err:
        return err
    return cap.execute(None, file=file, sheet=sheet, range=range)


@mcp.tool()
def analyze_trend(
    file: str, sheet: str, range: str, value_column: str, time_column: str = None
) -> dict:
    """[勾股章] 趋势分析

    使用场景：
    - 需要分析数据随时间的变化趋势
    - 识别增长/下降趋势
    - 时间序列分析

    Args:
        file: Excel 文件路径（必填）
            示例："/path/to/data.xlsx"
        sheet: 工作表名称（必填）
            示例："Sheet1"
        range: 数据范围，使用 A1 表示法（必填）
            示例："A1:C100"
        value_column: 值列名（必填）
            示例："Sales"
        time_column: 时间列名（可选）
            示例："Date"

    Returns:
        包含以下字段的字典：
        - trend: 趋势方向（up/down/stable）
        - slope: 斜率
        - r_squared: R² 值
        - forecast: 预测值

    示例：
        analyze_trend(file="data.xlsx", sheet="Sheet1", range="A1:C100", value_column="Sales")
    """
    cap, err = _get_capability("analyze_trend")
    if err:
        return err
    return cap.execute(
        None,
        file=file,
        sheet=sheet,
        range=range,
        value_column=value_column,
        time_column=time_column,
    )


@mcp.tool()
def analyze_correlation(file: str, sheet: str, range: str, column1: str, column2: str) -> dict:
    """[勾股章] 相关性分析

    使用场景：
    - 需要分析两个变量之间的相关性
    - 识别变量间的关联关系
    - 数据探索和特征选择

    Args:
        file: Excel 文件路径（必填）
            示例："/path/to/data.xlsx"
        sheet: 工作表名称（必填）
            示例："Sheet1"
        range: 数据范围，使用 A1 表示法（必填）
            示例："A1:C100"
        column1: 第一列名（必填）
            示例："Sales"
        column2: 第二列名（必填）
            示例："Profit"

    Returns:
        包含以下字段的字典：
        - correlation: 相关系数
        - strength: 相关强度（strong/moderate/weak）
        - direction: 相关方向（positive/negative）

    示例：
        analyze_correlation(file="data.xlsx", sheet="Sheet1", range="A1:C100", column1="Sales", column2="Profit")
    """
    cap, err = _get_capability("analyze_correlation")
    if err:
        return err
    return cap.execute(None, file=file, sheet=sheet, range=range, column1=column1, column2=column2)


@mcp.tool()
def create_chart(
    file: str,
    sheet: str,
    range: str,
    chart_type: str,
    title: str | None = None,
    x_axis: str | None = None,
    y_axis: str | None = None,
    output_sheet: str | None = None,
    position: str = "A1",
    width: float = 15,
    height: float = 10,
) -> dict:
    """[商功章] 创建图表（柱形图、折线图、饼图、面积图、散点图）

    使用场景：
    - 需要将数据可视化
    - 创建各种类型的图表
    - 报表中添加图表

    Args:
        file: Excel 文件路径（必填）
            示例："/path/to/data.xlsx"
        sheet: 工作表名称（必填）
            示例："Sheet1"
        range: 数据范围，使用 A1 表示法（必填）
            示例："A1:C10"
        chart_type: 图表类型（必填）
            可选值：bar（柱形图）、line（折线图）、pie（饼图）、area（面积图）、scatter（散点图）
            示例："bar"
        title: 图表标题（可选）
            示例："销售趋势"
        x_axis: X 轴标题（可选）
            示例："月份"
        y_axis: Y 轴标题（可选）
            示例："销售额"
        output_sheet: 输出工作表名称（可选）
            示例："Charts"
        position: 图表位置（可选，默认 "A1"）
            示例："A1"
        width: 图表宽度（可选，默认 15）
            示例：15
        height: 图表高度（可选，默认 10）
            示例：10

    Returns:
        包含以下字段的字典：
        - success: 是否成功
        - chart_index: 图表索引
        - chart_type: 图表类型

    示例：
        create_chart(file="data.xlsx", sheet="Sheet1", range="A1:C10", chart_type="bar", title="销售趋势")
    """
    cap, err = _get_capability("create_chart")
    if err:
        return err
    return cap.execute(
        None,
        file=file,
        sheet=sheet,
        range=range,
        chart_type=chart_type,
        title=title,
        x_axis=x_axis,
        y_axis=y_axis,
        output_sheet=output_sheet,
        position=position,
        width=width,
        height=height,
    )


@mcp.tool()
def update_chart(file: str, sheet: str, chart_index: int, title: str | None = None) -> dict:
    """[商功章] 更新图表

    使用场景：
    - 需要修改图表标题
    - 更新图表属性
    - 图表维护

    Args:
        file: Excel 文件路径（必填）
            示例："/path/to/data.xlsx"
        sheet: 工作表名称（必填）
            示例："Sheet1"
        chart_index: 图表索引（必填）
            示例：0
        title: 新图表标题（可选）
            示例："更新后的标题"

    Returns:
        包含以下字段的字典：
        - success: 是否成功
        - chart_index: 图表索引

    示例：
        update_chart(file="data.xlsx", sheet="Sheet1", chart_index=0, title="新标题")
    """
    cap, err = _get_capability("update_chart")
    if err:
        return err
    return cap.execute(None, file=file, sheet=sheet, chart_index=chart_index, title=title)


@mcp.tool()
def list_charts(file: str, sheet: str | None = None) -> dict:
    """[商功章] 列出所有图表

    使用场景：
    - 需要查看文件中有哪些图表
    - 图表管理和维护
    - 图表索引查询

    Args:
        file: Excel 文件路径（必填）
            示例："/path/to/data.xlsx"
        sheet: 工作表名称（可选，默认返回所有工作表的图表）
            示例："Sheet1"

    Returns:
        包含以下字段的字典：
        - charts: 图表列表，每个包含 index、type、title 等
        - count: 图表数量

    示例：
        list_charts(file="data.xlsx", sheet="Sheet1")
    """
    cap, err = _get_capability("list_charts")
    if err:
        return err
    return cap.execute(None, file=file, sheet=sheet)


@mcp.tool()
def delete_chart(file: str, sheet: str, chart_index: int) -> dict:
    """[商功章] 删除图表

    使用场景：
    - 需要删除不需要的图表
    - 清理文件中的图表
    - 图表管理

    Args:
        file: Excel 文件路径（必填）
            示例："/path/to/data.xlsx"
        sheet: 工作表名称（必填）
            示例："Sheet1"
        chart_index: 图表索引（必填）
            示例：0

    Returns:
        包含以下字段的字典：
        - success: 是否成功
        - deleted_chart: 已删除的图表信息

    示例：
        delete_chart(file="data.xlsx", sheet="Sheet1", chart_index=0)
    """
    cap, err = _get_capability("delete_chart")
    if err:
        return err
    return cap.execute(None, file=file, sheet=sheet, chart_index=chart_index)


@mcp.tool()
def create_advanced_chart(
    file: str,
    data: dict,
    chart_type: str,
    title: str | None = None,
    x_axis: str | None = None,
    y_axis: str | None = None,
) -> dict:
    """[商功章] 创建高级图表（组合图、双轴图、瀑布图、甘特图）

    使用场景：
    - 创建组合图（柱形图+折线图）
    - 创建双轴图（左右Y轴）
    - 创建瀑布图（展示增减变化）
    - 创建甘特图（项目进度）

    Args:
        file: 输出文件路径（必填）
            示例："/path/to/chart.xlsx"
        data: 图表数据（必填），包含 headers 和 rows
            示例：{"headers": list["月份", "销售额", "利润"], "rows": list[["1月", 100, 20], ["2月", 120, 25]]}
        chart_type: 图表类型（必填）
            可选值：combo（组合图）、dual_axis（双轴图）、waterfall（瀑布图）、gantt（甘特图）
            示例："combo"
        title: 图表标题（可选）
            示例："销售趋势"
        x_axis: X轴标题（可选）
            示例："月份"
        y_axis: Y轴标题（可选）
            示例："金额"

    Returns:
        包含以下字段的字典：
        - file: 输出文件路径
        - chart_type: 图表类型
        - title: 图表标题
        - created: 是否创建成功

    示例：
        create_advanced_chart(file="chart.xlsx", data={"headers": list["月份", "销售额"], "rows": [["1月", 100]]}, chart_type="combo")
    """
    cap, err = _get_capability("create_advanced_chart")
    if err:
        return err
    return cap.execute(
        None, file=file, data=data, chart_type=chart_type, title=title, x_axis=x_axis, y_axis=y_axis
    )


@mcp.tool()
def visualize(
    file: str,
    output: str,
    chart_type: str,
    x_column: str | None = None,
    y_column: str | None = None,
    sheet: str | None = None,
    title: str | None = None,
    width: float = 10,
    height: float = 6,
) -> dict:
    """[勾股章] 数据可视化（生成 PNG/SVG/PDF 图表）

    使用场景：
    - 将 Excel/CSV 数据生成图片格式的图表
    - 创建柱状图、折线图、饼图、散点图、热力图
    - 数据报告和演示

    Args:
        file: 数据文件路径（必填）
            示例："/path/to/data.xlsx"
        output: 输出图片路径（必填）
            示例："/path/to/chart.png"
        chart_type: 图表类型（必填）
            可选值：bar（柱状图）、line（折线图）、pie（饼图）、scatter（散点图）、heatmap（热力图）
            示例："bar"
        x_column: X轴列名（可选）
            示例："月份"
        y_column: Y轴列名（可选）
            示例："销售额"
        sheet: 工作表名称（可选）
            示例："Sheet1"
        title: 图表标题（可选）
            示例："销售趋势"
        width: 图片宽度，单位英寸（可选，默认 10）
            示例：10
        height: 图片高度，单位英寸（可选，默认 6）
            示例：6

    Returns:
        包含以下字段的字典：
        - file: 数据文件路径
        - output: 输出图片路径
        - chart_type: 图表类型
        - created: 是否创建成功

    示例：
        visualize(file="data.xlsx", output="chart.png", chart_type="bar", title="销售趋势")
    """
    cap, err = _get_capability("visualize")
    if err:
        return err
    return cap.execute(
        None,
        file=file,
        output=output,
        chart_type=chart_type,
        x_column=x_column,
        y_column=y_column,
        sheet=sheet,
        title=title,
        width=width,
        height=height,
    )


# 范围扩展工具
@mcp.tool()
def clear_range(file: str, sheet: str, range: str, clear_type: str = "all") -> dict:
    """[范围扩展] 清除范围内容（值/公式/格式/全部）

    使用场景：
    - 需要清除单元格内容
    - 重置数据区域
    - 清理格式

    Args:
        file: Excel 文件路径（必填）
            示例："/path/to/data.xlsx"
        sheet: 工作表名称（必填）
            示例："Sheet1"
        range: 数据范围，使用 A1 表示法（必填）
            示例："A1:D10"
        clear_type: 清除类型（可选，默认 "all"）
            可选值：all（全部）、contents（内容）、formats（格式）
            示例："contents"

    Returns:
        包含以下字段的字典：
        - success: 是否成功
        - cleared_range: 已清除的范围

    示例：
        clear_range(file="data.xlsx", sheet="Sheet1", range="A1:D10", clear_type="contents")
    """
    cap, err = _get_capability("clear_range")
    if err:
        return err
    return cap.execute(None, file=file, sheet=sheet, range=range, clear_type=clear_type)


@mcp.tool()
def copy_range(file: str, sheet: str, source: str, target: str, copy_type: str = "all") -> dict:
    """[范围扩展] 复制范围（值/公式/格式）

    使用场景：
    - 需要复制数据到其他位置
    - 复制公式或格式
    - 数据备份

    Args:
        file: Excel 文件路径（必填）
            示例："/path/to/data.xlsx"
        sheet: 工作表名称（必填）
            示例："Sheet1"
        source: 源范围，使用 A1 表示法（必填）
            示例："A1:D10"
        target: 目标位置（必填）
            示例："F1"
        copy_type: 复制类型（可选，默认 "all"）
            可选值：all（全部）、values（值）、formulas（公式）、formats（格式）
            示例："values"

    Returns:
        包含以下字段的字典：
        - success: 是否成功
        - source: 源范围
        - target: 目标位置

    示例：
        copy_range(file="data.xlsx", sheet="Sheet1", source="A1:D10", target="F1", copy_type="values")
    """
    cap, err = _get_capability("copy_range")
    if err:
        return err
    return cap.execute(
        None, file=file, sheet=sheet, source=source, target=target, copy_type=copy_type
    )


@mcp.tool()
def find_replace(
    file: str, sheet: str, find: str, replace: str | None = None, range: str | None = None
) -> dict:
    """[范围扩展] 查找替换文本

    使用场景：
    - 需要批量替换文本内容
    - 数据清洗和修正
    - 文本标准化

    Args:
        file: Excel 文件路径（必填）
            示例："/path/to/data.xlsx"
        sheet: 工作表名称（必填）
            示例："Sheet1"
        find: 查找内容（必填）
            示例："旧文本"
        replace: 替换内容（可选）
            示例："新文本"
        range: 查找范围（可选，默认整个工作表）
            示例："A1:D100"

    Returns:
        包含以下字段的字典：
        - success: 是否成功
        - replacements: 替换次数

    示例：
        find_replace(file="data.xlsx", sheet="Sheet1", find="旧文本", replace="新文本")
    """
    cap, err = _get_capability("find_replace")
    if err:
        return err
    return cap.execute(None, file=file, sheet=sheet, find=find, replace=replace, range=range)


@mcp.tool()
def manage_hyperlink(
    file: str, sheet: str, action: str, cell: str | None = None, url: str | None = None
) -> dict:
    """[范围扩展] 管理超链接（添加/删除/列出）

    使用场景：
    - 需要在单元格中添加超链接
    - 管理文件中的超链接
    - 创建导航链接

    Args:
        file: Excel 文件路径（必填）
            示例："/path/to/data.xlsx"
        sheet: 工作表名称（必填）
            示例："Sheet1"
        action: 操作（必填）
            可选值：add（添加）、remove（删除）、（列出）
            示例："add"
        cell: 单元格位置（action 为 add/remove 时必填）
            示例："A1"
        url: 链接地址（action 为 add 时必填）
            示例："https://example.com"

    Returns:
        包含以下字段的字典：
        - success: 是否成功
        - hyperlinks: 超链接列表（action 为  时）

    示例：
        manage_hyperlink(file="data.xlsx", sheet="Sheet1", action="add", cell="A1", url="https://example.com")
    """
    cap, err = _get_capability("manage_hyperlink")
    if err:
        return err
    return cap.execute(None, file=file, sheet=sheet, action=action, cell=cell, url=url)


@mcp.tool()
def manage_cell_lock(file: str, sheet: str, range: str, locked: bool = True) -> dict:
    """[范围扩展] 管理单元格锁定（锁定/解锁）

    使用场景：
    - 需要保护特定单元格不被修改
    - 解锁单元格以便编辑
    - 工作表保护设置

    Args:
        file: Excel 文件路径（必填）
            示例："/path/to/data.xlsx"
        sheet: 工作表名称（必填）
            示例："Sheet1"
        range: 数据范围，使用 A1 表示法（必填）
            示例："A1:D10"
        locked: 是否锁定（可选，默认 True）
            示例：True

    Returns:
        包含以下字段的字典：
        - success: 是否成功
        - locked: 锁定状态

    示例：
        manage_cell_lock(file="data.xlsx", sheet="Sheet1", range="A1:D10", locked=True)
    """
    cap, err = _get_capability("manage_cell_lock")
    if err:
        return err
    return cap.execute(None, file=file, sheet=sheet, range=range, locked=locked)


@mcp.tool()
def manage_size(
    file: str, sheet: str, action: str, dimension: str, index: int, size: list[float] = None
) -> dict:
    """[范围扩展] 管理行列大小（列宽/行高）

    使用场景：
    - 需要调整列宽或行高
    - 自动调整行列大小
    - 获取当前行列大小

    Args:
        file: Excel 文件路径（必填）
            示例："/path/to/data.xlsx"
        sheet: 工作表名称（必填）
            示例："Sheet1"
        action: 操作（必填）
            可选值：set（设置）、get（获取）、auto（自动调整）
            示例："set"
        dimension: 维度（必填）
            可选值：row（行）、column（列）
            示例："column"
        index: 行号或列号（必填）
            示例：1
        size: 大小（action 为 set 时必填）
            示例：20

    Returns:
        包含以下字段的字典：
        - success: 是否成功
        - size: 当前大小（action 为 get 时）

    示例：
        manage_size(file="data.xlsx", sheet="Sheet1", action="set", dimension="column", index=1, size=20)
    """
    cap, err = _get_capability("manage_size")
    if err:
        return err
    return cap.execute(
        None, file=file, sheet=sheet, action=action, dimension=dimension, index=index, size=size
    )


# 表格管理工具
@mcp.tool()
def manage_table(
    file: str,
    sheet: str,
    action: str,
    table_name: str | None = None,
    range: str | None = None,
    style: str | None = None,
    data: list | None = None,
) -> dict:
    """[商功章] 管理 Excel 表格（创建/列出/删除/追加）

    使用场景：
    - 需要创建结构化表格
    - 管理文件中的表格
    - 向表格追加数据

    Args:
        file: Excel 文件路径（必填）
            示例："/path/to/data.xlsx"
        sheet: 工作表名称（必填）
            示例："Sheet1"
        action: 操作（必填）
            可选值：create（创建）、（列出）、delete（删除）、append（追加）
            示例："create"
        table_name: 表格名称（action 为 create/delete/append 时必填）
            示例："SalesTable"
        range: 数据范围（action 为 create 时必填）
            示例："A1:D100"
        style: 表格样式（可选）
            示例："TableStyleMedium9"
        data: 追加数据（action 为 append 时必填）
            示例：[{"Name": "Product A", "Sales": 100}]

    Returns:
        包含以下字段的字典：
        - success: 是否成功
        - tables: 表格列表（action 为  时）

    示例：
        manage_table(file="data.xlsx", sheet="Sheet1", action="create", table_name="SalesTable", range="A1:D100")
    """
    cap, err = _get_capability("manage_table")
    if err:
        return err
    return cap.execute(
        None,
        file=file,
        sheet=sheet,
        action=action,
        table_name=table_name,
        range=range,
        style=style,
        data=data,
    )


# 命名范围管理工具
@mcp.tool()
def manage_named_range(
    file: str,
    action: str,
    name: str | None = None,
    refers_to: str | None = None,
) -> dict:
    """[方田章] 管理命名范围（创建/列出/读取/删除）

    使用场景：
    - 需要创建命名范围简化引用
    - 管理文件中的命名范围
    - 使用命名范围提高公式可读性

    Args:
        file: Excel 文件路径（必填）
            示例："/path/to/data.xlsx"
        action: 操作（必填）
            可选值：create（创建）、（列出）、read（读取）、delete（删除）
            示例："create"
        name: 命名范围名称（action 为 create/read/delete 时必填）
            示例："SalesData"
        refers_to: 引用位置（action 为 create 时必填）
            示例："Sheet1!$A$1:$D$10"

    Returns:
        包含以下字段的字典：
        - success: 是否成功
        - named_ranges: 命名范围列表（action 为  时）

    示例：
        manage_named_range(file="data.xlsx", action="create", name="SalesData", refers_to="Sheet1!$A$1:$D$10")
    """
    cap, err = _get_capability("manage_named_range")
    if err:
        return err
    return cap.execute(None, file=file, action=action, name=name, refers_to=refers_to)


# 工作表扩展工具
@mcp.tool()
def manage_sheet_style(file: str, sheet: str, action: str, color: str | None = None) -> dict:
    """[方田章] 管理工作表样式（标签颜色）

    使用场景：
    - 需要设置工作表标签颜色
    - 获取当前标签颜色
    - 清除标签颜色

    Args:
        file: Excel 文件路径（必填）
            示例："/path/to/data.xlsx"
        sheet: 工作表名称（必填）
            示例："Sheet1"
        action: 操作（必填）
            可选值：set（设置）、get（获取）、clear（清除）
            示例："set"
        color: 颜色（action 为 set 时必填，使用十六进制颜色代码）
            示例："FF0000"

    Returns:
        包含以下字段的字典：
        - success: 是否成功
        - color: 当前颜色（action 为 get 时）

    示例：
        manage_sheet_style(file="data.xlsx", sheet="Sheet1", action="set", color="FF0000")
    """
    cap, err = _get_capability("manage_sheet_style")
    if err:
        return err
    return cap.execute(None, file=file, sheet=sheet, action=action, color=color)


@mcp.tool()
def manage_sheet_visibility(file: str, sheet: str, action: str) -> dict:
    """[方田章] 管理工作表可见性（显示/隐藏/非常隐藏）

    使用场景：
    - 需要隐藏工作表
    - 显示隐藏的工作表
    - 获取工作表可见性状态

    Args:
        file: Excel 文件路径（必填）
            示例："/path/to/data.xlsx"
        sheet: 工作表名称（必填）
            示例："Sheet1"
        action: 操作（必填）
            可选值：show（显示）、hide（隐藏）、very-hide（非常隐藏）、get（获取状态）
            示例："hide"

    Returns:
        包含以下字段的字典：
        - success: 是否成功
        - visibility: 可见性状态（action 为 get 时）

    示例：
        manage_sheet_visibility(file="data.xlsx", sheet="Sheet1", action="hide")
    """
    cap, err = _get_capability("manage_sheet_visibility")
    if err:
        return err
    return cap.execute(None, file=file, sheet=sheet, action=action)


# 审计工具
@mcp.tool()
def excel_lint(code: str | None = None, file: str | None = None) -> dict:
    """[审计] 检查 openpyxl 代码的 10 类常见问题

    使用场景：
    - 检查 openpyxl 代码质量
    - 发现常见编程错误
    - 代码审查

    Args:
        code: Python 代码内容（与 file 二选一）
            示例："import openpyxl\nwb = openpyxl.load_workbook('test.xlsx')"
        file: Python 文件路径（与 code 二选一）
            示例："/path/to/script.py"

    Returns:
        包含以下字段的字典：
        - issues: 问题列表
        - issue_count: 问题数量
        - severity: 严重程度

    示例：
        excel_lint(code="import openpyxl\nwb = openpyxl.load_workbook('test.xlsx')")
    """
    cap, err = _get_capability("excel_lint")
    if err:
        return err
    return cap.execute(None, code=code, file=file)


@mcp.tool()
def file_analyze(file: str) -> dict:
    """[审计] 检查 Excel 文件的 10 类常见问题

    使用场景：
    - 检查 Excel 文件质量
    - 发现文件中的潜在问题
    - 文件健康检查

    Args:
        file: Excel 文件路径（必填）
            示例："/path/to/data.xlsx"

    Returns:
        包含以下字段的字典：
        - issues: 问题列表
        - issue_count: 问题数量
        - severity: 严重程度

    示例：
        file_analyze(file="data.xlsx")
    """
    cap, err = _get_capability("file_analyze")
    if err:
        return err
    return cap.execute(None, file=file)


# 样式工具
@mcp.tool()
def manage_style(
    file: str, sheet: str, action: str, range: str | None = None, industry: str = "finance"
) -> dict:
    """[样式] 管理样式（行业品牌色、表头、KPI 格式）

    使用场景：
    - 应用行业品牌色
    - 设置表头样式
    - 应用 KPI 格式
    - 自动调整列宽

    Args:
        file: Excel 文件路径（必填）
            示例："/path/to/data.xlsx"
        sheet: 工作表名称（必填）
            示例："Sheet1"
        action: 操作（必填）
            可选值：apply_header（应用表头）、apply_kpi（应用 KPI 格式）、auto_width（自动列宽）
            示例："apply_header"
        range: 数据范围（可选）
            示例："A1:D10"
        industry: 行业（可选，默认 "finance"）
            可选值：finance、ecommerce、saas、internet
            示例："finance"

    Returns:
        包含以下字段的字典：
        - success: 是否成功
        - applied_style: 应用的样式

    示例：
        manage_style(file="data.xlsx", sheet="Sheet1", action="apply_header", range="A1:D10", industry="finance")
    """
    cap, err = _get_capability("manage_style")
    if err:
        return err
    return cap.execute(None, file=file, sheet=sheet, action=action, range=range, industry=industry)


# 公式生成器工具
@mcp.tool()
def generate_formula(
    formula_type: str,
    params: dict,
    file: str | None = None,
    sheet: str | None = None,
    cell: str | None = None,
) -> dict:
    """[方程章] 生成常用 Excel 公式（VLOOKUP、SUMIFS、IF 等）

    使用场景：
    - 需要生成复杂公式
    - 不熟悉 Excel 公式语法
    - 快速创建常用公式

    Args:
        formula_type: 公式类型（必填）
            可选值：vlookup、sumifs、if、today、npv、pmt、irr 等
            示例："vlookup"
        params: 公式参数（必填）
            示例：{"lookup_value": "A1", "table_array": "B:C", "col_index": 2}
        file: Excel 文件路径（可选，写入公式时使用）
            示例："/path/to/data.xlsx"
        sheet: 工作表名称（可选）
            示例："Sheet1"
        cell: 单元格位置（可选）
            示例："E1"

    Returns:
        包含以下字段的字典：
        - formula: 生成的公式
        - description: 公式说明

    示例：
        generate_formula(formula_type="vlookup", params={"lookup_value": "A1", "table_array": "B:C", "col_index": 2})
    """
    cap, err = _get_capability("generate_formula")
    if err:
        return err
    return cap.execute(
        None, formula_type=formula_type, params=params, file=file, sheet=sheet, cell=cell
    )


# 数据分析工具
@mcp.tool()
def analyze_data(file: str, sheet: str | None = None, analysis_type: str = "auto") -> dict:
    """[勾股章] 智能数据分析（自动检测数据类型、统计摘要、相关性分析）

    使用场景：
    - 快速了解数据概况
    - 自动检测数据类型
    - 生成统计摘要
    - 发现数据间的相关性

    Args:
        file: 文件路径（必填，支持 Excel 和 CSV）
            示例："/path/to/data.xlsx"
        sheet: 工作表名称（可选，Excel 文件时使用）
            示例："Sheet1"
        analysis_type: 分析类型（可选，默认 "auto"）
            可选值：auto（自动）、summary（摘要）、correlation（相关性）
            示例："auto"

    Returns:
        包含以下字段的字典：
        - data_types: 数据类型检测结果
        - statistics: 统计摘要
        - correlations: 相关性分析结果

    示例：
        analyze_data(file="data.xlsx", sheet="Sheet1", analysis_type="auto")
    """
    cap, err = _get_capability("analyze_data")
    if err:
        return err
    return cap.execute(None, file=file, sheet=sheet, analysis_type=analysis_type)


@mcp.tool()
def clean_data(
    file: str,
    sheet: str | None = None,
    output: str | None = None,
    operations: list[str] | None = None,
) -> dict:
    """[粟米章] 数据清洗（去重、缺失值处理、格式化）

    使用场景：
    - 清洗脏数据
    - 处理缺失值
    - 去除重复数据
    - 格式化文本

    Args:
        file: 文件路径（必填）
            示例："/path/to/data.xlsx"
        sheet: 工作表名称（可选）
            示例："Sheet1"
        output: 输出文件路径（可选，默认覆盖原文件）
            示例："/path/to/cleaned_data.xlsx"
        operations: 清洗操作列表（可选，默认执行所有操作）
            可选值：remove_duplicates（去重）、handle_missing（处理缺失值）、strip_whitespace（去除空白）、convert_types（类型转换）
            示例：["remove_duplicates", "handle_missing"]

    Returns:
        包含以下字段的字典：
        - success: 是否成功
        - rows_affected: 受影响的行数
        - operations_performed: 执行的操作列表

    示例：
        clean_data(file="data.xlsx", operations=["remove_duplicates", "handle_missing"])
    """
    cap, err = _get_capability("clean_data")
    if err:
        return err
    return cap.execute(None, file=file, sheet=sheet, output=output, operations=operations)


@mcp.tool()
def pivot_analysis(
    file: str,
    group_by: str,
    value_field: str,
    sheet: str | None = None,
    agg_function: str = "sum",
    output: str | None = None,
) -> dict:
    """[衰分章] 数据透视分析（分组汇总、交叉分析）

    使用场景：
    - 需要对数据进行分组汇总
    - 交叉分析
    - 数据聚合统计

    Args:
        file: 文件路径（必填）
            示例："/path/to/data.xlsx"
        group_by: 分组字段（必填）
            示例："Category"
        value_field: 值字段（必填）
            示例："Sales"
        sheet: 工作表名称（可选）
            示例："Sheet1"
        agg_function: 聚合函数（可选，默认 "sum"）
            可选值：sum、mean、count、min、max
            示例："sum"
        output: 输出文件路径（可选）
            示例："/path/to/pivot_result.xlsx"

    Returns:
        包含以下字段的字典：
        - success: 是否成功
        - pivot_table: 透视表数据

    示例：
        pivot_analysis(file="data.xlsx", group_by="Category", value_field="Sales", agg_function="sum")
    """
    cap, err = _get_capability("pivot_analysis")
    if err:
        return err
    return cap.execute(
        None,
        file=file,
        sheet=sheet,
        group_by=group_by,
        value_field=value_field,
        agg_function=agg_function,
        output=output,
    )


# 金融建模工具
@mcp.tool()
def variance_analysis(
    file: str,
    budget_sheet: str,
    actual_sheet: str,
    output: str | None = None,
    threshold: float = 0.1,
) -> dict:
    """[勾股章] 预算与实际差异分析

    使用场景：
    - 预算执行情况分析
    - 实际与预算对比
    - 差异原因分析

    Args:
        file: Excel 文件路径（必填）
            示例："/path/to/data.xlsx"
        budget_sheet: 预算数据工作表（必填）
            示例："Budget"
        actual_sheet: 实际数据工作表（必填）
            示例："Actual"
        output: 输出文件路径（可选）
            示例："/path/to/variance_report.xlsx"
        threshold: 重要性阈值（可选，默认 0.1，表示 10%）
            示例：0.1

    Returns:
        包含以下字段的字典：
        - success: 是否成功
        - variances: 差异列表
        - significant_variances: 重大差异列表

    示例：
        variance_analysis(file="data.xlsx", budget_sheet="Budget", actual_sheet="Actual", threshold=0.1)
    """
    cap, err = _get_capability("variance_analysis")
    if err:
        return err
    return cap.execute(
        None,
        file=file,
        budget_sheet=budget_sheet,
        actual_sheet=actual_sheet,
        output=output,
        threshold=threshold,
    )


# 报表生成工具
@mcp.tool()
def create_basic_report(
    data_source: str,
    output: str,
    sheet_name: str = "Data",
    title: str | None = None,
) -> dict:
    """[商功章] 从数据生成基础 Excel 报表（自动格式化、列宽调整、冻结首行）

    使用场景：
    - 快速生成报表
    - 从 CSV 数据创建 Excel 报表
    - 自动格式化数据

    Args:
        data_source: 数据源（必填，CSV 文件路径）
            示例："/path/to/data.csv"
        output: 输出文件路径（必填）
            示例："/path/to/report.xlsx"
        sheet_name: 工作表名称（可选，默认 "Data"）
            示例："Sales"
        title: 报表标题（可选）
            示例："月度销售报表"

    Returns:
        包含以下字段的字典：
        - success: 是否成功
        - output: 输出文件路径
        - rows: 数据行数

    示例：
        create_basic_report(data_source="data.csv", output="report.xlsx", title="月度销售报表")
    """
    cap, err = _get_capability("create_basic_report")
    if err:
        return err
    return cap.execute(
        None, data_source=data_source, output=output, sheet_name=sheet_name, title=title
    )


@mcp.tool()
def create_advanced_report(
    data_source: str,
    output: str,
    chart_type: str = "bar",
    include_dashboard: bool = True,
) -> dict:
    """[商功章] 生成高级 Excel 报表（图表、条件格式、透视汇总、仪表板）

    使用场景：
    - 生成专业级报表
    - 包含图表和仪表板
    - 高级数据可视化

    Args:
        data_source: 数据源（必填，CSV 文件路径）
            示例："/path/to/data.csv"
        output: 输出文件路径（必填）
            示例："/path/to/advanced_report.xlsx"
        chart_type: 图表类型（可选，默认 "bar"）
            可选值：bar（柱形图）、line（折线图）、pie（饼图）
            示例："bar"
        include_dashboard: 是否包含仪表板（可选，默认 True）
            示例：True

    Returns:
        包含以下字段的字典：
        - success: 是否成功
        - output: 输出文件路径
        - charts_created: 创建的图表数量

    示例：
        create_advanced_report(data_source="data.csv", output="report.xlsx", chart_type="bar", include_dashboard=True)
    """
    cap, err = _get_capability("create_advanced_report")
    if err:
        return err
    return cap.execute(
        None,
        data_source=data_source,
        output=output,
        chart_type=chart_type,
        include_dashboard=include_dashboard,
    )


@mcp.tool()
def fill_template(
    template: str,
    output: str,
    data_source: str | None = None,
    data: list | None = None,
    sheet_name: str | None = None,
    start_cell: str = "A1",
) -> dict:
    """[商功章] 基于模板填充数据生成报表（支持命名单元格、批量填充）

    使用场景：
    - 基于模板生成报表
    - 批量填充数据
    - 自动化报表生成

    Args:
        template: 模板文件路径（必填）
            示例："/path/to/template.xlsx"
        output: 输出文件路径（必填）
            示例："/path/to/output.xlsx"
        data_source: 数据源（可选，CSV 文件路径）
            示例："/path/to/data.csv"
        data: 填充数据（可选，字典格式）
            示例：{"Name": "John", "Sales": 1000}
        sheet_name: 工作表名称（可选）
            示例："Sheet1"
        start_cell: 起始单元格（可选，默认 "A1"）
            示例："A1"

    Returns:
        包含以下字段的字典：
        - success: 是否成功
        - output: 输出文件路径
        - cells_filled: 填充的单元格数量

    示例：
        fill_template(template="template.xlsx", output="output.xlsx", data={"Name": "John", "Sales": 1000})
    """
    cap, err = _get_capability("fill_template")
    if err:
        return err
    return cap.execute(
        None,
        template=template,
        output=output,
        data_source=data_source,
        data=data,
        sheet_name=sheet_name,
        start_cell=start_cell,
    )


# CSV 处理工具
@mcp.tool()
def merge_files(
    files: list[str],
    output: str,
    merge_type: str = "concat",
    on: str | None = None,
    dedup: bool = False,
    dedup_columns: list[list[str]] = None,
) -> dict:
    """[均输章] 合并多个 CSV/Excel 文件

    使用场景：
    - 合并多个数据文件
    - 数据整合
    - 批量数据处理

    Args:
        files: 文件路径列表（必填）
            示例：["/path/to/file1.csv", "/path/to/file2.csv"]
        output: 输出文件路径（必填）
            示例："/path/to/merged.csv"
        merge_type: 合并类型（可选，默认 "concat"）
            可选值：concat（纵向合并）、merge（横向合并）、join（连接）
            示例："concat"
        on: 合并键（merge/join 时必填）
            示例："ID"
        dedup: 是否去重（可选，默认 False）
            示例：True
        dedup_columns: 去重列（可选）
            示例：["ID", "Name"]

    Returns:
        包含以下字段的字典：
        - success: 是否成功
        - rows_merged: 合并行数
        - output: 输出文件路径

    示例：
        merge_files(files=["file1.csv", "file2.csv"], output="merged.csv", merge_type="concat")
    """
    cap, err = _get_capability("merge_files")
    if err:
        return err
    return cap.execute(
        None,
        files=files,
        output=output,
        merge_type=merge_type,
        on=on,
        dedup=dedup,
        dedup_columns=dedup_columns,
    )


@mcp.tool()
def visualize_data(
    file: str,
    output: str,
    chart_type: str = "auto",
    include_dashboard: bool = True,
    include_stats: bool = True,
) -> dict:
    """[勾股章] CSV 数据可视化（自动生成图表、仪表板、统计摘要）

    使用场景：
    - 快速数据可视化
    - 自动生成图表
    - 数据探索分析

    Args:
        file: 文件路径（必填，支持 CSV 和 Excel）
            示例："/path/to/data.csv"
        output: 输出文件路径（必填）
            示例："/path/to/visualization.xlsx"
        chart_type: 图表类型（可选，默认 "auto"）
            可选值：bar（柱形图）、line（折线图）、pie（饼图）、auto（自动选择）
            示例："auto"
        include_dashboard: 是否包含仪表板（可选，默认 True）
            示例：True
        include_stats: 是否包含统计摘要（可选，默认 True）
            示例：True

    Returns:
        包含以下字段的字典：
        - success: 是否成功
        - output: 输出文件路径
        - charts_created: 创建的图表数量

    示例：
        visualize_data(file="data.csv", output="visualization.xlsx", chart_type="auto")
    """
    cap, err = _get_capability("visualize_data")
    if err:
        return err
    return cap.execute(
        None,
        file=file,
        output=output,
        chart_type=chart_type,
        include_dashboard=include_dashboard,
        include_stats=include_stats,
    )


# 格式转换工具
@mcp.tool()
def excel_to_markdown(
    file: str,
    sheet: str | None = None,
    output: str | None = None,
    merge_mode: str = "tl",
    include_styles: bool = True,
) -> dict:
    """[均输章] 将 Excel 表格转换为 Markdown 格式

    使用场景：
    - 将 Excel 表格转换为 Markdown
    - 在文档中嵌入表格
    - 生成技术文档

    Args:
        file: Excel 文件路径（必填）
            示例："/path/to/data.xlsx"
        sheet: 工作表名称（可选，默认全部）
            示例："Sheet1"
        output: 输出文件路径（可选）
            示例："/path/to/output.md"
        merge_mode: 合并单元格处理模式（可选，默认 "tl"）
            可选值：tl（使用左上角值）、fill（填充）
            示例："tl"
        include_styles: 是否包含样式（可选，默认 True）
            示例：True

    Returns:
        包含以下字段的字典：
        - success: 是否成功
        - markdown: Markdown 内容
        - output: 输出文件路径

    示例：
        excel_to_markdown(file="data.xlsx", sheet="Sheet1", include_styles=True)
    """
    cap, err = _get_capability("excel_to_markdown")
    if err:
        return err
    return cap.execute(
        None,
        file=file,
        sheet=sheet,
        output=output,
        merge_mode=merge_mode,
        include_styles=include_styles,
    )


@mcp.tool()
def split_sheet(
    file: str,
    sheet: str,
    output_dir: str,
    split_by: str,
    split_column: str | None = None,
    row_count: int | None = None,
    prefix: str = "split",
) -> dict:
    """[商功章] 将 Excel 工作表按条件拆分为多个文件

    使用场景：
    - 将大文件拆分为小文件
    - 按类别拆分数据
    - 数据分发

    Args:
        file: Excel 文件路径（必填）
            示例："/path/to/data.xlsx"
        sheet: 工作表名称（必填）
            示例："Sheet1"
        output_dir: 输出目录（必填）
            示例："/path/to/output"
        split_by: 拆分方式（必填）
            可选值：column（按列拆分）、row_count（按行数拆分）、range（按范围拆分）
            示例："column"
        split_column: 拆分列名（split_by=column 时必填）
            示例："Category"
        row_count: 每文件行数（split_by=row_count 时必填）
            示例：1000
        prefix: 输出文件前缀（可选，默认 "split"）
            示例："category"

    Returns:
        包含以下字段的字典：
        - success: 是否成功
        - files_created: 创建的文件数量
        - output_dir: 输出目录

    示例：
        split_sheet(file="data.xlsx", sheet="Sheet1", output_dir="output", split_by="column", split_column="Category")
    """
    cap, err = _get_capability("split_sheet")
    if err:
        return err
    return cap.execute(
        None,
        file=file,
        sheet=sheet,
        output_dir=output_dir,
        split_by=split_by,
        split_column=split_column,
        row_count=row_count,
        prefix=prefix,
    )


@mcp.tool()
def advanced_analysis(
    file: str,
    analysis_type: str,
    sheet: str | None = None,
    x_column: str | None = None,
    y_column: str | None = None,
    periods: int = 10,
) -> dict:
    """[勾股章] 高级数据分析（回归分析、时间序列、预测）

    使用场景：
    - 线性回归分析，获取斜率、截距和 R² 值
    - 时间序列分析，识别趋势和季节性
    - 线性外推预测

    Args:
        file: 文件路径（必填）
            示例："/path/to/data.xlsx"
        analysis_type: 分析类型（必填）
            可选值：regression（回归分析）、timeseries（时间序列）、forecast（预测）
            示例："regression"
        sheet: 工作表名称（可选）
            示例："Sheet1"
        x_column: 自变量列名（regression 时必填）
            示例："X"
        y_column: 因变量列名（regression/timeseries/forecast 时必填）
            示例："Y"
        periods: 预测期数（forecast 时可选，默认 10）
            示例：10

    Returns:
        包含以下字段的字典：
        - type: 分析类型
        - regression: 斜率、截距、R²（regression 时）
        - timeseries: 统计信息、趋势（timeseries 时）
        - forecast: 历史数据和预测值（forecast 时）

    示例：
        advanced_analysis(file="data.xlsx", analysis_type="regression", x_column="X", y_column="Y")
        advanced_analysis(file="data.xlsx", analysis_type="forecast", y_column="Sales", periods=12)
    """
    cap, err = _get_capability("advanced_analysis")
    if err:
        return err
    return cap.execute(
        None,
        file=file,
        analysis_type=analysis_type,
        sheet=sheet,
        x_column=x_column,
        y_column=y_column,
        periods=periods,
    )


@mcp.tool()
def transform_data(
    file: str,
    transform_type: str,
    sheet: str | None = None,
    params: list | None = None,
    output: str | None = None,
) -> dict:
    """[粟米章] 高级数据转换（透视、转置、合并、重塑）

    使用场景：
    - 透视表转换（宽表聚合）
    - 逆透视（宽表转长表）
    - 合并多个数据源
    - 数据重塑

    Args:
        file: 文件路径（必填）
            示例："/path/to/data.xlsx"
        transform_type: 转换类型（必填）
            可选值：pivot（透视）、melt（逆透视）、merge（合并）、reshape（重塑）
            示例："pivot"
        sheet: 工作表名称（可选）
            示例："Sheet1"
        params: 转换参数（可选）
            pivot 示例：{"index": "Category", "values": "Sales", "aggfunc": "sum"}
            melt 示例：{"id_vars": list["ID"], "value_vars": ["A", "B"]}
            merge 示例：{"other_file": "other.xlsx", "on": "ID"}
            reshape 示例：{"pivot_column": "Type", "value_column": "Amount"}
        output: 输出文件路径（可选）
            示例："/path/to/output.xlsx"

    Returns:
        包含以下字段的字典：
        - transform_type: 转换类型
        - input_rows: 输入行数
        - output_rows: 输出行数
        - output_columns: 输出列数

    示例：
        transform_data(file="data.xlsx", transform_type="pivot", params={"index": "Category", "values": "Sales"})
        transform_data(file="data.xlsx", transform_type="melt", params={"id_vars": ["ID"]})
    """
    cap, err = _get_capability("transform_data")
    if err:
        return err
    return cap.execute(
        None,
        file=file,
        transform_type=transform_type,
        sheet=sheet,
        params=params or {},
        output=output,
    )


# 保护工具
@mcp.tool()
def protect_workbook(file: str, password: str | None = None) -> dict:
    """[商功章] 保护工作簿

    使用场景：
    - 需要保护工作簿结构不被修改
    - 防止添加/删除/重命名工作表
    - 保护工作簿配置

    Args:
        file: Excel 文件路径（必填）
            示例："/path/to/data.xlsx"
        password: 保护密码（可选）
            示例："mypassword"

    Returns:
        包含以下字段的字典：
        - success: 是否成功
        - action: 执行的操作
        - file: 文件路径

    示例：
        protect_workbook(file="data.xlsx")
        protect_workbook(file="data.xlsx", password="mypassword")
    """
    cap, err = _get_capability("protect_workbook")
    if err:
        return err
    return cap.execute(None, file=file, password=password)


@mcp.tool()
def protect_sheet(file: str, sheet: str, password: str | None = None) -> dict:
    """[商功章] 保护工作表

    使用场景：
    - 需要保护工作表不被修改
    - 防止意外编辑
    - 保护公式和数据

    Args:
        file: Excel 文件路径（必填）
            示例："/path/to/data.xlsx"
        sheet: 工作表名称（必填）
            示例："Sheet1"
        password: 保护密码（可选）
            示例："mypassword"

    Returns:
        包含以下字段的字典：
        - success: 是否成功
        - action: 执行的操作
        - sheet: 工作表名称

    示例：
        protect_sheet(file="data.xlsx", sheet="Sheet1")
        protect_sheet(file="data.xlsx", sheet="Sheet1", password="mypassword")
    """
    cap, err = _get_capability("protect_sheet")
    if err:
        return err
    return cap.execute(None, file=file, sheet=sheet, password=password)


@mcp.tool()
def unprotect_sheet(file: str, sheet: str, password: str | None = None) -> dict:
    """[商功章] 解除工作表保护

    使用场景：
    - 需要编辑受保护的工作表
    - 解除保护以进行修改
    - 更改受保护的内容

    Args:
        file: Excel 文件路径（必填）
            示例："/path/to/data.xlsx"
        sheet: 工作表名称（必填）
            示例："Sheet1"
        password: 保护密码（可选）
            示例："mypassword"

    Returns:
        包含以下字段的字典：
        - success: 是否成功
        - action: 执行的操作
        - sheet: 工作表名称

    示例：
        unprotect_sheet(file="data.xlsx", sheet="Sheet1")
        unprotect_sheet(file="data.xlsx", sheet="Sheet1", password="mypassword")
    """
    cap, err = _get_capability("unprotect_sheet")
    if err:
        return err
    return cap.execute(None, file=file, sheet=sheet, password=password)


@mcp.tool()
def set_array_formula(file: str, sheet: str, range: str, formula: str) -> dict:
    """[方程章] 设置数组公式

    使用场景：
    - 需要创建数组公式
    - 执行多单元格计算
    - 复杂的数据分析

    Args:
        file: Excel 文件路径（必填）
            示例："/path/to/data.xlsx"
        sheet: 工作表名称（必填）
            示例："Sheet1"
        range: 数组公式范围（必填）
            示例："A1:A10"
        formula: 数组公式内容（必填）
            示例："SUM(B1:B10*C1:C10)"

    Returns:
        包含以下字段的字典：
        - success: 是否成功
        - action: 执行的操作
        - sheet: 工作表名称
        - range: 数组公式范围
        - formula: 数组公式内容

    示例：
        set_array_formula(file="data.xlsx", sheet="Sheet1", range="A1:A10", formula="SUM(B1:B10*C1:C10)")
    """
    cap, err = _get_capability("set_array_formula")
    if err:
        return err
    return cap.execute(None, file=file, sheet=sheet, range=range, formula=formula)


# P1 新增工具
@mcp.tool()
def insert_excel_image(
    file: str, sheet: str, cell: str, image_path: str, width: int = None, height: int = None
) -> dict:
    """[商功章] 插入图片到单元格

    使用场景：
    - 需要在 Excel 中插入产品图片
    - 添加 Logo 或水印
    - 嵌入图表截图

    Args:
        file: Excel 文件路径（必填）
            示例："/path/to/data.xlsx"
        sheet: 工作表名称（必填）
            示例："Sheet1"
        cell: 单元格位置（必填）
            示例："A1"
        image_path: 图片文件路径（必填）
            示例："/path/to/image.png"
        width: 图片宽度（像素，可选）
            示例：200
        height: 图片高度（像素，可选）
            示例：150

    Returns:
        包含以下字段的字典：
        - success: 是否成功
        - cell: 单元格位置
        - image_path: 图片路径
        - width: 图片宽度
        - height: 图片高度

    示例：
        insert_excel_image(file="data.xlsx", sheet="Sheet1", cell="A1", image_path="logo.png")
    """
    cap, err = _get_capability("insert_excel_image")
    if err:
        return err
    return cap.execute(
        None, file=file, sheet=sheet, cell=cell, image_path=image_path, width=width, height=height
    )


@mcp.tool()
def group_rows(file: str, sheet: str, start_row: int, end_row: int, level: int = 1) -> dict:
    """[商功章] 分组行（折叠/展开）

    使用场景：
    - 需要折叠明细数据
    - 创建层级结构
    - 报表分组展示

    Args:
        file: Excel 文件路径（必填）
            示例："/path/to/data.xlsx"
        sheet: 工作表名称（必填）
            示例："Sheet1"
        start_row: 起始行号（必填）
            示例：2
        end_row: 结束行号（必填）
            示例：10
        level: 分组层级（可选，默认 1）
            示例：1

    Returns:
        包含以下字段的字典：
        - success: 是否成功
        - start_row: 起始行号
        - end_row: 结束行号
        - level: 分组层级
        - rows_grouped: 分组行数

    示例：
        group_rows(file="data.xlsx", sheet="Sheet1", start_row=2, end_row=10)
    """
    cap, err = _get_capability("group_rows")
    if err:
        return err
    return cap.execute(
        None, file=file, sheet=sheet, start_row=start_row, end_row=end_row, level=level
    )


@mcp.tool()
def subtotal(file: str, sheet: str, range: str, group_column: str, function: str = "sum") -> dict:
    """[衰分章] 分类汇总（按字段分组聚合）

    使用场景：
    - 按类别汇总销售额
    - 按部门统计平均工资
    - 按产品分组计算总量

    Args:
        file: Excel 文件路径（必填）
            示例："/path/to/data.xlsx"
        sheet: 工作表名称（必填）
            示例："Sheet1"
        range: 数据范围，使用 A1 表示法（必填）
            示例："A1:D100"
        group_column: 分组列名（必填）
            示例："Category"
        function: 聚合函数（可选，默认 "sum"）
            可选值：sum、mean、count、min、max
            示例："sum"

    Returns:
        包含以下字段的字典：
        - success: 是否成功
        - group_column: 分组列名
        - function: 聚合函数
        - groups_count: 分组数量
        - summary: 汇总结果

    示例：
        subtotal(file="data.xlsx", sheet="Sheet1", range="A1:D100", group_column="Category", function="sum")
    """
    cap, err = _get_capability("subtotal")
    if err:
        return err
    return cap.execute(
        None, file=file, sheet=sheet, range=range, group_column=group_column, function=function
    )


@mcp.tool()
def transpose(file: str, sheet: str, range: str, output_sheet: str = None) -> dict:
    """[粟米章] 转置数据（行列互换）

    使用场景：
    - 需要将行数据转为列
    - 数据结构调整
    - 报表格式转换

    Args:
        file: Excel 文件路径（必填）
            示例："/path/to/data.xlsx"
        sheet: 工作表名称（必填）
            示例："Sheet1"
        range: 数据范围，使用 A1 表示法（必填）
            示例："A1:D10"
        output_sheet: 输出工作表名称（可选，默认覆盖原工作表）
            示例："Transposed"

    Returns:
        包含以下字段的字典：
        - success: 是否成功
        - source_range: 源范围
        - source_rows: 源行数
        - source_columns: 源列数
        - output_sheet: 输出工作表
        - output_rows: 输出行数
        - output_columns: 输出列数

    示例：
        transpose(file="data.xlsx", sheet="Sheet1", range="A1:D10", output_sheet="Transposed")
    """
    cap, err = _get_capability("transpose")
    if err:
        return err
    return cap.execute(None, file=file, sheet=sheet, range=range, output_sheet=output_sheet)


@mcp.tool()
def text_to_columns(file: str, sheet: str, column: str, delimiter: str = ",") -> dict:
    """[粟米章] 文本分列（按分隔符拆分）

    使用场景：
    - 将逗号分隔的数据拆分到多列
    - 处理 CSV 格式的数据
    - 拆分姓名、地址等复合字段

    Args:
        file: Excel 文件路径（必填）
            示例："/path/to/data.xlsx"
        sheet: 工作表名称（必填）
            示例："Sheet1"
        column: 列标识（必填）
            示例："A"
        delimiter: 分隔符（可选，默认 ","）
            示例：","

    Returns:
        包含以下字段的字典：
        - success: 是否成功
        - column: 源列
        - delimiter: 分隔符
        - rows_split: 拆分行数
        - columns_created: 创建列数

    示例：
        text_to_columns(file="data.xlsx", sheet="Sheet1", column="A", delimiter=",")
    """
    cap, err = _get_capability("text_to_columns")
    if err:
        return err
    return cap.execute(None, file=file, sheet=sheet, column=column, delimiter=delimiter)


@mcp.tool()
def auto_sum(file: str, sheet: str, range: str, direction: str = "down") -> dict:
    """[方程章] 自动求和（在范围内设置 SUM 公式）

    使用场景：
    - 快速为数据列添加合计行
    - 为数据行添加合计列
    - 批量设置求和公式

    Args:
        file: Excel 文件路径（必填）
            示例："/path/to/data.xlsx"
        sheet: 工作表名称（必填）
            示例："Sheet1"
        range: 数据范围，使用 A1 表示法（必填）
            示例："A1:D10"
        direction: 求和方向（可选，默认 "down"）
            可选值：down（向下求和）、right（向右求和）
            示例："down"

    Returns:
        包含以下字段的字典：
        - success: 是否成功
        - range: 数据范围
        - direction: 求和方向
        - formulas_set: 设置的公式数量

    示例：
        auto_sum(file="data.xlsx", sheet="Sheet1", range="A1:D10", direction="down")
    """
    cap, err = _get_capability("auto_sum")
    if err:
        return err
    return cap.execute(None, file=file, sheet=sheet, range=range, direction=direction)


@mcp.tool()
def export_chart_as_image(file: str, sheet: str, chart_index: int, output: str) -> dict:
    """[商功章] 导出图表为图片（使用 LibreOffice）

    使用场景：
    - 需要将 Excel 图表导出为 PNG 图片
    - 在报告中嵌入图表
    - 分享图表图片

    Args:
        file: Excel 文件路径（必填）
            示例："/path/to/data.xlsx"
        sheet: 工作表名称（必填）
            示例："Sheet1"
        chart_index: 图表索引（必填，从 0 开始）
            示例：0
        output: 输出图片路径（必填）
            示例："/path/to/chart.png"

    Returns:
        包含以下字段的字典：
        - success: 是否成功
        - file: Excel 文件路径
        - chart_index: 图表索引
        - output: 输出图片路径
        - note: 备注信息

    示例：
        export_chart_as_image(file="data.xlsx", sheet="Sheet1", chart_index=0, output="chart.png")
    """
    cap, err = _get_capability("export_chart_as_image")
    if err:
        return err
    return cap.execute(None, file=file, sheet=sheet, chart_index=chart_index, output=output)


# P3 新增工具
@mcp.tool()
def pack_file(file: str, output: str | None = None) -> dict:
    """[商功章] 将 Excel 文件打包为 ZIP

    使用场景：
    - 需要将 Excel 文件打包为 ZIP
    - 用于调试和分析
    - 文件备份

    Args:
        file: Excel 文件路径（必填）
            示例："/path/to/data.xlsx"
        output: 输出 ZIP 文件路径（可选）
            示例："/path/to/output.zip"

    Returns:
        包含以下字段的字典：
        - file: Excel 文件路径
        - output: 输出 ZIP 文件路径
        - packed: 是否打包成功

    示例：
        pack_file(file="data.xlsx")
        pack_file(file="data.xlsx", output="backup.zip")
    """
    cap, err = _get_capability("pack_file")
    if err:
        return err
    return cap.execute(None, file=file, output=output)


@mcp.tool()
def unpack_file(file: str, output: str | None = None) -> dict:
    """[商功章] 将 ZIP 解包为 Excel 文件

    使用场景：
    - 需要将 ZIP 解包为 Excel 文件
    - 用于调试和分析
    - 文件恢复

    Args:
        file: ZIP 文件路径（必填）
            示例："/path/to/file.zip"
        output: 输出目录（可选）
            示例："/path/to/output"

    Returns:
        包含以下字段的字典：
        - file: ZIP 文件路径
        - output: 输出目录
        - files: 解包的文件列表
        - unpacked: 是否解包成功

    示例：
        unpack_file(file="backup.zip")
        unpack_file(file="backup.zip", output="output_dir")
    """
    cap, err = _get_capability("unpack_file")
    if err:
        return err
    return cap.execute(None, file=file, output=output)


@mcp.tool()
def set_print_area(file: str, sheet: str, range: str) -> dict:
    """[商功章] 设置打印区域

    使用场景：
    - 需要设置打印区域
    - 控制打印范围
    - 优化打印布局

    Args:
        file: Excel 文件路径（必填）
            示例："/path/to/data.xlsx"
        sheet: 工作表名称（必填）
            示例："Sheet1"
        range: 打印区域（必填，如 A1:D10）
            示例："A1:D10"

    Returns:
        包含以下字段的字典：
        - file: Excel 文件路径
        - sheet: 工作表名称
        - print_area: 打印区域
        - set: 是否设置成功

    示例：
        set_print_area(file="data.xlsx", sheet="Sheet1", range="A1:D10")
    """
    cap, err = _get_capability("set_print_area")
    if err:
        return err
    return cap.execute(None, file=file, sheet=sheet, range=range)


@mcp.tool()
def set_zoom(file: str, sheet: str, zoom: int) -> dict:
    """[商功章] 控制工作表缩放（10-400%）

    使用场景：
    - 需要调整工作表缩放比例
    - 优化查看体验
    - 控制显示大小

    Args:
        file: Excel 文件路径（必填）
            示例："/path/to/data.xlsx"
        sheet: 工作表名称（必填）
            示例："Sheet1"
        zoom: 缩放比例（必填，10-400）
            示例：100

    Returns:
        包含以下字段的字典：
        - file: Excel 文件路径
        - sheet: 工作表名称
        - zoom: 缩放比例
        - set: 是否设置成功

    示例：
        set_zoom(file="data.xlsx", sheet="Sheet1", zoom=100)
    """
    cap, err = _get_capability("set_zoom")
    if err:
        return err
    return cap.execute(None, file=file, sheet=sheet, zoom=zoom)


# ============================================================================
# 知识图谱 MCP 工具
# ============================================================================


@mcp.tool()
def abacus_skill_search(query: str, limit: int = 10) -> dict:
    """[知识图谱] 搜索 SKILL.md 和知识文件

    使用 FTS5 全文搜索，返回匹配的 skill 和知识文件。

    Args:
        query: 搜索关键词
        limit: 返回结果数量（默认 10）

    Returns:
        包含搜索结果的字典：
        - results: 结果列表，每个结果包含 type/name/description/file_path
        - total: 总结果数

    示例：
        abacus_skill_search(query="公式验证")
        abacus_skill_search(query="chart", limit=5)
    """
    from pathlib import Path

    from .skill.indexer import SkillIndexer

    indexer = SkillIndexer(Path("skills_index.db"))
    try:
        results = indexer.search(query, limit)
        return {"query": query, "results": results, "total": len(results)}
    finally:
        indexer.close()


@mcp.tool()
def abacus_skill_graph(skill_name: str) -> dict:
    """[知识图谱] 获取 skill 关联图谱

    返回指定 skill 的完整关联信息：
    - 使用的能力
    - 引用的知识文件
    - 关联的其他 skill

    Args:
        skill_name: skill 名称（如 "abacus-field"）

    Returns:
        包含图谱信息的字典：
        - skill: skill 名称
        - chapter: 所属章节
        - description: 描述
        - uses: 使用的能力列表
        - referenced_by: 引用的知识文件
        - related_skills: 关联的 skill

    示例：
        abacus_skill_graph(skill_name="abacus-field")
    """
    from pathlib import Path

    from .skill.indexer import SkillIndexer

    indexer = SkillIndexer(Path("skills_index.db"))
    try:
        return indexer.graph(skill_name)
    finally:
        indexer.close()


@mcp.tool()
def abacus_skill_index_build() -> dict:
    """[知识图谱] 重建 SKILL.md 索引

    扫描 skills/ 目录，重建知识图谱索引。
    包括：9 个章节 SKILL.md + 29 个知识文件。

    Returns:
        包含索引统计的字典：
        - status: 状态
        - stats: 索引统计（skills/knowledge/relations 数量）

    示例：
        abacus_skill_index_build()
    """
    from pathlib import Path

    from .skill.indexer import SkillIndexer

    indexer = SkillIndexer(Path("skills_index.db"))
    try:
        indexer.scan(Path("skills"))
        stats = indexer.stats()
        return {"status": "ok", "stats": stats}
    finally:
        indexer.close()


@mcp.tool()
def abacus_skill_stats() -> dict:
    """[知识图谱] 获取索引统计

    返回知识图谱的统计信息。

    Returns:
        包含统计的字典：
        - skills: skill 数量
        - knowledge_files: 知识文件数量
        - relations: 关系数量
        - chapters: 各章节 skill 分布
        - sources: 知识文件来源分布
    """
    from pathlib import Path

    from .skill.indexer import SkillIndexer

    indexer = SkillIndexer(Path("skills_index.db"))
    try:
        return indexer.stats()
    finally:
        indexer.close()


# ============================================================================
# P2 新增工具 - 智能数据匹配和关联
# ============================================================================


@mcp.tool()
def fuzzy_match_columns(
    file: str, sheet: str, target_columns: list[str], threshold: float = 0.6
) -> dict:
    """[粟米章] 模糊匹配列名（自动识别相似列名）

    使用场景：
    - 自动识别相似列名（如"销售额"和"销售金额"）
    - 多表合并时自动对齐列名
    - 数据整合时的列名匹配

    Args:
        file: Excel 文件路径（必填）
            示例："/path/to/data.xlsx"
        sheet: 工作表名称（必填）
            示例："Sheet1"
        target_columns: 目标列名列表（必填）
            示例：["销售额", "利润"]
        threshold: 相似度阈值（可选，默认 0.6）
            示例：0.6

    Returns:
        包含以下字段的字典：
        - source_columns: 源列名列表
        - matches: 匹配结果列表
        - match_count: 匹配成功的数量

    示例：
        fuzzy_match_columns(file="data.xlsx", sheet="Sheet1", target_columns=["销售额", "利润"])
    """
    cap, err = _get_capability("fuzzy_match_columns")
    if err:
        return err
    return cap.execute(
        None, file=file, sheet=sheet, target_columns=target_columns, threshold=threshold
    )


@mcp.tool()
def data_quality_check(file: str, sheet: str, range: str | None = None) -> dict:
    """[盈不足章] 数据质量检测（自动检测空值、异常值、重复数据）

    使用场景：
    - 自动检测空值、异常值、重复数据、格式不一致
    - 数据清洗前的质量评估
    - 数据质量报告

    Args:
        file: Excel 文件路径（必填）
            示例："/path/to/data.xlsx"
        sheet: 工作表名称（必填）
            示例："Sheet1"
        range: 数据范围（可选，默认全部）
            示例："A1:D100"

    Returns:
        包含以下字段的字典：
        - quality_score: 质量分数（0-100）
        - issues: 问题列表
        - issue_count: 问题数量
        - null_counts: 空值统计
        - type_distribution: 类型分布

    示例：
        data_quality_check(file="data.xlsx", sheet="Sheet1")
    """
    cap, err = _get_capability("data_quality_check")
    if err:
        return err
    return cap.execute(None, file=file, sheet=sheet, range=range)


@mcp.tool()
def join_tables(
    left_file: str,
    left_sheet: str,
    right_file: str,
    right_sheet: str,
    on: list[str],
    how: str = "inner",
    output: str | None = None,
) -> dict:
    """[均输章] SQL 风格关联（LEFT/RIGHT/INNER/OUTER JOIN）

    使用场景：
    - 替代 VLOOKUP，多表关联
    - LEFT/RIGHT/INNER/OUTER JOIN
    - 数据整合

    Args:
        left_file: 左表 Excel 文件路径（必填）
            示例："/path/to/left.xlsx"
        left_sheet: 左表工作表名称（必填）
            示例："Sheet1"
        right_file: 右表 Excel 文件路径（必填）
            示例："/path/to/right.xlsx"
        right_sheet: 右表工作表名称（必填）
            示例："Sheet1"
        on: 关联键（必填）
            示例：["ID"]
        how: 关联类型（可选，默认 "inner"）
            可选值：left、right、inner、outer
            示例："inner"
        output: 输出文件路径（可选）
            示例："/path/to/output.xlsx"

    Returns:
        包含以下字段的字典：
        - success: 是否成功
        - left_rows: 左表行数
        - right_rows: 右表行数
        - result_rows: 结果行数
        - result_columns: 结果列名列表

    示例：
        join_tables(left_file="left.xlsx", left_sheet="Sheet1", right_file="right.xlsx", right_sheet="Sheet1", on=["ID"], how="inner")
    """
    cap, err = _get_capability("join_tables")
    if err:
        return err
    return cap.execute(
        None,
        left_file=left_file,
        left_sheet=left_sheet,
        right_file=right_file,
        right_sheet=right_sheet,
        on=on,
        how=how,
        output=output,
    )


@mcp.tool()
def batch_merge(
    folder: str, pattern: str = "*.xlsx", sheet: str | None = None, output: str = None
) -> dict:
    """[均输章] 多表批量合并（从文件夹批量合并多个 Excel 文件）

    使用场景：
    - 年终汇总、月度数据合并
    - 从文件夹批量合并多个 Excel 文件
    - 数据整合

    Args:
        folder: 文件夹路径（必填）
            示例："/path/to/data"
        pattern: 文件匹配模式（可选，默认 "*.xlsx")
            示例："*.xlsx"
        sheet: 工作表名称（可选）
            示例："Sheet1"
        output: 输出文件路径（必填）
            示例："/path/to/output.xlsx"

    Returns:
        包含以下字段的字典：
        - success: 是否成功
        - file_count: 文件数量
        - success_count: 成功数量
        - total_rows: 总行数
        - columns: 列名列表

    示例：
        batch_merge(folder="/path/to/data", pattern="*.xlsx", output="merged.xlsx")
    """
    cap, err = _get_capability("batch_merge")
    if err:
        return err
    return cap.execute(None, folder=folder, pattern=pattern, sheet=sheet, output=output)


@mcp.tool()
def auto_type_infer(
    file: str, sheet: str, range: str | None = None, output: str | None = None
) -> dict:
    """[粟米章] 自动类型推断（自动检测并转换数据类型）

    使用场景：
    - 自动检测并转换数据类型（文本→数字、日期等）
    - 导入数据时自动标准化
    - 数据清洗

    Args:
        file: Excel 文件路径（必填）
            示例："/path/to/data.xlsx"
        sheet: 工作表名称（必填）
            示例："Sheet1"
        range: 数据范围（可选，默认全部）
            示例："A1:D100"
        output: 输出文件路径（可选）
            示例："/path/to/output.xlsx"

    Returns:
        包含以下字段的字典：
        - success: 是否成功
        - inferred_types: 推断的类型
        - conversions: 转换结果
        - conversion_count: 转换数量

    示例：
        auto_type_infer(file="data.xlsx", sheet="Sheet1")
    """
    cap, err = _get_capability("auto_type_infer")
    if err:
        return err
    return cap.execute(None, file=file, sheet=sheet, range=range, output=output)


@mcp.tool()
def standardize_data(
    file: str,
    sheet: str,
    date_format: str | None = None,
    number_format: str | None = None,
    text_case: str | None = None,
    output: str | None = None,
) -> dict:
    """[粟米章] 数据标准化（统一日期、数字、文本格式）

    使用场景：
    - 统一日期格式（如 %Y-%m-%d）
    - 统一数字格式（如 %.2f）
    - 统一文本大小写（lower/upper/title）

    Args:
        file: Excel 文件路径（必填）
            示例："/path/to/data.xlsx"
        sheet: 工作表名称（必填）
            示例："Sheet1"
        date_format: 日期格式（可选）
            示例："%Y-%m-%d"
        number_format: 数字格式（可选）
            示例："%.2f"
        text_case: 文本大小写（可选）
            可选值：lower、upper、title
            示例："lower"
        output: 输出文件路径（可选）
            示例："/path/to/output.xlsx"

    Returns:
        包含以下字段的字典：
        - success: 是否成功
        - operations: 操作列表
        - operation_count: 操作数量

    示例：
        standardize_data(file="data.xlsx", sheet="Sheet1", date_format="%Y-%m-%d", text_case="lower")
    """
    cap, err = _get_capability("standardize_data")
    if err:
        return err
    return cap.execute(
        None,
        file=file,
        sheet=sheet,
        date_format=date_format,
        number_format=number_format,
        text_case=text_case,
        output=output,
    )


@mcp.tool()
def transform_pipeline(
    file: str,
    steps: list,
    sheet: str | None = None,
    stop_on_error: bool = True,
) -> dict:
    """[粟米章] 数据转换管道（链式执行多个转换步骤）

    使用场景：
    - 需要按顺序执行多个数据转换
    - 数据清洗流水线
    - 批量格式化和类型转换

    支持的步骤类型：
    - convert_type: 类型转换（target_type: int/float/str）
    - convert_format: 格式转换（format_type: number/currency/percentage/date/text）
    - convert_unit: 单位转换（factor: 转换系数）
    - standardize: 标准化（text_case: upper/lower/title）
    - fill_value: 填充值（value: 填充值）
    - replace_value: 替换值（old_value/new_value）

    Args:
        file: Excel 文件路径（必填）
            示例："/path/to/data.xlsx"
        steps: 转换步骤列表（必填）
            示例：[
                {"type": "convert_type", "range": "A1:A10", "target_type": "float"},
                {"type": "fill_value", "range": "B1:B10", "value": 0}
            ]
        sheet: 工作表名称（可选，默认活动工作表）
            示例："Sheet1"
        stop_on_error: 遇到错误是否停止（可选，默认 True）
            示例：True

    Returns:
        包含以下字段的字典：
        - file: 文件路径
        - sheet: 工作表名称
        - steps_executed: 执行的步骤数
        - steps_succeeded: 成功的步骤数
        - results: 每个步骤的结果

    示例：
        transform_pipeline(file="data.xlsx", steps=[{"type": "convert_type", "range": "A1:A10", "target_type": "float"}])
    """
    cap, err = _get_capability("transform_pipeline")
    if err:
        return err
    return cap.execute(None, file=file, sheet=sheet, steps=steps, stop_on_error=stop_on_error)


@mcp.tool()
def generate_summary_report(file: str, sheet: str) -> dict:
    """[商功章] 数据摘要报告（自动生成数据摘要）

    使用场景：
    - 快速了解数据全貌
    - 自动生成数据摘要（行列数、类型分布、质量问题）
    - 数据探索

    Args:
        file: Excel 文件路径（必填）
            示例："/path/to/data.xlsx"
        sheet: 工作表名称（必填）
            示例："Sheet1"

    Returns:
        包含以下字段的字典：
        - total_rows: 总行数
        - total_columns: 总列数
        - columns: 列名列表
        - dtypes: 数据类型
        - null_counts: 空值统计
        - numeric_stats: 数值统计
        - categorical_stats: 分类统计

    示例：
        generate_summary_report(file="data.xlsx", sheet="Sheet1")
    """
    cap, err = _get_capability("generate_summary_report")
    if err:
        return err
    return cap.execute(None, file=file, sheet=sheet)


@mcp.tool()
def generate_diff_report(
    old_file: str,
    old_sheet: str,
    new_file: str,
    new_sheet: str,
    key_columns: list[list[str]] = None,
) -> dict:
    """[商功章] 变化检测报告（对比两个版本的数据，检测变化）

    使用场景：
    - 数据版本对比
    - 审计追踪
    - 检测数据变化

    Args:
        old_file: 旧版本文件路径（必填）
            示例："/path/to/old.xlsx"
        old_sheet: 旧版本工作表名称（必填）
            示例："Sheet1"
        new_file: 新版本文件路径（必填）
            示例："/path/to/new.xlsx"
        new_sheet: 新版本工作表名称（必填）
            示例："Sheet1"
        key_columns: 用于匹配的键列（可选）
            示例：["ID"]

    Returns:
        包含以下字段的字典：
        - old_rows: 旧行数
        - new_rows: 新行数
        - row_diff: 行数差异
        - added_columns: 新增列
        - removed_columns: 删除列
        - changes: 变化详情

    示例：
        generate_diff_report(old_file="old.xlsx", old_sheet="Sheet1", new_file="new.xlsx", new_sheet="Sheet1", key_columns=["ID"])
    """
    cap, err = _get_capability("generate_diff_report")
    if err:
        return err
    return cap.execute(
        None,
        old_file=old_file,
        old_sheet=old_sheet,
        new_file=new_file,
        new_sheet=new_sheet,
        key_columns=key_columns,
    )


@mcp.tool()
def manage_data_view(
    file: str,
    sheet: str,
    action: str,
    view_name: str | None = None,
    columns: list[list[str]] = None,
    filters: list | None = None,
) -> dict:
    """[商功章] 数据视图管理（创建和管理不同角色的数据视图）

    使用场景：
    - 创建和管理不同角色的数据视图
    - 不同角色看不同数据子集
    - 数据权限管理

    Args:
        file: Excel 文件路径（必填）
            示例："/path/to/data.xlsx"
        sheet: 工作表名称（必填）
            示例："Sheet1"
        action: 操作（必填）
            可选值：create（创建）、（列出）、get（获取）、delete（删除）
            示例："create"
        view_name: 视图名称（create//get/delete 时必填）
            示例："sales_view"
        columns: 视图包含的列（create 时必填）
            示例：["ID", "Name", "Sales"]
        filters: 过滤条件（可选）
            示例：{"Region": ["East", "West"]}

    Returns:
        包含以下字段的字典：
        - success: 是否成功
        - view_name: 视图名称
        - views: 视图列表（ 时）
        - data: 数据（get 时）

    示例：
        manage_data_view(file="data.xlsx", sheet="Sheet1", action="create", view_name="sales_view", columns=["ID", "Name", "Sales"])
    """
    cap, err = _get_capability("manage_data_view")
    if err:
        return err
    return cap.execute(
        None,
        file=file,
        sheet=sheet,
        action=action,
        view_name=view_name,
        columns=columns,
        filters=filters,
    )


if __name__ == "__main__":
    mcp.run()
