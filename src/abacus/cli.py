"""CLI 入口 - 完整实现"""

import json
from pathlib import Path

from .logging import setup_logging

setup_logging()

import click

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
    AutoTypeInferCapability,
    ConvertFormatCapability,
    ConvertTypeCapability,
    ConvertUnitCapability,
    DataTransformCapability,
    FuzzyMatchCapability,
    StandardizeCapability,
    TextToColumnsCapability,
    TransformPipelineCapability,
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
    ExportChartAsImageCapability,
    GroupRowsCapability,
    InsertImageCapability,
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
from .core.work.format import FormatRangeCapability
from .core.work.freeze import FreezePaneCapability
from .core.work.hide_show import HideShowCapability
from .core.work.mapping_template import CreateMappingTemplateCapability
from .core.work.pack import PackFileCapability
from .core.work.protection import (
    ProtectSheetCapability,
    ProtectWorkbookCapability,
    SetArrayFormulaCapability,
)
from .core.work.unpack import UnpackFileCapability
from .core.workflow import FormattingWorkflowCapability, SpreadsheetWorkflowCapability

# 初始化注册表
registry = CapabilityRegistry()

# 方田章
registry.register(MeasureRangeCapability())
registry.register(MeasureCellsCapability())
registry.register(MeasureStructureCapability())
registry.register(NamedRangeCapability())

# 粟米章
registry.register(ConvertFormatCapability())
registry.register(ConvertUnitCapability())
registry.register(ConvertTypeCapability())
registry.register(TransposeCapability())
registry.register(TextToColumnsCapability())
registry.register(FuzzyMatchCapability())
registry.register(AutoTypeInferCapability())
registry.register(StandardizeCapability())
registry.register(TransformPipelineCapability())

# 衰分章
registry.register(GroupByCapability())
registry.register(DistributeCapability())
registry.register(SummarizeCapability())
registry.register(SubtotalCapability())

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
registry.register(InsertImageCapability())
registry.register(GroupRowsCapability())
registry.register(ExportChartAsImageCapability())
registry.register(ProtectWorkbookCapability())
registry.register(ProtectSheetCapability())
registry.register(SetArrayFormulaCapability())
registry.register(PackFileCapability())
registry.register(UnpackFileCapability())
registry.register(AdvancedFilterCapability())
registry.register(CreateMappingTemplateCapability())
registry.register(SummaryReportCapability())
registry.register(DiffReportCapability())
registry.register(DataViewCapability())

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


@click.group()
@click.version_option(version="1.0.0")
def main():
    """Abacus - 本地 Excel 自动化框架（算盘）

    基于《九章算术》的章节体系设计，提供 Excel 数据读取、格式转换、分组汇总、
    反向计算、批量操作、导入导出、数据验证、公式计算、数据分析等完整能力。

    使用 abacus capabilities 查看所有可用能力。
    """
    pass


@main.command()
def capabilities():
    """[方田章] 列出所有能力

    使用场景：
    - 查看 Abacus 框架支持的所有功能模块
    - 了解各章节（方田、粟米、衰分等）的能力分布
    - 开发调试时确认能力注册状态

    示例：
        abacus capabilities
    """
    click.echo("可用能力 (Capabilities):")
    click.echo("=" * 50)

    chapters = {}
    for cap in registry.list_all():
        if cap.chapter not in chapters:
            chapters[cap.chapter] = []
        chapters[cap.chapter].append(cap)

    chapter_names = {
        "field": "方田章 (Field) - 数据读取",
        "grain": "粟米章 (Grain) - 格式转换",
        "share": "衰分章 (Share) - 分组汇总",
        "dimension": "少广章 (Dimension) - 反向计算",
        "work": "商功章 (Work) - 批量操作",
        "transport": "均输章 (Transport) - 导入导出",
        "balance": "盈不足章 (Balance) - 数据验证",
        "equation": "方程章 (Equation) - 公式计算",
        "triangle": "勾股章 (Triangle) - 数据分析",
    }

    for chapter, caps in chapters.items():
        click.echo(f"\n{chapter_names.get(chapter, chapter)}:")
        for cap in caps:
            click.echo(f"  - {cap.name}: {cap.description}")


# 方田章命令
@main.command()
@click.option("--file", "-f", required=True, help="Excel 文件路径（如 data.xlsx）")
@click.option("--sheet", "-s", required=True, help="工作表名称（如 Sheet1）")
@click.option("--range", "-r", required=True, help="数据范围，Excel 格式（如 A1:D10）")
def read(file: str, sheet: str, range: str):
    """[方田章] 读取指定范围数据

    使用场景：
    - 需要读取 Excel 文件中的数据
    - 需要查看特定工作表的特定范围
    - 数据分析、报表生成的第一步

    示例：
        abacus read -f data.xlsx -s Sales -r A1:C10
    """
    cap = registry.get("measure_range")
    result = cap.execute(None, file=file, sheet=sheet, range=range)
    click.echo(json.dumps(result, ensure_ascii=False, indent=2))


@main.command()
@click.option("--file", "-f", required=True, help="Excel 文件路径（如 data.xlsx）")
@click.option("--sheet", "-s", required=True, help="工作表名称（如 Sheet1）")
@click.option("--range", "-r", required=True, help="数据范围，Excel 格式（如 A1:D10）")
def cells(file: str, sheet: str, range: str):
    """[方田章] 读取单元格详细信息

    使用场景：
    - 需要查看单元格的值、类型、公式等详细信息
    - 调试公式错误时检查单元格状态
    - 批量检查数据格式一致性

    示例：
        abacus cells -f data.xlsx -s Sheet1 -r A1:C5
    """
    cap = registry.get("measure_cells")
    result = cap.execute(None, file=file, sheet=sheet, range=range)
    click.echo(json.dumps(result, ensure_ascii=False, indent=2))


@main.command()
@click.option("--file", "-f", required=True, help="Excel 文件路径（如 data.xlsx）")
def structure(file: str):
    """[方田章] 查看工作表结构

    使用场景：
    - 快速了解 Excel 文件包含哪些工作表
    - 查看每个工作表的行列数、数据范围
    - 数据导入前的结构检查

    示例：
        abacus structure -f data.xlsx
    """
    cap = registry.get("measure_structure")
    result = cap.execute(None, file=file)
    click.echo(json.dumps(result, ensure_ascii=False, indent=2))


@main.command()
@click.option("--file", "-f", required=True, help="Excel 文件路径（如 data.xlsx）")
def sheets(file: str):
    """[方田章] 列出所有工作表名称

    使用场景：
    - 快速获取 Excel 文件包含哪些工作表
    - 获取工作表名称后进行精准读取
    - 批量操作前确认工作表列表

    示例：
        abacus sheets -f data.xlsx
    """
    cap = registry.get("list_sheets")
    result = cap.execute(None, file=file)
    click.echo(json.dumps(result, ensure_ascii=False, indent=2))


@main.command()
@click.option("--file", "-f", required=True, help="Excel 文件路径（如 data.xlsx）")
@click.option("--rows", "-n", default=5, help="预览行数（默认 5）")
@click.option("--sheet", "-s", help="工作表名称（可选，默认预览所有）")
def peek(file: str, rows: int, sheet: str):
    """[方田章] 快速预览数据

    使用场景：
    - 快速了解数据长什么样
    - 查看表头和前几行数据
    - 判断数据格式和内容

    示例：
        abacus peek -f data.xlsx
        abacus peek -f data.xlsx -n 3
        abacus peek -f data.xlsx -s Sales
    """
    cap = registry.get("peek_preview")
    result = cap.execute(None, file=file, rows=rows, sheet=sheet)
    click.echo(json.dumps(result, ensure_ascii=False, indent=2))


@main.command()
@click.option("--file", "-f", required=True, help="Excel 文件路径（如 data.xlsx）")
@click.option("--sheet", "-s", required=True, help="工作表名称（如 Sheet1）")
@click.option("--sample-rows", "-n", default=100, help="采样行数（默认 100）")
def columns(file: str, sheet: str, sample_rows: int):
    """[方田章] 检测列名和数据类型

    使用场景：
    - 了解有哪些列可用
    - 检测每列的数据类型
    - 为后续操作选择合适的列

    示例：
        abacus columns -f data.xlsx -s Sheet1
        abacus columns -f data.xlsx -s Sales --sample-rows 200
    """
    cap = registry.get("detect_columns")
    result = cap.execute(None, file=file, sheet=sheet, sample_rows=sample_rows)
    click.echo(json.dumps(result, ensure_ascii=False, indent=2))


@main.command()
@click.option("--file", "-f", required=True, help="Excel 文件路径（如 data.xlsx）")
@click.option("--keyword", "-k", required=True, help="搜索关键词")
@click.option("--sheet", "-s", help="工作表名称（可选，默认搜索所有）")
@click.option("--max-results", "-m", default=50, help="最大结果数（默认 50）")
def search(file: str, keyword: str, sheet: str, max_results: int):
    """[方田章] 搜索内容

    使用场景：
    - 快速定位包含特定内容的单元格
    - 查找特定文本、数字或公式
    - 数据检索和验证

    示例：
        abacus search -f data.xlsx -k 销售
        abacus search -f data.xlsx -k "产品A" -s Sales
    """
    cap = registry.get("search_content")
    result = cap.execute(None, file=file, keyword=keyword, sheet=sheet, max_results=max_results)
    click.echo(json.dumps(result, ensure_ascii=False, indent=2))


@main.command()
@click.option("--file", "-f", required=True, help="Excel 文件路径（如 data.xlsx）")
def summary(file: str):
    """[方田章] 获取文件摘要

    使用场景：
    - 快速了解文件规模
    - 获取每个工作表的行列数
    - 评估数据量大小

    示例：
        abacus summary -f data.xlsx
    """
    cap = registry.get("get_summary")
    result = cap.execute(None, file=file)
    click.echo(json.dumps(result, ensure_ascii=False, indent=2))


@main.command()
@click.option("--file", "-f", required=True, help="Excel 文件路径（如 data.xlsx）")
@click.option("--sheet", "-s", required=True, help="工作表名称（如 Sheet1）")
@click.option("--rows", "-n", default=10, help="样本行数（默认 10）")
def sample(file: str, sheet: str, rows: int):
    """[方田章] 获取样本数据

    使用场景：
    - 精准预览某个工作表的数据
    - 获取前 N 行数据用于分析
    - 数据探索和验证

    示例：
        abacus sample -f data.xlsx -s Sheet1
        abacus sample -f data.xlsx -s Sales -n 20
    """
    cap = registry.get("get_sample_data")
    result = cap.execute(None, file=file, sheet=sheet, rows=rows)
    click.echo(json.dumps(result, ensure_ascii=False, indent=2))


@main.command()
@click.option("--file", "-f", required=True, help="Excel 文件路径（如 data.xlsx）")
@click.option("--sheet", "-s", required=True, help="工作表名称（如 Sheet1）")
@click.option("--range", "-r", required=True, help="数据范围，Excel 格式（如 A1:D10）")
@click.option(
    "--format-type",
    required=True,
    type=click.Choice(["date", "number", "text", "percentage", "currency"]),
    help="目标格式类型",
)
def convert_format(file: str, sheet: str, range: str, format_type: str):
    """[粟米章] 转换数据格式

    使用场景：
    - 将文本格式的数字转换为数值格式
    - 将字符串日期转换为日期格式
    - 统一数据格式以便后续计算和分析

    示例：
        abacus convert-format -f data.xlsx -s Sheet1 -r A1:A100 --format-type number
        abacus convert-format -f data.xlsx -s Sales -r B1:B50 --format-type date
    """
    cap = registry.get("convert_format")
    result = cap.execute(None, file=file, sheet=sheet, range=range, format_type=format_type)
    click.echo(json.dumps(result, ensure_ascii=False, indent=2))


@main.command()
@click.option("--file", "-f", required=True, help="Excel 文件路径（如 data.xlsx）")
@click.option("--sheet", "-s", required=True, help="工作表名称（如 Sheet1）")
@click.option("--range", "-r", required=True, help="数据范围，Excel 格式（如 A1:D10）")
@click.option("--from-unit", required=True, help="源单位（如 cm、kg、USD）")
@click.option("--to-unit", required=True, help="目标单位（如 m、g、CNY）")
def convert_unit(file: str, sheet: str, range: str, from_unit: str, to_unit: str):
    """[粟米章] 转换数据单位

    使用场景：
    - 将厘米转换为米、千克转换为克等
    - 货币单位转换（需配合汇率数据）
    - 统一数据单位以便比较和汇总

    示例：
        abacus convert-unit -f data.xlsx -s Sheet1 -r B1:B50 --from-unit cm --to-unit m
        abacus convert-unit -f data.xlsx -s Sales -r D1:D100 --from-unit USD --to-unit CNY
    """
    cap = registry.get("convert_unit")
    result = cap.execute(
        None, file=file, sheet=sheet, range=range, from_unit=from_unit, to_unit=to_unit
    )
    click.echo(json.dumps(result, ensure_ascii=False, indent=2))


@main.command()
@click.option("--file", "-f", required=True, help="Excel 文件路径（如 data.xlsx）")
@click.option("--sheet", "-s", required=True, help="工作表名称（如 Sheet1）")
@click.option("--range", "-r", required=True, help="数据范围，Excel 格式（如 A1:D10）")
@click.option(
    "--target-type",
    required=True,
    type=click.Choice(["int", "float", "str", "date"]),
    help="目标数据类型",
)
def convert_type(file: str, sheet: str, range: str, target_type: str):
    """[粟米章] 转换数据类型

    使用场景：
    - 将字符串数字转换为整数或浮点数
    - 将数值转换为字符串以便拼接
    - 统一数据类型以避免计算错误

    示例：
        abacus convert-type -f data.xlsx -s Sheet1 -r A1:A100 --target-type int
        abacus convert-type -f data.xlsx -s Sheet1 -r B1:B50 --target-type float
    """
    cap = registry.get("convert_type")
    result = cap.execute(None, file=file, sheet=sheet, range=range, target_type=target_type)
    click.echo(json.dumps(result, ensure_ascii=False, indent=2))


# 衰分章命令
@main.command()
@click.option("--file", "-f", required=True, help="Excel 文件路径（如 data.xlsx）")
@click.option("--sheet", "-s", required=True, help="工作表名称（如 Sheet1）")
@click.option("--range", "-r", required=True, help="数据范围，Excel 格式（如 A1:D100）")
@click.option("--group-column", required=True, help="分组列名（如 Category、Region）")
def group_by(file: str, sheet: str, range: str, group_column: str):
    """[衰分章] 按字段分组数据

    使用场景：
    - 按类别、地区、时间等维度对数据进行分组
    - 为后续的汇总统计做准备
    - 数据探索和分组分析

    示例：
        abacus group-by -f data.xlsx -s Sales -r A1:E100 --group-column Category
        abacus group-by -f data.xlsx -s Orders -r A1:F500 --group-column Region
    """
    cap = registry.get("group_by")
    result = cap.execute(None, file=file, sheet=sheet, range=range, group_column=group_column)
    click.echo(json.dumps(result, ensure_ascii=False, indent=2))


@main.command()
@click.option("--file", "-f", required=True, help="Excel 文件路径（如 data.xlsx）")
@click.option("--sheet", "-s", required=True, help="工作表名称（如 Sheet1）")
@click.option("--range", "-r", required=True, help="数据范围，Excel 格式（如 A1:D100）")
@click.option("--total", required=True, type=float, help="待分配的总数值（如 10000）")
@click.option(
    "--method",
    default="equal",
    type=click.Choice(["equal", "proportional"]),
    help="分配方法：equal=平均分配，proportional=按比例分配",
)
def distribute(file: str, sheet: str, range: str, total: float, method: str):
    """[衰分章] 按比例分配数值

    使用场景：
    - 将预算按部门人数比例分配
    - 将销售额按区域权重分配
    - 将成本按产品线分配

    示例：
        abacus distribute -f data.xlsx -s Budget -r A1:C10 --total 100000 --method proportional
        abacus distribute -f data.xlsx -s Staff -r A1:B20 --total 50000 --method equal
    """
    cap = registry.get("distribute")
    result = cap.execute(None, file=file, sheet=sheet, range=range, total=total, method=method)
    click.echo(json.dumps(result, ensure_ascii=False, indent=2))


@main.command()
@click.option("--file", "-f", required=True, help="Excel 文件路径（如 data.xlsx）")
@click.option("--sheet", "-s", required=True, help="工作表名称（如 Sheet1）")
@click.option("--range", "-r", required=True, help="数据范围，Excel 格式（如 A1:D100）")
@click.option("--group-column", required=True, help="分组列名（如 Category、Region）")
@click.option(
    "--agg-function",
    default="sum",
    type=click.Choice(["sum", "avg", "count", "min", "max"]),
    help="聚合函数",
)
def summarize(file: str, sheet: str, range: str, group_column: str, agg_function: str):
    """[衰分章] 分组汇总统计

    使用场景：
    - 按类别汇总销售额
    - 按地区统计平均价格
    - 按时间分组计算最大/最小值

    示例：
        abacus summarize -f data.xlsx -s Sales -r A1:E100 --group-column Category --agg-function sum
        abacus summarize -f data.xlsx -s Orders -r A1:F500 --group-column Region --agg-function avg
    """
    cap = registry.get("summarize")
    result = cap.execute(
        None,
        file=file,
        sheet=sheet,
        range=range,
        group_column=group_column,
        agg_function=agg_function,
    )
    click.echo(json.dumps(result, ensure_ascii=False, indent=2))


# 少广章命令
@main.command()
@click.option("--file", "-f", required=True, help="Excel 文件路径（如 data.xlsx）")
@click.option("--sheet", "-s", required=True, help="工作表名称（如 Sheet1）")
@click.option("--cell", "-c", required=True, help="结果写入的单元格位置（如 A1）")
@click.option("--area", required=True, type=float, help="已知面积值（如 100.5）")
@click.option(
    "--shape",
    default="rectangle",
    type=click.Choice(["rectangle", "square", "circle"]),
    help="形状类型",
)
def find_dimension(file: str, sheet: str, cell: str, area: float, shape: str):
    """[少广章] 已知面积反推边长

    使用场景：
    - 已知矩形面积求长和宽
    - 已知圆面积求半径
    - 几何计算中的反向推导

    示例：
        abacus find-dimension -f data.xlsx -s Calc -c A1 --area 100 --shape rectangle
        abacus find-dimension -f data.xlsx -s Calc -c B1 --area 78.5 --shape circle
    """
    cap = registry.get("find_dimension")
    result = cap.execute(None, file=file, sheet=sheet, cell=cell, area=area, shape=shape)
    click.echo(json.dumps(result, ensure_ascii=False, indent=2))


@main.command()
@click.option("--file", "-f", required=True, help="Excel 文件路径（如 data.xlsx）")
@click.option("--sheet", "-s", required=True, help="工作表名称（如 Sheet1）")
@click.option("--cell", "-c", required=True, help="结果写入的单元格位置（如 A1）")
@click.option("--target-value", required=True, type=float, help="目标值（如 500）")
@click.option("--formula", required=True, help="公式模板，使用 {x} 作为未知数（如 {x}*2+10）")
def derive(file: str, sheet: str, cell: str, target_value: float, formula: str):
    """[少广章] 已知结果反推参数

    使用场景：
    - 已知目标利润反推需要的销量
    - 已知目标值反推公式中的未知参数
    - 逆向工程和敏感性分析

    示例：
        abacus derive -f data.xlsx -s Calc -c A1 --target-value 1000 --formula "{x}*1.1-50"
        abacus derive -f data.xlsx -s Model -c B5 --target-value 50000 --formula "{x}*{x}+100"
    """
    cap = registry.get("derive")
    result = cap.execute(
        None, file=file, sheet=sheet, cell=cell, target_value=target_value, formula=formula
    )
    click.echo(json.dumps(result, ensure_ascii=False, indent=2))


# 商功章命令
@main.command()
@click.option("--file", "-f", required=True, help="Excel 文件路径（如 data.xlsx）")
@click.option("--sheet", "-s", required=True, help="工作表名称（如 Sheet1）")
@click.option("--range", "-r", required=True, help="数据范围，Excel 格式（如 A1:D10）")
@click.option("--font", help='字体设置，JSON 格式（如 \'{"name":"Arial","size":12,"bold":true}\'）')
@click.option("--fill", help='填充设置，JSON 格式（如 \'{"color":"FF0000","pattern":"solid"}\'）')
@click.option("--border", help='边框设置，JSON 格式（如 \'{"style":"thin","color":"000000"}\'）')
@click.option(
    "--alignment", help='对齐设置，JSON 格式（如 \'{"horizontal":"center","vertical":"center"}\'）'
)
@click.option("--number-format", help="数字格式（如 '#,##0.00'、'yyyy-mm-dd'）")
@click.option("--conditional", help="条件格式设置，JSON 格式")
def format(
    file: str,
    sheet: str,
    range: str,
    font: str,
    fill: str,
    border: str,
    alignment: str,
    number_format: str,
    conditional: str,
):
    """[商功章] 格式化单元格样式

    使用场景：
    - 设置表头字体、颜色、对齐方式
    - 为数据区域添加边框和填充色
    - 设置数字格式（货币、百分比、日期等）
    - 添加条件格式（数据条、色阶等）

    示例：
        abacus format -f data.xlsx -s Sheet1 -r A1:D1 --font '{"bold":true,"size":14}' --fill '{"color":"4472C4"}'
        abacus format -f data.xlsx -s Sheet1 -r B2:B100 --number-format '#,##0.00'
    """
    import json as json_mod

    font_dict = json_mod.loads(font) if font else None
    fill_dict = json_mod.loads(fill) if fill else None
    border_dict = json_mod.loads(border) if border else None
    alignment_dict = json_mod.loads(alignment) if alignment else None
    conditional_dict = json_mod.loads(conditional) if conditional else None
    cap = registry.get("format_range")
    result = cap.execute(
        None,
        file=file,
        sheet=sheet,
        range=range,
        font=font_dict,
        fill=fill_dict,
        border=border_dict,
        alignment=alignment_dict,
        number_format=number_format,
        conditional=conditional_dict,
    )
    click.echo(json.dumps(result, ensure_ascii=False, indent=2))


@main.command()
@click.option("--file", "-f", required=True, help="Excel 文件路径（如 data.xlsx）")
@click.option(
    "--operations",
    "-o",
    required=True,
    help='操作列表，JSON 格式（如 \'[{"action":"write","sheet":"S1","cell":"A1","value":100}]\'）',
)
def batch(file: str, operations: str):
    """[商功章] 批量执行多个操作

    使用场景：
    - 一次性写入多个单元格
    - 批量修改多个工作表
    - 自动化重复性操作

    示例：
        abacus batch -f data.xlsx -o '[{"action":"write","sheet":"S1","cell":"A1","value":"Hello"}]'
    """
    import json as json_mod

    ops = json_mod.loads(operations)
    cap = registry.get("batch_execute")
    result = cap.execute(None, file=file, operations=ops)
    click.echo(json.dumps(result, ensure_ascii=False, indent=2))


@main.command()
@click.option("--file", "-f", required=True, help="Excel 文件路径（如 data.xlsx）")
@click.option(
    "--operations",
    "-o",
    required=True,
    help='转换操作列表，JSON 格式（如 \'[{"sheet":"S1","range":"A1:A10","type":"int"}]\'）',
)
def batch_transform(file: str, operations: str):
    """[商功章] 批量转换数据

    使用场景：
    - 批量转换多个范围的数据类型
    - 批量转换单位或格式
    - 数据清洗时的批量处理

    示例：
        abacus batch-transform -f data.xlsx -o '[{"sheet":"S1","range":"A1:A100","type":"float"}]'
    """
    import json as json_mod

    ops = json_mod.loads(operations)
    cap = registry.get("batch_transform")
    result = cap.execute(None, file=file, operations=ops)
    click.echo(json.dumps(result, ensure_ascii=False, indent=2))


@main.command()
@click.option("--file", "-f", required=True, help="Excel 文件路径（如 data.xlsx）")
@click.option(
    "--operations",
    "-o",
    required=True,
    help='验证操作列表，JSON 格式（如 \'[{"sheet":"S1","range":"A1:A10","type":"int"}]\'）',
)
def batch_validate(file: str, operations: str):
    """[商功章] 批量验证数据

    使用场景：
    - 批量验证多个范围的数据类型
    - 批量检查数据范围约束
    - 数据质量批量检查

    示例：
        abacus batch-validate -f data.xlsx -o '[{"sheet":"S1","range":"A1:A100","type":"int"}]'
    """
    import json as json_mod

    ops = json_mod.loads(operations)
    cap = registry.get("batch_validate")
    result = cap.execute(None, file=file, operations=ops)
    click.echo(json.dumps(result, ensure_ascii=False, indent=2))


@main.command()
@click.option("--file", "-f", required=True, help="Excel 文件路径（如 data.xlsx）")
@click.option("--sheet", "-s", required=True, help="源数据工作表名称（如 Sheet1）")
@click.option("--range", "-r", required=True, help="源数据范围，Excel 格式（如 A1:E100）")
@click.option(
    "--row-fields", required=True, help='行字段列表，JSON 格式（如 \'["Category","Region"]\'）'
)
@click.option("--value-field", required=True, help="值字段名（如 Sales、Quantity）")
@click.option(
    "--agg-function",
    default="sum",
    type=click.Choice(["sum", "avg", "count", "min", "max"]),
    help="聚合函数",
)
@click.option("--output-sheet", help="输出工作表名称（可选，默认自动创建）")
def create_pivot(
    file: str,
    sheet: str,
    range: str,
    row_fields: str,
    value_field: str,
    agg_function: str,
    output_sheet: str,
):
    """[商功章] 创建数据透视表

    使用场景：
    - 按多维度汇总分析数据
    - 生成交叉报表
    - 数据探索和多维分析

    示例：
        abacus create-pivot -f data.xlsx -s Sales -r A1:E1000 --row-fields '["Category","Region"]' --value-field Sales --agg-function sum
    """
    import json as json_mod

    row_fields_list = json_mod.loads(row_fields)
    cap = registry.get("create_pivot")
    result = cap.execute(
        None,
        file=file,
        sheet=sheet,
        range=range,
        row_fields=row_fields_list,
        value_field=value_field,
        agg_function=agg_function,
        output_sheet=output_sheet,
    )
    click.echo(json.dumps(result, ensure_ascii=False, indent=2))


@main.command("advanced-filter")
@click.option("--file", "-f", required=True, help="Excel 文件路径（如 data.xlsx）")
@click.option("--sheet", "-s", required=True, help="工作表名称（如 Sheet1）")
@click.option("--range", "-r", help="数据范围（如 A1:D100，可选）")
@click.option("--conditions", "-c", required=True, help="筛选条件，JSON 格式")
@click.option(
    "--return-type",
    default="data",
    type=click.Choice(["data", "rows"]),
    help="返回类型：data=数据，rows=行号",
)
def advanced_filter(file, sheet, range, conditions, return_type):
    """[商功章] 高级筛选（支持复杂条件的数据筛选）

    使用场景：
    - 需要执行复杂的多条件筛选
    - 支持 AND/OR/NOT 逻辑组合
    - 数值范围、文本匹配、日期范围筛选

    示例：
        abacus advanced-filter -f data.xlsx -s Sheet1 -c '{"type":"condition","field":"Sales",">",1000}'
        abacus advanced-filter -f data.xlsx -s Sheet1 -c '{"type":"group","logic":"AND","conditions":[{"type":"condition","field":"Sales",">",1000},{"type":"condition","field":"Region","==","East"}]}'
    """
    import json as json_mod

    conditions_dict = json_mod.loads(conditions)
    cap = registry.get("advanced_filter")
    result = cap.execute(
        None,
        file=file,
        sheet=sheet,
        range=range,
        conditions=conditions_dict,
        return_type=return_type,
    )
    click.echo(json.dumps(result, ensure_ascii=False, indent=2))


@main.command("create-mapping-template")
@click.option("--output", "-o", help="输出文件路径（可选，默认带时间戳）")
@click.option("--source-count", default=4, type=int, help="源表数量（默认 4）")
@click.option("--quiet/--no-quiet", default=False, help="静默模式")
def create_mapping_template(output, source_count, quiet):
    """[商功章] 创建数据映射模板

    使用场景：
    - 创建数据仓库的映射模板
    - 定义目标表与源表的映射关系
    - 标准化数据开发流程

    示例：
        abacus create-mapping-template
        abacus create-mapping-template -o my_template.xlsx --source-count 3
    """
    cap = registry.get("create_mapping_template")
    result = cap.execute(None, output=output, source_count=source_count, quiet=quiet)
    click.echo(json.dumps(result, ensure_ascii=False, indent=2))


# 均输章命令
@main.command()
@click.option("--file", "-f", required=True, help="目标 Excel 文件路径（如 output.xlsx）")
@click.option("--source", required=True, help="源文件路径（如 data.csv、data.json）")
@click.option("--source-type", default="csv", type=click.Choice(["csv", "json"]), help="源文件类型")
@click.option("--sheet", "-s", default="Sheet1", help="目标工作表名称（默认 Sheet1）")
def import_data(file: str, source: str, source_type: str, sheet: str):
    """[均输章] 导入外部数据到 Excel

    使用场景：
    - 将 CSV 文件导入 Excel 工作表
    - 将 JSON 数据导入 Excel
    - 数据迁移和整合

    示例：
        abacus import -f output.xlsx --source data.csv --source-type csv --sheet Sales
        abacus import -f output.xlsx --source data.json --source-type json --sheet Config
    """
    cap = registry.get("import_data")
    result = cap.execute(None, file=file, source=source, source_type=source_type, sheet=sheet)
    click.echo(json.dumps(result, ensure_ascii=False, indent=2))


@main.command()
@click.option("--file", "-f", required=True, help="源 Excel 文件路径（如 data.xlsx）")
@click.option("--sheet", "-s", required=True, help="工作表名称（如 Sheet1）")
@click.option("--range", "-r", required=True, help="数据范围，Excel 格式（如 A1:D100）")
@click.option("--target", required=True, help="目标文件路径（如 output.csv、output.json）")
@click.option(
    "--target-type", default="csv", type=click.Choice(["csv", "json"]), help="目标文件类型"
)
def export_data(file: str, sheet: str, range: str, target: str, target_type: str):
    """[均输章] 导出 Excel 数据到外部格式

    使用场景：
    - 将 Excel 数据导出为 CSV 供其他系统使用
    - 将 Excel 数据导出为 JSON 供程序处理
    - 数据交换和备份

    示例：
        abacus export -f data.xlsx -s Sales -r A1:E100 --target output.csv --target-type csv
        abacus export -f data.xlsx -s Config -r A1:B50 --target config.json --target-type json
    """
    cap = registry.get("export_data")
    result = cap.execute(
        None, file=file, sheet=sheet, range=range, target=target, target_type=target_type
    )
    click.echo(json.dumps(result, ensure_ascii=False, indent=2))


@main.command()
@click.option("--source", required=True, help="源 Excel 文件路径（如 old.xlsx）")
@click.option("--target", required=True, help="目标 Excel 文件路径（如 new.xlsx）")
@click.option(
    "--sheets", help='指定迁移的工作表列表，JSON 格式（如 \'["Sheet1","Sheet2"]\'，默认全部）'
)
def migrate(source: str, target: str, sheets: str):
    """[均输章] 数据迁移（跨文件）

    使用场景：
    - 将数据从一个 Excel 文件迁移到另一个
    - 合并多个 Excel 文件的工作表
    - 数据备份和归档

    示例：
        abacus migrate --source old.xlsx --target new.xlsx
        abacus migrate --source old.xlsx --target new.xlsx --sheets '["Sales","Orders"]'
    """
    import json as json_mod

    sheets_list = json_mod.loads(sheets) if sheets else None
    cap = registry.get("migrate")
    result = cap.execute(None, source=source, target=target, sheets=sheets_list)
    click.echo(json.dumps(result, ensure_ascii=False, indent=2))


# 盈不足章命令
@main.command()
@click.option("--file", "-f", required=True, help="Excel 文件路径（如 data.xlsx）")
@click.option("--sheet", "-s", required=True, help="工作表名称（如 Sheet1）")
@click.option("--range", "-r", required=True, help="数据范围，Excel 格式（如 A1:A100）")
@click.option("--min-value", type=float, help="最小值约束（如 0）")
@click.option("--max-value", type=float, help="最大值约束（如 1000）")
def validate_range(file: str, sheet: str, range: str, min_value: float, max_value: float):
    """[盈不足章] 验证数据是否在指定范围内

    使用场景：
    - 检查数值是否在合理范围内
    - 数据质量检查（如年龄不能为负数）
    - 业务规则验证（如价格不能超过上限）

    示例：
        abacus validate-range -f data.xlsx -s Sheet1 -r A1:A100 --min-value 0 --max-value 100
        abacus validate-range -f data.xlsx -s Sales -r D1:D500 --min-value 0
    """
    cap = registry.get("validate_range")
    result = cap.execute(
        None, file=file, sheet=sheet, range=range, min_value=min_value, max_value=max_value
    )
    click.echo(json.dumps(result, ensure_ascii=False, indent=2))


@main.command("set-data-validation")
@click.option("--file", "-f", required=True, help="Excel 文件路径（如 data.xlsx）")
@click.option("--sheet", "-s", required=True, help="工作表名称（如 Sheet1）")
@click.option("--range", "-r", required=True, help="数据范围，Excel 格式（如 A1:A10）")
@click.option(
    "--validation-type",
    required=True,
    type=click.Choice(["list", "number", "date", "text_length"]),
    help="验证类型：list=下拉列表，number=数值范围，date=日期范围，text_length=文本长度",
)
@click.option("--operator", help="运算符（如 between、notBetween、equal、notEqual 等）")
@click.option("--formula1", help="验证公式1（list 类型时为逗号分隔的选项，如 '是,否'）")
@click.option("--formula2", help="验证公式2（between 时需要）")
@click.option("--error-message", help="错误提示消息")
def set_data_validation(
    file: str,
    sheet: str,
    range: str,
    validation_type: str,
    operator: str,
    formula1: str,
    formula2: str,
    error_message: str,
):
    """[盈不足章] 设置单元格数据验证规则

    使用场景：
    - 创建下拉列表选择
    - 限制数值输入范围
    - 限制日期输入范围
    - 限制文本长度

    示例：
        abacus set-data-validation -f data.xlsx -s Sheet1 -r A1:A10 --validation-type list --formula1 '是,否'
        abacus set-data-validation -f data.xlsx -s Sheet1 -r B1:B10 --validation-type number --operator between --formula1 0 --formula2 100
        abacus set-data-validation -f data.xlsx -s Sheet1 -r C1:C10 --validation-type date --operator between --formula1 2020-01-01 --formula2 2030-12-31
    """
    cap = registry.get("set_data_validation")
    result = cap.execute(
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
    click.echo(json.dumps(result, ensure_ascii=False, indent=2))


@main.command()
@click.option("--file", "-f", required=True, help="Excel 文件路径（如 data.xlsx）")
@click.option("--sheet", "-s", required=True, help="工作表名称（如 Sheet1）")
@click.option("--range", "-r", required=True, help="数据范围，Excel 格式（如 A1:D100）")
@click.option(
    "--expected-type",
    required=True,
    type=click.Choice(["int", "float", "str", "date"]),
    help="期望的数据类型",
)
def validate_type(file: str, sheet: str, range: str, expected_type: str):
    """[盈不足章] 验证数据类型是否符合预期

    使用场景：
    - 检查列数据是否全为数字
    - 验证日期列格式是否正确
    - 数据导入前的类型检查

    示例：
        abacus validate-type -f data.xlsx -s Sheet1 -r A1:A100 --expected-type int
        abacus validate-type -f data.xlsx -s Orders -r C1:C500 --expected-type date
    """
    cap = registry.get("validate_type")
    result = cap.execute(None, file=file, sheet=sheet, range=range, expected_type=expected_type)
    click.echo(json.dumps(result, ensure_ascii=False, indent=2))


@main.command()
@click.option("--file", "-f", required=True, help="Excel 文件路径（如 data.xlsx）")
@click.option("--sheet", "-s", required=True, help="工作表名称（如 Sheet1）")
@click.option("--cell", "-c", required=True, help="包含公式的单元格位置（如 A1）")
def validate_formula(file: str, sheet: str, cell: str):
    """[盈不足章] 验证单元格公式的正确性

    使用场景：
    - 检查公式是否存在语法错误
    - 验证公式引用的单元格是否有效
    - 调试公式错误

    示例：
        abacus validate-formula -f data.xlsx -s Sheet1 -c A1
        abacus validate-formula -f data.xlsx -s Calc -c B10
    """
    cap = registry.get("validate_formula")
    result = cap.execute(None, file=file, sheet=sheet, cell=cell)
    click.echo(json.dumps(result, ensure_ascii=False, indent=2))


@main.command()
@click.option("--file", "-f", required=True, help="Excel 文件路径（如 data.xlsx）")
def validate_file(file: str):
    """[盈不足章] 验证 Excel 文件结构（ZIP 格式、XML 结构、公式错误）

    使用场景：
    - 检查 Excel 文件质量
    - 发现文件中的潜在问题
    - 文件健康检查

    示例：
        abacus validate-file -f data.xlsx
    """
    cap = registry.get("validate_file")
    result = cap.execute(None, file=file)
    click.echo(json.dumps(result, ensure_ascii=False, indent=2))


@main.command("validation-engine")
@click.option("--file", "-f", required=True, help="Excel 文件路径（如 data.xlsx）")
@click.option("--sheet", "-s", required=True, help="工作表名称（如 Sheet1）")
@click.option("--range", "-r", required=True, help="数据范围，Excel 格式（如 A1:D100）")
@click.option(
    "--rules",
    "-o",
    required=True,
    help='验证规则列表，JSON 格式（如 \'[{"type":"type","params":{"expected_type":"int"}}]\'）',
)
def validation_engine(file: str, sheet: str, range: str, rules: str):
    """[盈不足章] 数据验证规则引擎（自定义规则、AND/OR 组合、规则链）

    使用场景：
    - 执行复杂的多规则验证
    - 支持 AND/OR 逻辑组合验证
    - 按顺序执行规则链
    - 自定义 Python 表达式验证

    示例：
        abacus validation-engine -f data.xlsx -s Sheet1 -r A1:A10 -o '[{"type":"not_empty"},{"type":"type","params":{"expected_type":"int"}}]'
        abacus validation-engine -f data.xlsx -s Sheet1 -r B1:B10 -o '[{"type":"range","params":{"min_val":0,"max_val":100}}]'
        abacus validation-engine -f data.xlsx -s Sheet1 -r C1:C10 -o '[{"type":"and","rules":[{"type":"not_empty"},{"type":"regex","params":{"pattern":"^[a-z]+$"}}]}]'
    """
    import json as json_mod

    rules_list = json_mod.loads(rules)
    cap = registry.get("validation_engine")
    result = cap.execute(None, file=file, sheet=sheet, range=range, rules=rules_list)
    click.echo(json.dumps(result, ensure_ascii=False, indent=2))


# 方程章命令
@main.command()
@click.option("--file", "-f", required=True, help="Excel 文件路径（如 data.xlsx）")
@click.option("--sheet", "-s", help="工作表名称（可选，默认检查所有工作表）")
@click.option("--cell", "-c", help="单元格位置（可选，默认检查所有公式单元格）")
def diagnose_formula(file: str, sheet: str, cell: str):
    """[方程章] 诊断公式错误

    使用场景：
    - 检查 Excel 文件中的公式错误
    - 诊断公式语法问题
    - 查找包含错误的单元格
    - 公式调试和修复

    示例：
        abacus diagnose-formula -f data.xlsx
        abacus diagnose-formula -f data.xlsx -s Sheet1
        abacus diagnose-formula -f data.xlsx -s Sheet1 -c A1
    """
    cap = registry.get("diagnose_formula")
    result = cap.execute(None, file=file, sheet=sheet, cell=cell)
    click.echo(json.dumps(result, ensure_ascii=False, indent=2))


@main.command()
@click.option("--file", "-f", required=True, help="Excel 文件路径（如 data.xlsx）")
@click.option("--sheet", "-s", required=True, help="工作表名称（如 Sheet1）")
@click.option("--cell", "-c", required=True, help="公式写入的单元格位置（如 A1）")
@click.option(
    "--formula", required=True, help="公式内容（如 =SUM(A1:A10)、=VLOOKUP(B1,Sheet2!A:C,3,0)）"
)
def formula(file: str, sheet: str, cell: str, formula: str):
    """[方程章] 在单元格中创建公式

    使用场景：
    - 为单元格添加计算公式
    - 创建引用其他单元格的公式
    - 批量设置公式

    示例：
        abacus formula -f data.xlsx -s Sheet1 -c D1 --formula '=SUM(A1:C1)'
        abacus formula -f data.xlsx -s Sheet1 -c E1 --formula '=VLOOKUP(D1,Sheet2!A:B,2,0)'
    """
    cap = registry.get("create_formula")
    result = cap.execute(None, file=file, sheet=sheet, cell=cell, formula=formula)
    click.echo(json.dumps(result, ensure_ascii=False, indent=2))


@main.command()
@click.option("--file", "-f", required=True, help="Excel 文件路径（如 data.xlsx）")
@click.option("--sheet", "-s", required=True, help="工作表名称（如 Sheet1）")
@click.option("--cell", "-c", required=True, help="结果写入的单元格位置（如 A1）")
@click.option("--equation", required=True, help="方程表达式（如 2*x+10=100、x^2-4=0）")
def solve_equation(file: str, sheet: str, cell: str, equation: str):
    """[方程章] 解方程并将结果写入单元格

    使用场景：
    - 求解一元一次方程
    - 求解一元二次方程
    - 数学建模中的方程求解

    示例：
        abacus solve-equation -f data.xlsx -s Calc -c A1 --equation '2*x+10=100'
        abacus solve-equation -f data.xlsx -s Calc -c B1 --equation 'x^2-4=0'
    """
    cap = registry.get("solve_equation")
    result = cap.execute(None, file=file, sheet=sheet, cell=cell, equation=equation)
    click.echo(json.dumps(result, ensure_ascii=False, indent=2))


@main.command()
@click.option("--file", "-f", required=True, help="Excel 文件路径（如 data.xlsx）")
@click.option("--sheet", "-s", required=True, help="工作表名称（如 Sheet1）")
@click.option("--cell", "-c", required=True, help="结果写入的单元格位置（如 A1）")
@click.option("--expression", required=True, help="计算表达式（如 100*1.1+50、sqrt(144)）")
def calculate(file: str, sheet: str, cell: str, expression: str):
    """[方程章] 执行数学计算并将结果写入单元格

    使用场景：
    - 执行复杂的数学运算
    - 计算结果直接写入 Excel
    - 自动化计算流程

    示例：
        abacus calculate -f data.xlsx -s Calc -c A1 --expression '100*1.1+50'
        abacus calculate -f data.xlsx -s Calc -c B1 --expression 'sqrt(144)+2^3'
    """
    cap = registry.get("calculate")
    result = cap.execute(None, file=file, sheet=sheet, cell=cell, expression=expression)
    click.echo(json.dumps(result, ensure_ascii=False, indent=2))


@main.command()
@click.option("--file", "-f", required=True, help="Excel 文件路径（如 data.xlsx）")
@click.option("--output", "-o", help="输出文件路径（可选）")
def recalc_formulas(file: str, output: str):
    """[方程章] 公式重算：使用 LibreOffice 重算 Excel 公式（扫描所有错误）

    使用场景：
    - 需要重算 Excel 文件中的公式
    - 检查公式错误
    - 批量公式验证

    示例：
        abacus recalc-formulas -f data.xlsx
        abacus recalc-formulas -f data.xlsx -o recalc.xlsx
    """
    cap = registry.get("recalc_formulas")
    result = cap.execute(None, file=file, output=output)
    click.echo(json.dumps(result, ensure_ascii=False, indent=2))


# 勾股章命令
@main.command()
@click.option("--file", "-f", required=True, help="Excel 文件路径（如 data.xlsx）")
@click.option("--sheet", "-s", required=True, help="工作表名称（如 Sheet1）")
@click.option("--range", "-r", required=True, help="数据范围，Excel 格式（如 A1:D100）")
def analyze_stats(file: str, sheet: str, range: str):
    """[勾股章] 统计分析（描述性统计）

    使用场景：
    - 计算数据的均值、中位数、标准差等
    - 了解数据分布特征
    - 数据探索和质量评估

    示例：
        abacus analyze-stats -f data.xlsx -s Sales -r A1:E100
        abacus analyze-stats -f data.xlsx -s Survey -r B1:B500
    """
    cap = registry.get("analyze_stats")
    result = cap.execute(None, file=file, sheet=sheet, range=range)
    click.echo(json.dumps(result, ensure_ascii=False, indent=2))


@main.command()
@click.option("--file", "-f", required=True, help="Excel 文件路径（如 data.xlsx）")
@click.option("--sheet", "-s", required=True, help="工作表名称（如 Sheet1）")
@click.option("--range", "-r", required=True, help="数据范围，Excel 格式（如 A1:D100）")
@click.option("--time-column", required=True, help="时间列名（如 Date、Month）")
def analyze_trend(file: str, sheet: str, range: str, time_column: str):
    """[勾股章] 时间序列趋势分析

    使用场景：
    - 分析销售数据的月度/年度趋势
    - 识别数据的增长/下降趋势
    - 预测未来走势

    示例：
        abacus analyze-trend -f data.xlsx -s Sales -r A1:D100 --time-column Month
        abacus analyze-trend -f data.xlsx -s Revenue -r A1:C50 --time-column Year
    """
    cap = registry.get("analyze_trend")
    result = cap.execute(None, file=file, sheet=sheet, range=range, time_column=time_column)
    click.echo(json.dumps(result, ensure_ascii=False, indent=2))


@main.command()
@click.option("--file", "-f", required=True, help="Excel 文件路径（如 data.xlsx）")
@click.option("--sheet", "-s", required=True, help="工作表名称（如 Sheet1）")
@click.option("--range", "-r", required=True, help="数据范围，Excel 格式（如 A1:D100）")
@click.option("--column1", required=True, help="第一列名（如 Price）")
@click.option("--column2", required=True, help="第二列名（如 Sales）")
def analyze_correlation(file: str, sheet: str, range: str, column1: str, column2: str):
    """[勾股章] 两列数据的相关性分析

    使用场景：
    - 分析价格与销量的关系
    - 探索变量之间的线性相关性
    - 因果关系初步判断

    示例：
        abacus analyze-correlation -f data.xlsx -s Sales -r A1:C100 --column1 Price --column2 Quantity
        abacus analyze-correlation -f data.xlsx -s Survey -r A1:D200 --column1 Age --column2 Income
    """
    cap = registry.get("analyze_correlation")
    result = cap.execute(
        None, file=file, sheet=sheet, range=range, column1=column1, column2=column2
    )
    click.echo(json.dumps(result, ensure_ascii=False, indent=2))


@main.command()
@click.option("--file", "-f", required=True, help="Excel 文件路径（如 data.xlsx）")
@click.option("--sheet", "-s", required=True, help="数据工作表名称（如 Sheet1）")
@click.option("--range", "-r", required=True, help="数据范围，Excel 格式（如 A1:C10）")
@click.option(
    "--chart-type",
    required=True,
    type=click.Choice(["bar", "line", "pie", "area", "scatter"]),
    help="图表类型",
)
@click.option("--title", help="图表标题（如 '月度销售趋势'）")
@click.option("--x-axis", help="X轴标题（如 '月份'）")
@click.option("--y-axis", help="Y轴标题（如 '销售额'）")
@click.option("--output-sheet", help="输出工作表名称（可选，默认在数据工作表中创建）")
@click.option("--position", default="A1", help="图表左上角位置（默认 A1）")
@click.option("--width", default=15, type=float, help="图表宽度，单位厘米（默认 15）")
@click.option("--height", default=10, type=float, help="图表高度，单位厘米（默认 10）")
def create_chart(
    file, sheet, range, chart_type, title, x_axis, y_axis, output_sheet, position, width, height
):
    """[商功章] 创建 Excel 图表

    使用场景：
    - 将数据可视化为柱状图、折线图、饼图等
    - 创建数据仪表板
    - 生成报表图表

    示例：
        abacus create-chart -f data.xlsx -s Sales -r A1:C10 --chart-type bar --title '月度销售'
        abacus create-chart -f data.xlsx -s Data -r A1:B20 --chart-type line --title '趋势图' --x-axis '时间' --y-axis '数值'
    """
    cap = registry.get("create_chart")
    result = cap.execute(
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
    click.echo(json.dumps(result, ensure_ascii=False, indent=2))


@main.command()
@click.option("--file", "-f", required=True, help="Excel 文件路径（如 data.xlsx）")
def list_charts(file):
    """[商功章] 列出 Excel 文件中的所有图表

    使用场景：
    - 查看文件中有哪些图表
    - 获取图表的索引以便删除或更新
    - 图表管理和审计

    示例：
        abacus list-charts -f data.xlsx
    """
    cap = registry.get("list_charts")
    result = cap.execute(None, file=file)
    click.echo(json.dumps(result, ensure_ascii=False, indent=2))


@main.command()
@click.option("--file", "-f", required=True, help="Excel 文件路径（如 data.xlsx）")
@click.option("--sheet", "-s", required=True, help="工作表名称（如 Sheet1）")
@click.option(
    "--chart-index", required=True, type=int, help="图表索引（从 0 开始，可通过 list-charts 获取）"
)
def delete_chart(file, sheet, chart_index):
    """[商功章] 删除指定图表

    使用场景：
    - 删除不需要的图表
    - 清理报表中的旧图表
    - 图表管理

    示例：
        abacus delete-chart -f data.xlsx -s Sheet1 --chart-index 0
        abacus delete-chart -f data.xlsx -s Dashboard --chart-index 2
    """
    cap = registry.get("delete_chart")
    result = cap.execute(None, file=file, sheet=sheet, chart_index=chart_index)
    click.echo(json.dumps(result, ensure_ascii=False, indent=2))


@main.command("create-advanced-chart")
@click.option("--file", "-f", required=True, help="输出文件路径（如 chart.xlsx）")
@click.option("--data", "-d", required=True, help="图表数据，JSON 格式（包含 headers 和 rows）")
@click.option(
    "--chart-type",
    required=True,
    type=click.Choice(["combo", "dual_axis", "waterfall", "gantt"]),
    help="图表类型：combo=组合图，dual_axis=双轴图，waterfall=瀑布图，gantt=甘特图",
)
@click.option("--title", help="图表标题（如 '销售趋势'）")
@click.option("--x-axis", help="X轴标题（如 '月份'）")
@click.option("--y-axis", help="Y轴标题（如 '金额'）")
def create_advanced_chart(file, data, chart_type, title, x_axis, y_axis):
    """[商功章] 创建高级图表（组合图、双轴图、瀑布图、甘特图）

    使用场景：
    - 创建组合图（柱形图+折线图）展示多维度数据
    - 创建双轴图（左右Y轴）对比不同量纲的数据
    - 创建瀑布图展示增减变化过程
    - 创建甘特图展示项目进度

    示例：
        abacus create-advanced-chart -f chart.xlsx -d '{"headers": list["月份","销售额","利润"],"rows": list[["1月",100,20],["2月",120,25]]}' --chart-type combo --title '销售趋势'
        abacus create-advanced-chart -f gantt.xlsx -d '{"headers": list["任务","开始","持续"],"rows": list[["设计",0,5],["开发",5,10]]}' --chart-type gantt --title '项目进度'
    """
    import json as json_mod

    data_dict = json_mod.loads(data)
    cap = registry.get("create_advanced_chart")
    result = cap.execute(
        None,
        file=file,
        data=data_dict,
        chart_type=chart_type,
        title=title,
        x_axis=x_axis,
        y_axis=y_axis,
    )
    click.echo(json.dumps(result, ensure_ascii=False, indent=2))


@main.command()
@click.option("--file", "-f", required=True, help="数据文件路径（如 data.xlsx、data.csv）")
@click.option("--output", "-o", required=True, help="输出图片路径（如 chart.png、chart.svg）")
@click.option(
    "--chart-type",
    required=True,
    type=click.Choice(["bar", "line", "pie", "scatter", "heatmap"]),
    help="图表类型：bar=柱状图，line=折线图，pie=饼图，scatter=散点图，heatmap=热力图",
)
@click.option("--x-column", help="X轴列名（可选）")
@click.option("--y-column", help="Y轴列名（可选）")
@click.option("--sheet", "-s", help="工作表名称（Excel 文件时可选）")
@click.option("--title", help="图表标题（可选）")
@click.option("--width", default=10, type=float, help="图片宽度，单位英寸（默认 10）")
@click.option("--height", default=6, type=float, help="图片高度，单位英寸（默认 6）")
def visualize(file, output, chart_type, x_column, y_column, sheet, title, width, height):
    """[勾股章] 数据可视化（生成 PNG/SVG/PDF 图表）

    使用场景：
    - 将 Excel/CSV 数据生成图片格式的图表
    - 创建柱状图、折线图、饼图、散点图、热力图
    - 数据报告和演示文稿中的图表

    示例：
        abacus visualize -f data.xlsx -o chart.png --chart-type bar --title '月度销售'
        abacus visualize -f data.csv -o trend.png --chart-type line --x-column 月份 --y-column 销售额
        abacus visualize -f data.xlsx -s Sales -o pie.png --chart-type pie --title '产品占比'
    """
    cap = registry.get("visualize")
    result = cap.execute(
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
    click.echo(json.dumps(result, ensure_ascii=False, indent=2))


# 范围扩展命令
@main.command()
@click.option("--file", "-f", required=True, help="Excel 文件路径（如 data.xlsx）")
@click.option("--sheet", "-s", required=True, help="工作表名称（如 Sheet1）")
@click.option("--range", "-r", required=True, help="数据范围，Excel 格式（如 A1:D10）")
@click.option(
    "--clear-type",
    default="all",
    type=click.Choice(["all", "contents", "formats"]),
    help="清除类型：all=全部，contents=仅内容，formats=仅格式",
)
def clear_range(file, sheet, range, clear_type):
    """[范围扩展] 清除指定范围的内容或格式

    使用场景：
    - 清空数据区域的内容
    - 清除单元格格式（保留数据）
    - 重置工作表区域

    示例：
        abacus clear-range -f data.xlsx -s Sheet1 -r A1:D10 --clear-type contents
        abacus clear-range -f data.xlsx -s Sheet1 -r A1:Z100 --clear-type all
    """
    cap = registry.get("clear_range")
    result = cap.execute(None, file=file, sheet=sheet, range=range, clear_type=clear_type)
    click.echo(json.dumps(result, ensure_ascii=False, indent=2))


@main.command("copy-range")
@click.option("--file", "-f", required=True, help="Excel 文件路径（如 data.xlsx）")
@click.option("--sheet", "-s", required=True, help="源工作表名称（如 Sheet1）")
@click.option("--source", required=True, help="源范围，Excel 格式（如 A1:D10）")
@click.option("--target", required=True, help="目标位置（如 F1 或其他工作表的 A1）")
@click.option(
    "--copy-type",
    default="all",
    type=click.Choice(["all", "values", "formulas", "formats"]),
    help="复制类型：all=全部，values=仅值，formulas=仅公式，formats=仅格式",
)
def copy_range(file, sheet, source, target, copy_type):
    """[范围扩展] 复制范围数据到目标位置

    使用场景：
    - 复制数据区域到其他位置
    - 仅复制值（去除公式）
    - 复制格式到其他区域

    示例：
        abacus copy-range -f data.xlsx -s Sheet1 --source A1:D10 --target F1 --copy-type values
        abacus copy-range -f data.xlsx -s Sheet1 --source A1:A100 --target Sheet2!A1 --copy-type all
    """
    cap = registry.get("copy_range")
    result = cap.execute(
        None, file=file, sheet=sheet, source=source, target=target, copy_type=copy_type
    )
    click.echo(json.dumps(result, ensure_ascii=False, indent=2))


@main.command()
@click.option("--file", "-f", required=True, help="Excel 文件路径（如 data.xlsx）")
@click.option("--sheet", "-s", required=True, help="工作表名称（如 Sheet1）")
@click.option("--find", required=True, help="查找内容（如 '旧文本'、'123'）")
@click.option("--replace", help="替换内容（如 '新文本'、'456'，不填则删除匹配内容）")
@click.option("--range", "-r", help="查找范围，Excel 格式（如 A1:D100，默认整个工作表）")
def find_replace(file, sheet, find, replace, range):
    """[范围扩展] 查找并替换文本

    使用场景：
    - 批量替换错误的文本
    - 更新旧数据为新值
    - 清理数据中的特殊字符

    示例：
        abacus find-replace -f data.xlsx -s Sheet1 --find '旧值' --replace '新值'
        abacus find-replace -f data.xlsx -s Sheet1 --find 'N/A' --replace '' -r A1:Z1000
    """
    cap = registry.get("find_replace")
    result = cap.execute(None, file=file, sheet=sheet, find=find, replace=replace, range=range)
    click.echo(json.dumps(result, ensure_ascii=False, indent=2))


@main.command()
@click.option("--file", "-f", required=True, help="Excel 文件路径（如 data.xlsx）")
@click.option("--sheet", "-s", required=True, help="工作表名称（如 Sheet1）")
@click.option(
    "--action", required=True, type=click.Choice(["add", "remove", "list"]), help="操作类型"
)
@click.option("--cell", help="单元格位置（如 A1，add/remove 时必填）")
@click.option("--url", help="链接地址（如 https://example.com，add 时必填）")
def hyperlink(file, sheet, action, cell, url):
    """[范围扩展] 管理单元格超链接

    使用场景：
    - 为单元格添加网页链接
    - 删除不需要的超链接
    - 列出工作表中的所有超链接

    示例：
        abacus hyperlink -f data.xlsx -s Sheet1 --action add --cell A1 --url 'https://example.com'
        abacus hyperlink -f data.xlsx -s Sheet1 --action list
        abacus hyperlink -f data.xlsx -s Sheet1 --action remove --cell A1
    """
    cap = registry.get("manage_hyperlink")
    result = cap.execute(None, file=file, sheet=sheet, action=action, cell=cell, url=url)
    click.echo(json.dumps(result, ensure_ascii=False, indent=2))


@main.command()
@click.option("--file", "-f", required=True, help="Excel 文件路径（如 data.xlsx）")
@click.option("--sheet", "-s", required=True, help="工作表名称（如 Sheet1）")
@click.option("--range", "-r", required=True, help="数据范围，Excel 格式（如 A1:D10）")
@click.option(
    "--locked/--unlocked", default=True, help="是否锁定单元格（--locked 锁定，--unlocked 解锁）"
)
def cell_lock(file, sheet, range, locked):
    """[范围扩展] 管理单元格锁定状态

    使用场景：
    - 保护公式单元格不被误修改
    - 锁定关键数据区域
    - 配合工作表保护使用

    示例：
        abacus cell-lock -f data.xlsx -s Sheet1 -r A1:D1 --locked
        abacus cell-lock -f data.xlsx -s Sheet1 -r E1:E100 --unlocked
    """
    cap = registry.get("manage_cell_lock")
    result = cap.execute(None, file=file, sheet=sheet, range=range, locked=locked)
    click.echo(json.dumps(result, ensure_ascii=False, indent=2))


@main.command()
@click.option("--file", "-f", required=True, help="Excel 文件路径（如 data.xlsx）")
@click.option("--sheet", "-s", required=True, help="工作表名称（如 Sheet1）")
@click.option(
    "--action",
    required=True,
    type=click.Choice(["set", "get", "auto"]),
    help="操作类型：set=设置，get=获取，auto=自动调整",
)
@click.option(
    "--dimension",
    required=True,
    type=click.Choice(["row", "column"]),
    help="维度：row=行高，column=列宽",
)
@click.option("--index", required=True, type=int, help="行号或列号（从 1 开始）")
@click.option("--size", type=float, help="大小（set 时必填，单位：行高为磅，列宽为字符数）")
def manage_size(file, sheet, action, dimension, index, size):
    """[范围扩展] 管理行高和列宽

    使用场景：
    - 调整特定行的高度
    - 调整特定列的宽度
    - 自动调整以适应内容

    示例：
        abacus manage-size -f data.xlsx -s Sheet1 --action set --dimension row --index 1 --size 20
        abacus manage-size -f data.xlsx -s Sheet1 --action auto --dimension column --index 1
        abacus manage-size -f data.xlsx -s Sheet1 --action get --dimension row --index 5
    """
    cap = registry.get("manage_size")
    result = cap.execute(
        None, file=file, sheet=sheet, action=action, dimension=dimension, index=index, size=size
    )
    click.echo(json.dumps(result, ensure_ascii=False, indent=2))


# 表格管理命令
@main.command()
@click.option("--file", "-f", required=True, help="Excel 文件路径（如 data.xlsx）")
@click.option("--sheet", "-s", required=True, help="工作表名称（如 Sheet1）")
@click.option(
    "--action",
    required=True,
    type=click.Choice(["add", "delete", "get", "list"]),
    help="操作类型：add=添加批注，delete=删除批注，get=获取批注，list=列出所有批注",
)
@click.option("--cell", help="单元格位置（如 A1，add/delete/get 时必填）")
@click.option("--text", help="批注内容（add 时必填）")
@click.option("--author", default="Abacus", help="批注作者（默认 Abacus）")
def comment(file, sheet, action, cell, text, author):
    """[商功章] 管理单元格批注

    使用场景：
    - 为单元格添加说明性批注
    - 删除不需要的批注
    - 查看批注内容
    - 列出工作表中的所有批注

    示例：
        abacus comment -f data.xlsx -s Sheet1 --action add --cell A1 --text '这是重要数据'
        abacus comment -f data.xlsx -s Sheet1 --action list
        abacus comment -f data.xlsx -s Sheet1 --action get --cell A1
        abacus comment -f data.xlsx -s Sheet1 --action delete --cell A1
    """
    cap = registry.get("manage_comment")
    result = cap.execute(
        None, file=file, sheet=sheet, action=action, cell=cell, text=text, author=author
    )
    click.echo(json.dumps(result, ensure_ascii=False, indent=2))


@main.command()
@click.option("--file", "-f", required=True, help="Excel 文件路径（如 data.xlsx）")
@click.option("--sheet", "-s", required=True, help="工作表名称（如 Sheet1）")
@click.option("--rows", type=int, help="冻结行数（如 1 表示冻结第一行）")
@click.option("--columns", type=int, help="冻结列数（如 1 表示冻结第一列）")
@click.option("--cell", help="冻结位置（如 B2 表示冻结第一行和第一列，与 --rows/--columns 二选一）")
def freeze(file, sheet, rows, columns, cell):
    """[商功章] 冻结窗格

    使用场景：
    - 冻结首行表头，滚动时始终可见
    - 冻结首列标签，横向滚动时始终可见
    - 同时冻结行和列
    - 解除已有的冻结设置

    示例：
        abacus freeze -f data.xlsx -s Sheet1 --rows 1
        abacus freeze -f data.xlsx -s Sheet1 --cell B2
        abacus freeze -f data.xlsx -s Sheet1 --rows 2 --columns 1
        abacus freeze -f data.xlsx -s Sheet1  # 不带参数表示解除冻结
    """
    cap = registry.get("freeze_panes")
    result = cap.execute(None, file=file, sheet=sheet, rows=rows, columns=columns, cell=cell)
    click.echo(json.dumps(result, ensure_ascii=False, indent=2))


@main.command("set-auto-filter")
@click.option("--file", "-f", required=True, help="Excel 文件路径（如 data.xlsx）")
@click.option("--sheet", "-s", required=True, help="工作表名称（如 Sheet1）")
@click.option(
    "--action",
    required=True,
    type=click.Choice(["set", "remove", "get"]),
    help="操作类型：set=设置筛选，remove=删除筛选，get=查询筛选",
)
@click.option("--range", "-r", help="筛选范围，Excel 格式（如 A1:D100，set 时必填）")
@click.option("--column", help="筛选列（可选）")
@click.option("--criteria", help="筛选条件（可选）")
def set_auto_filter(file, sheet, action, range, column, criteria):
    """[商功章] 设置和管理自动筛选

    使用场景：
    - 为数据区域添加自动筛选下拉箭头
    - 删除已有的自动筛选
    - 查询当前自动筛选状态和条件

    示例：
        abacus set-auto-filter -f data.xlsx -s Sheet1 --action set -r A1:D100
        abacus set-auto-filter -f data.xlsx -s Sheet1 --action get
        abacus set-auto-filter -f data.xlsx -s Sheet1 --action remove
    """
    cap = registry.get("set_auto_filter")
    result = cap.execute(
        None, file=file, sheet=sheet, action=action, range=range, column=column, criteria=criteria
    )
    click.echo(json.dumps(result, ensure_ascii=False, indent=2))


@main.command("manage-visibility")
@click.option("--file", "-f", required=True, help="Excel 文件路径（如 data.xlsx）")
@click.option("--sheet", "-s", required=True, help="工作表名称（如 Sheet1）")
@click.option(
    "--action",
    required=True,
    type=click.Choice(["hide", "show"]),
    help="操作类型：hide=隐藏，show=显示",
)
@click.option(
    "--dimension",
    required=True,
    type=click.Choice(["row", "column"]),
    help="维度：row=行，column=列",
)
@click.option("--index", required=True, type=int, help="行号或列号（从 1 开始）")
def manage_visibility(file, sheet, action, dimension, index):
    """[商功章] 隐藏或显示行和列

    使用场景：
    - 隐藏辅助数据行或列
    - 显示已隐藏的行或列
    - 数据展示时隐藏不需要的列

    示例：
        abacus manage-visibility -f data.xlsx -s Sheet1 --action hide --dimension row --index 3
        abacus manage-visibility -f data.xlsx -s Sheet1 --action show --dimension column --index 2
        abacus manage-visibility -f data.xlsx -s Sheet1 --action hide --dimension column --index 5
    """
    cap = registry.get("manage_visibility")
    result = cap.execute(
        None, file=file, sheet=sheet, action=action, dimension=dimension, index=index
    )
    click.echo(json.dumps(result, ensure_ascii=False, indent=2))


# 命名范围管理命令
@main.command()
@click.option("--file", "-f", required=True, help="Excel 文件路径（如 data.xlsx）")
@click.option(
    "--action",
    required=True,
    type=click.Choice(["create", "list", "read", "delete"]),
    help="操作类型",
)
@click.option("--name", help="命名范围名称（如 SalesData，create/read/delete 时必填）")
@click.option("--refers-to", help="引用位置（如 Sheet1!$A$1:$D$100，create 时必填）")
def manage_named_range(file, action, name, refers_to):
    """[方田章] 管理命名范围

    使用场景：
    - 创建有意义的范围名称便于公式引用
    - 列出文件中的所有命名范围
    - 读取或删除命名范围

    示例：
        abacus manage-named-range -f data.xlsx --action create --name SalesData --refers-to 'Sheet1!$A$1:$E$100'
        abacus manage-named-range -f data.xlsx --action list
        abacus manage-named-range -f data.xlsx --action delete --name SalesData
    """
    cap = registry.get("manage_named_range")
    result = cap.execute(None, file=file, action=action, name=name, refers_to=refers_to)
    click.echo(json.dumps(result, ensure_ascii=False, indent=2))


# 工作表扩展命令
@main.command()
@click.option("--file", "-f", required=True, help="Excel 文件路径（如 data.xlsx）")
@click.option("--sheet", "-s", required=True, help="工作表名称（如 Sheet1）")
@click.option(
    "--action",
    required=True,
    type=click.Choice(["set", "get", "clear"]),
    help="操作类型：set=设置颜色，get=获取颜色，clear=清除颜色",
)
@click.option("--color", help="颜色值，6位十六进制（如 FF0000=红色，00FF00=绿色，set 时必填）")
def sheet_style(file, sheet, action, color):
    """[工作表扩展] 管理工作表标签颜色

    使用场景：
    - 用颜色区分不同类型的工作表
    - 标记重要工作表
    - 视觉化工作表分类

    示例：
        abacus sheet-style -f data.xlsx -s Sales --action set --color FF0000
        abacus sheet-style -f data.xlsx -s Sales --action get
        abacus sheet-style -f data.xlsx -s Sales --action clear
    """
    cap = registry.get("manage_sheet_style")
    result = cap.execute(None, file=file, sheet=sheet, action=action, color=color)
    click.echo(json.dumps(result, ensure_ascii=False, indent=2))


@main.command()
@click.option("--file", "-f", required=True, help="Excel 文件路径（如 data.xlsx）")
@click.option("--sheet", "-s", required=True, help="工作表名称（如 Sheet1）")
@click.option(
    "--action",
    required=True,
    type=click.Choice(["show", "hide", "very-hide", "get"]),
    help="操作类型：show=显示，hide=隐藏，very-hide=深度隐藏，get=获取状态",
)
def sheet_visibility(file, sheet, action):
    """[工作表扩展] 管理工作表可见性

    使用场景：
    - 隐藏辅助工作表（如配置表、临时数据表）
    - 深度隐藏（无法通过 Excel 界面取消隐藏）
    - 检查工作表的可见状态

    示例：
        abacus sheet-visibility -f data.xlsx -s Config --action hide
        abacus sheet-visibility -f data.xlsx -s TempData --action very-hide
        abacus sheet-visibility -f data.xlsx -s Sales --action get
    """
    cap = registry.get("manage_sheet_visibility")
    result = cap.execute(None, file=file, sheet=sheet, action=action)
    click.echo(json.dumps(result, ensure_ascii=False, indent=2))


# 审计命令
@main.command()
@click.option("--code", help="Python 代码内容（直接传入代码字符串）")
@click.option("--file", "-f", help="Python 文件路径（从文件读取代码）")
def excel_lint(code, file):
    """[审计] 检查 openpyxl 代码的潜在问题

    使用场景：
    - 检查 openpyxl 代码中的性能问题
    - 发现潜在的兼容性问题
    - 代码审查和质量保证

    示例：
        abacus excel-lint --code 'from openpyxl import Workbook; wb = Workbook()'
        abacus excel-lint -f my_script.py
    """
    cap = registry.get("excel_lint")
    result = cap.execute(None, code=code, file=file)
    click.echo(json.dumps(result, ensure_ascii=False, indent=2))


@main.command()
@click.option("--file", "-f", required=True, help="Excel 文件路径（如 data.xlsx）")
def file_analyze(file):
    """[审计] 分析 Excel 文件的结构和潜在问题

    使用场景：
    - 检查 Excel 文件的健康状态
    - 发现文件中的潜在问题（如循环引用、损坏的对象）
    - 文件迁移前的兼容性检查

    示例：
        abacus analyze -f data.xlsx
        abacus analyze -f report.xlsx
    """
    cap = registry.get("file_analyze")
    result = cap.execute(None, file=file)
    click.echo(json.dumps(result, ensure_ascii=False, indent=2))


# 样式命令
@main.command()
@click.option("--file", "-f", required=True, help="Excel 文件路径（如 data.xlsx）")
@click.option("--sheet", "-s", required=True, help="工作表名称（如 Sheet1）")
@click.option(
    "--action",
    required=True,
    type=click.Choice(["apply_header", "apply_kpi", "auto_width"]),
    help="操作类型：apply_header=应用表头样式，apply_kpi=应用 KPI 样式，auto_width=自动列宽",
)
@click.option(
    "--range", "-r", help="数据范围，Excel 格式（如 A1:E1，apply_header/apply_kpi 时必填）"
)
@click.option(
    "--industry",
    default="finance",
    type=click.Choice(["finance", "ecommerce", "saas", "internet"]),
    help="行业风格（默认 finance）",
)
def manage_style(file, sheet, action, range, industry):
    """[商功章] 管理单元格预设样式

    使用场景：
    - 快速应用专业的表头样式
    - 为 KPI 数据添加可视化样式
    - 自动调整列宽以适应内容

    示例：
        abacus manage-style -f data.xlsx -s Sheet1 --action apply_header -r A1:E1 --industry finance
        abacus manage-style -f data.xlsx -s Dashboard --action apply_kpi -r B2:B5 --industry saas
        abacus manage-style -f data.xlsx -s Sheet1 --action auto_width
    """
    cap = registry.get("manage_style")
    result = cap.execute(
        None, file=file, sheet=sheet, action=action, range=range, industry=industry
    )
    click.echo(json.dumps(result, ensure_ascii=False, indent=2))


# 公式生成器命令
@main.command("generate-formula")
@click.option(
    "--formula-type", required=True, help="公式类型（如 vlookup、sumifs、if、today、npv、pmt 等）"
)
@click.option(
    "--params",
    "-p",
    required=True,
    help='公式参数，JSON 格式（如 \'{"lookup_value":"A1","table_array":"Sheet2!A:C","col_index":3}\'）',
)
@click.option("--file", "-f", help="Excel 文件路径（可选，指定后将公式写入文件）")
@click.option("--sheet", "-s", help="工作表名称（写入文件时可选）")
@click.option("--cell", "-c", help="单元格位置（如 A1，写入文件时可选）")
def generate_formula(formula_type, params, file, sheet, cell):
    """[方程章] 生成常用 Excel 公式

    使用场景：
    - 快速生成 VLOOKUP、SUMIFS 等常用公式
    - 不确定公式语法时获取正确公式
    - 批量生成类似公式

    示例：
        abacus generate-formula --formula-type vlookup -p '{"lookup_value":"A1","table_array":"Sheet2!A:C","col_index":3}'
        abacus generate-formula --formula-type sumifs -p '{"sum_range":"D:D","criteria_range1":"A:A","criteria1":"销售"}'
        abacus generate-formula --formula-type npv -p '{"rate":0.1,"values":"B1:B10"}' -f data.xlsx -s Calc -c A1
    """
    import json as json_mod

    params_dict = json_mod.loads(params)
    cap = registry.get("generate_formula")
    result = cap.execute(
        None, formula_type=formula_type, params=params_dict, file=file, sheet=sheet, cell=cell
    )
    click.echo(json.dumps(result, ensure_ascii=False, indent=2))


# 行业模板命令
@main.command("generate-template")
@click.option(
    "--industry",
    required=True,
    type=click.Choice(["finance", "ecommerce", "saas", "internet"]),
    help="行业类型",
)
@click.option(
    "--template-type",
    required=True,
    help="模板类型（如 income_statement、balance_sheet、cash_flow 等）",
)
@click.option("--output", "-o", required=True, help="输出文件路径（如 template.xlsx）")
def generate_template(industry, template_type, output):
    """[商功章] 生成行业专用 Excel 模板

    使用场景：
    - 快速生成财务报表模板
    - 创建电商数据分析模板
    - 生成 SaaS 业务指标模板

    示例：
        abacus generate-template --industry finance --template-type income_statement -o finance_template.xlsx
        abacus generate-template --industry ecommerce --template-type sales_report -o电商报表.xlsx
    """
    cap = registry.get("generate_template")
    result = cap.execute(None, industry=industry, template_type=template_type, output=output)
    click.echo(json.dumps(result, ensure_ascii=False, indent=2))


# 数据分析模块命令
@main.command("analyze-data")
@click.option(
    "--file", "-f", required=True, help="文件路径，支持 Excel 和 CSV 格式（如 data.xlsx、data.csv）"
)
@click.option("--sheet", "-s", help="工作表名称（Excel 文件时可选，默认第一个工作表）")
@click.option(
    "--analysis-type",
    default="auto",
    type=click.Choice(["auto", "summary", "correlation"]),
    help="分析类型：auto=自动选择，summary=摘要统计，correlation=相关性分析",
)
def analyze_data(file, sheet, analysis_type):
    """[勾股章] 智能数据分析

    使用场景：
    - 快速了解数据集的整体特征
    - 自动生成数据摘要报告
    - 发现数据中的模式和关联

    示例：
        abacus analyze-data -f data.xlsx --analysis-type auto
        abacus analyze-data -f data.csv --analysis-type summary
        abacus analyze-data -f data.xlsx -s Sales --analysis-type correlation
    """
    cap = registry.get("analyze_data")
    result = cap.execute(None, file=file, sheet=sheet, analysis_type=analysis_type)
    click.echo(json.dumps(result, ensure_ascii=False, indent=2))


@main.command("clean-data")
@click.option("--file", "-f", required=True, help="文件路径（如 data.xlsx）")
@click.option("--sheet", "-s", help="工作表名称（可选，默认第一个工作表）")
@click.option("--output", "-o", help="输出文件路径（可选，默认覆盖原文件）")
@click.option(
    "--operations",
    help='清洗操作列表，JSON 格式（如 \'[{"action":"remove_duplicates"},{"action":"fill_null","value":0}]\'）',
)
def clean_data(file, sheet, output, operations):
    """[粟米章] 数据清洗

    使用场景：
    - 删除重复数据
    - 填充空值
    - 修正数据格式问题
    - 数据预处理

    示例：
        abacus clean-data -f data.xlsx --operations '[{"action":"remove_duplicates"}]'
        abacus clean-data -f data.xlsx -s Sheet1 -o cleaned.xlsx --operations '[{"action":"fill_null","value":0}]'
    """
    import json as json_mod

    ops_list = json_mod.loads(operations) if operations else None
    cap = registry.get("clean_data")
    result = cap.execute(None, file=file, sheet=sheet, output=output, operations=ops_list)
    click.echo(json.dumps(result, ensure_ascii=False, indent=2))


@main.command("pivot-analysis")
@click.option("--file", "-f", required=True, help="文件路径（如 data.xlsx）")
@click.option("--group-by", required=True, help="分组字段名（如 Category、Region）")
@click.option("--value-field", required=True, help="值字段名（如 Sales、Quantity）")
@click.option("--sheet", "-s", help="工作表名称（可选，默认第一个工作表）")
@click.option(
    "--agg-function",
    default="sum",
    type=click.Choice(["sum", "mean", "count"]),
    help="聚合函数（默认 sum）",
)
@click.option("--output", "-o", help="输出文件路径（可选）")
def pivot_analysis(file, group_by, value_field, sheet, agg_function, output):
    """[衰分章] 数据透视分析

    使用场景：
    - 按维度汇总分析数据
    - 生成分组统计报表
    - 快速数据聚合

    示例：
        abacus pivot-analysis -f data.xlsx --group-by Category --value-field Sales --agg-function sum
        abacus pivot-analysis -f data.xlsx --group-by Region --value-field Quantity --agg-function mean -o result.xlsx
    """
    cap = registry.get("pivot_analysis")
    result = cap.execute(
        None,
        file=file,
        sheet=sheet,
        group_by=group_by,
        value_field=value_field,
        agg_function=agg_function,
        output=output,
    )
    click.echo(json.dumps(result, ensure_ascii=False, indent=2))


# 金融建模命令
@main.command("dcf-model")
@click.option("--output", "-o", required=True, help="输出文件路径（如 dcf_model.xlsx）")
@click.option("--revenue", required=True, type=float, help="当前营收（单位：万元，如 1000）")
@click.option(
    "--growth-rate", required=True, type=float, help="营收增长率（小数形式，如 0.15 表示 15%）"
)
@click.option(
    "--operating-margin", required=True, type=float, help="运营利润率（小数形式，如 0.2 表示 20%）"
)
@click.option("--tax-rate", required=True, type=float, help="税率（小数形式，如 0.25 表示 25%）")
@click.option(
    "--wacc", required=True, type=float, help="加权平均资本成本（小数形式，如 0.1 表示 10%）"
)
@click.option(
    "--terminal-growth", required=True, type=float, help="永续增长率（小数形式，如 0.03 表示 3%）"
)
@click.option("--capex-ratio", default=0.05, type=float, help="资本支出占营收比例（默认 0.05）")
@click.option("--nwc-ratio", default=0.1, type=float, help="净营运资本占营收比例（默认 0.1）")
def dcf_model(
    output,
    revenue,
    growth_rate,
    operating_margin,
    tax_rate,
    wacc,
    terminal_growth,
    capex_ratio,
    nwc_ratio,
):
    """[方程章] DCF 现金流折现估值模型

    使用场景：
    - 企业估值分析
    - 投资决策支持
    - 财务建模和预测

    示例：
        abacus dcf-model -o dcf.xlsx --revenue 1000 --growth-rate 0.15 --operating-margin 0.2 --tax-rate 0.25 --wacc 0.1 --terminal-growth 0.03
    """
    cap = registry.get("dcf_model")
    result = cap.execute(
        None,
        output=output,
        revenue=revenue,
        growth_rate=growth_rate,
        operating_margin=operating_margin,
        tax_rate=tax_rate,
        wacc=wacc,
        terminal_growth=terminal_growth,
        capex_ratio=capex_ratio,
        nwc_ratio=nwc_ratio,
    )
    click.echo(json.dumps(result, ensure_ascii=False, indent=2))


@main.command("lbo-model")
@click.option("--output", "-o", required=True, help="输出文件路径（如 lbo_model.xlsx）")
@click.option("--ebitda", required=True, type=float, help="目标公司 EBITDA（单位：万元，如 500）")
@click.option("--entry-multiple", required=True, type=float, help="入场 EV/EBITDA 倍数（如 10）")
@click.option("--exit-multiple", required=True, type=float, help="退出 EV/EBITDA 倍数（如 12）")
@click.option("--exit-year", default=5, type=int, help="退出年份（默认 5 年）")
@click.option(
    "--senior-debt-ratio", default=0.6, type=float, help="优先债务占总资本比例（默认 0.6）"
)
@click.option("--mezzanine-ratio", default=0.0, type=float, help="夹层债务占总资本比例（默认 0.0）")
@click.option("--interest-rate", default=0.06, type=float, help="债务利率（默认 0.06）")
def lbo_model(
    output,
    ebitda,
    entry_multiple,
    exit_multiple,
    exit_year,
    senior_debt_ratio,
    mezzanine_ratio,
    interest_rate,
):
    """[方程章] LBO 杠杆收购模型

    使用场景：
    - 私募股权投资分析
    - 杠杆收购可行性评估
    - 投资回报率测算

    示例：
        abacus lbo-model -o lbo.xlsx --ebitda 500 --entry-multiple 10 --exit-multiple 12 --exit-year 5
    """
    cap = registry.get("lbo_model")
    result = cap.execute(
        None,
        output=output,
        ebitda=ebitda,
        entry_multiple=entry_multiple,
        exit_multiple=exit_multiple,
        exit_year=exit_year,
        senior_debt_ratio=senior_debt_ratio,
        mezzanine_ratio=mezzanine_ratio,
        interest_rate=interest_rate,
    )
    click.echo(json.dumps(result, ensure_ascii=False, indent=2))


@main.command("variance-analysis")
@click.option("--file", "-f", required=True, help="Excel 文件路径（如 data.xlsx）")
@click.option("--budget-sheet", required=True, help="预算数据工作表名称（如 Budget）")
@click.option("--actual-sheet", required=True, help="实际数据工作表名称（如 Actual）")
@click.option("--output", "-o", help="输出文件路径（可选）")
@click.option(
    "--threshold",
    default=0.1,
    type=float,
    help="重要性阈值，小数形式（默认 0.1 表示 10%，超过此比例的差异会被标记）",
)
def variance_analysis(file, budget_sheet, actual_sheet, output, threshold):
    """[勾股章] 预算与实际差异分析

    使用场景：
    - 分析预算执行情况
    - 识别重大偏差项目
    - 财务控制和绩效评估

    示例：
        abacus variance-analysis -f data.xlsx --budget-sheet Budget --actual-sheet Actual --threshold 0.1
        abacus variance-analysis -f data.xlsx --budget-sheet Q1预算 --actual-sheet Q1实际 -o variance_report.xlsx
    """
    cap = registry.get("variance_analysis")
    result = cap.execute(
        None,
        file=file,
        budget_sheet=budget_sheet,
        actual_sheet=actual_sheet,
        output=output,
        threshold=threshold,
    )
    click.echo(json.dumps(result, ensure_ascii=False, indent=2))


# 报表生成命令
@main.command("create-basic-report")
@click.option("--data-source", "-d", required=True, help="数据源文件路径（CSV 格式，如 data.csv）")
@click.option("--output", "-o", required=True, help="输出文件路径（如 report.xlsx）")
@click.option("--sheet-name", default="Data", help="数据工作表名称（默认 Data）")
@click.option("--title", help="报表标题（如 '2024年度销售报表'）")
def create_basic_report(data_source, output, sheet_name, title):
    """[商功章] 生成基础 Excel 报表

    使用场景：
    - 从 CSV 数据快速生成 Excel 报表
    - 添加表头、格式和基本样式
    - 自动化报表生成

    示例：
        abacus create-basic-report -d sales.csv -o sales_report.xlsx --title '月度销售报表'
        abacus create-basic-report -d data.csv -o report.xlsx --sheet-name '数据汇总'
    """
    cap = registry.get("create_basic_report")
    result = cap.execute(
        None, data_source=data_source, output=output, sheet_name=sheet_name, title=title
    )
    click.echo(json.dumps(result, ensure_ascii=False, indent=2))


@main.command("create-advanced-report")
@click.option("--data-source", "-d", required=True, help="数据源文件路径（CSV 格式，如 data.csv）")
@click.option("--output", "-o", required=True, help="输出文件路径（如 report.xlsx）")
@click.option(
    "--chart-type",
    default="bar",
    type=click.Choice(["bar", "line", "pie"]),
    help="图表类型（默认 bar）",
)
@click.option("--include-dashboard/--no-dashboard", default=True, help="是否包含仪表板（默认包含）")
def create_advanced_report(data_source, output, chart_type, include_dashboard):
    """[商功章] 生成高级 Excel 报表（含图表和仪表板）

    使用场景：
    - 生成带图表的专业报表
    - 创建数据仪表板
    - 高级数据可视化报表

    示例：
        abacus create-advanced-report -d sales.csv -o advanced_report.xlsx --chart-type bar
        abacus create-advanced-report -d data.csv -o dashboard.xlsx --chart-type line --no-dashboard
    """
    cap = registry.get("create_advanced_report")
    result = cap.execute(
        None,
        data_source=data_source,
        output=output,
        chart_type=chart_type,
        include_dashboard=include_dashboard,
    )
    click.echo(json.dumps(result, ensure_ascii=False, indent=2))


@main.command("fill-template")
@click.option("--template", "-t", required=True, help="模板文件路径（如 template.xlsx）")
@click.option("--output", "-o", required=True, help="输出文件路径（如 filled.xlsx）")
@click.option("--data-source", "-d", help="数据源文件路径（CSV 格式，与 --data 二选一）")
@click.option(
    "--data",
    help='填充数据，JSON 格式（如 \'{"name":"张三","amount":1000}\'，与 --data-source 二选一）',
)
@click.option("--sheet-name", "-s", help="目标工作表名称（可选）")
@click.option("--start-cell", default="A1", help="数据起始单元格（默认 A1）")
def fill_template(template, output, data_source, data, sheet_name, start_cell):
    """[商功章] 基于模板填充数据

    使用场景：
    - 批量生成合同、发票等文档
    - 从数据源填充报表模板
    - 自动化文档生成

    示例：
        abacus fill-template -t template.xlsx -o output.xlsx --data '{"name":"张三","amount":1000}'
        abacus fill-template -t template.xlsx -o batch.xlsx -d data.csv --sheet-name '报表'
    """
    import json as json_mod

    data_dict = json_mod.loads(data) if data else None
    cap = registry.get("fill_template")
    result = cap.execute(
        None,
        template=template,
        output=output,
        data_source=data_source,
        data=data_dict,
        sheet_name=sheet_name,
        start_cell=start_cell,
    )
    click.echo(json.dumps(result, ensure_ascii=False, indent=2))


# CSV 处理命令
@main.command("merge-files")
@click.option(
    "--files",
    "-f",
    required=True,
    multiple=True,
    help="文件路径列表（可多次指定，如 -f file1.csv -f file2.csv）",
)
@click.option("--output", "-o", required=True, help="输出文件路径（如 merged.xlsx）")
@click.option(
    "--merge-type",
    default="concat",
    type=click.Choice(["concat", "merge", "join"]),
    help="合并类型：concat=纵向拼接，merge=横向合并，join=关联合并",
)
@click.option("--on", help="合并键（merge/join 时必填，如 'id'）")
@click.option("--dedup/--no-dedup", default=False, help="是否去重（默认不去重）")
@click.option("--dedup-columns", help="去重列，JSON 格式（如 ['id','name']，不填则按所有列去重）")
def merge_files(files, output, merge_type, on, dedup, dedup_columns):
    """[均输章] 合并多个 CSV/Excel 文件

    使用场景：
    - 合并多个 CSV 文件为一个 Excel
    - 按关键字段关联多个数据源
    - 数据整合和去重

    示例：
        abacus merge-files -f file1.csv -f file2.csv -o merged.xlsx
        abacus merge-files -f orders.csv -f customers.csv -o result.xlsx --merge-type join --on customer_id
        abacus merge-files -f data1.csv -f data2.csv -o unique.xlsx --dedup --dedup-columns '["id"]'
    """
    import json as json_mod

    dedup_cols = json_mod.loads(dedup_columns) if dedup_columns else None
    cap = registry.get("merge_files")
    result = cap.execute(
        None,
        files=list(files),
        output=output,
        merge_type=merge_type,
        on=on,
        dedup=dedup,
        dedup_columns=dedup_cols,
    )
    click.echo(json.dumps(result, ensure_ascii=False, indent=2))


@main.command("visualize-data")
@click.option("--file", "-f", required=True, help="CSV/Excel 文件路径（如 data.csv、data.xlsx）")
@click.option("--output", "-o", required=True, help="输出文件路径（如 visualization.xlsx）")
@click.option(
    "--chart-type",
    default="auto",
    type=click.Choice(["bar", "line", "pie", "auto"]),
    help="图表类型：auto=自动选择（默认）",
)
@click.option("--include-dashboard/--no-dashboard", default=True, help="是否包含仪表板（默认包含）")
@click.option("--include-stats/--no-stats", default=True, help="是否包含统计摘要（默认包含）")
def visualize_data(file, output, chart_type, include_dashboard, include_stats):
    """[勾股章] CSV 数据可视化

    使用场景：
    - 快速将 CSV 数据可视化
    - 自动生成图表和仪表板
    - 数据探索和展示

    示例：
        abacus visualize-data -f data.csv -o report.xlsx --chart-type auto
        abacus visualize-data -f sales.xlsx -o dashboard.xlsx --chart-type bar --no-stats
    """
    cap = registry.get("visualize_data")
    result = cap.execute(
        None,
        file=file,
        output=output,
        chart_type=chart_type,
        include_dashboard=include_dashboard,
        include_stats=include_stats,
    )
    click.echo(json.dumps(result, ensure_ascii=False, indent=2))


# 格式转换命令
@main.command("excel-to-markdown")
@click.option("--file", "-f", required=True, help="Excel 文件路径（如 data.xlsx）")
@click.option("--sheet", "-s", help="工作表名称（可选，默认转换所有工作表）")
@click.option("--output", "-o", help="输出文件路径（如 output.md，不填则输出到控制台）")
@click.option(
    "--merge-mode",
    default="tl",
    type=click.Choice(["tl", "fill"]),
    help="合并单元格处理模式：tl=使用左上角值（默认），fill=填充所有单元格",
)
@click.option("--include-styles/--no-styles", default=True, help="是否包含样式信息（默认包含）")
def excel_to_markdown(file, sheet, output, merge_mode, include_styles):
    """[均输章] 将 Excel 表格转换为 Markdown 格式

    使用场景：
    - 将 Excel 数据转换为 Markdown 用于文档
    - 在 GitHub、GitLab 等平台展示表格数据
    - 生成技术文档中的表格

    示例：
        abacus excel-to-markdown -f data.xlsx -o table.md
        abacus excel-to-markdown -f data.xlsx -s Sales --no-styles
        abacus excel-to-markdown -f data.xlsx -o output.md --merge-mode fill
    """
    cap = registry.get("excel_to_markdown")
    result = cap.execute(
        None,
        file=file,
        sheet=sheet,
        output=output,
        merge_mode=merge_mode,
        include_styles=include_styles,
    )
    click.echo(json.dumps(result, ensure_ascii=False, indent=2))


@main.command("split-sheet")
@click.option("--file", "-f", required=True, help="Excel 文件路径（如 data.xlsx）")
@click.option("--sheet", "-s", required=True, help="要拆分的工作表名称（如 Sheet1）")
@click.option("--output-dir", "-o", required=True, help="输出目录路径（如 ./output）")
@click.option(
    "--split-by",
    required=True,
    type=click.Choice(["column", "row_count", "range"]),
    help="拆分方式：column=按列值拆分，row_count=按行数拆分，range=按范围拆分",
)
@click.option("--split-column", help="拆分列名（split_by=column 时必填，如 Category）")
@click.option("--row-count", type=int, help="每个文件的行数（split_by=row_count 时必填，如 1000）")
@click.option("--prefix", default="split", help="输出文件前缀（默认 split）")
def split_sheet(file, sheet, output_dir, split_by, split_column, row_count, prefix):
    """[商功章] 将 Excel 工作表按条件拆分为多个文件

    使用场景：
    - 按类别将数据拆分为独立文件
    - 将大文件拆分为多个小文件
    - 数据分发和分批处理

    示例：
        abacus split-sheet -f data.xlsx -s Sheet1 -o ./output --split-by column --split-column Category
        abacus split-sheet -f data.xlsx -s Sheet1 -o ./output --split-by row_count --row-count 1000 --prefix data
    """
    cap = registry.get("split_sheet")
    result = cap.execute(
        None,
        file=file,
        sheet=sheet,
        output_dir=output_dir,
        split_by=split_by,
        split_column=split_column,
        row_count=row_count,
        prefix=prefix,
    )
    click.echo(json.dumps(result, ensure_ascii=False, indent=2))


@main.command()
@click.option("--file", required=True, help="账单 Excel 文件路径（如 bill.xlsx）")
@click.option("--output", help="输出文件路径（可选，默认覆盖原文件）")
@click.option(
    "--group-fields",
    multiple=True,
    help="分组字段列表（可多次指定，如 --group-fields Category --group-fields Month）",
)
@click.option(
    "--agg-fields",
    multiple=True,
    help="聚合字段列表（可多次指定，如 --agg-fields Amount --agg-fields Quantity）",
)
@click.option("--sheet-name", default="账单折扣总览", help="输出工作表名称（默认 '账单折扣总览'）")
def bill_pivot(file, output, group_fields, agg_fields, sheet_name):
    """[商功章] 账单透视表生成

    使用场景：
    - 按类别和时间汇总账单数据
    - 生成账单折扣总览报表
    - 费用分析和预算管理

    示例：
        abacus bill-pivot --file bill.xlsx --group-fields Category --group-fields Month --agg-fields Amount
        abacus bill-pivot --file bill.xlsx --output summary.xlsx --sheet-name '费用汇总'
    """
    cap = registry.get("bill_pivot")
    result = cap.execute(
        None,
        file=file,
        output=output,
        group_fields=list(group_fields) if group_fields else None,
        agg_fields=list(agg_fields) if agg_fields else None,
        sheet_name=sheet_name,
    )
    click.echo(json.dumps(result, ensure_ascii=False, indent=2))


@main.command("advanced-analysis")
@click.option("--file", "-f", required=True, help="文件路径（如 data.xlsx、data.csv）")
@click.option(
    "--analysis-type",
    required=True,
    type=click.Choice(["regression", "timeseries", "forecast"]),
    help="分析类型：regression=回归分析，timeseries=时间序列，forecast=预测",
)
@click.option("--sheet", "-s", help="工作表名称（可选）")
@click.option("--x-column", help="自变量列名（regression 时必填）")
@click.option("--y-column", help="因变量列名（regression/timeseries/forecast 时必填）")
@click.option("--periods", default=10, type=int, help="预测期数（forecast 时可选，默认 10）")
def advanced_analysis(file, analysis_type, sheet, x_column, y_column, periods):
    """[勾股章] 高级数据分析（回归、时间序列、预测）

    使用场景：
    - 线性回归分析，获取斜率、截距和 R² 值
    - 时间序列分析，识别趋势和季节性
    - 线性外推预测未来值

    示例：
        abacus advanced-analysis -f data.xlsx --analysis-type regression --x-column X --y-column Y
        abacus advanced-analysis -f data.xlsx --analysis-type timeseries --y-column Sales
        abacus advanced-analysis -f data.xlsx --analysis-type forecast --y-column Sales --periods 12
    """
    cap = registry.get("advanced_analysis")
    result = cap.execute(
        None,
        file=file,
        analysis_type=analysis_type,
        sheet=sheet,
        x_column=x_column,
        y_column=y_column,
        periods=periods,
    )
    click.echo(json.dumps(result, ensure_ascii=False, indent=2))


@main.command("transform-data")
@click.option("--file", "-f", required=True, help="文件路径（如 data.xlsx）")
@click.option(
    "--transform-type",
    required=True,
    type=click.Choice(["pivot", "melt", "merge", "reshape"]),
    help="转换类型：pivot=透视，melt=逆透视，merge=合并，reshape=重塑",
)
@click.option("--sheet", "-s", help="工作表名称（可选）")
@click.option("--params", "-p", help="转换参数，JSON 格式")
@click.option("--output", "-o", help="输出文件路径（可选）")
def transform_data(file, transform_type, sheet, params, output):
    """[粟米章] 高级数据转换（透视、转置、合并、重塑）

    使用场景：
    - 透视表转换（宽表聚合）
    - 逆透视（宽表转长表）
    - 合并多个数据源
    - 数据重塑

    示例：
        abacus transform-data -f data.xlsx --transform-type pivot -p '{"index":"Category","values":"Sales"}'
        abacus transform-data -f data.xlsx --transform-type melt -p '{"id_vars":["ID"]}'
        abacus transform-data -f data.xlsx --transform-type reshape -p '{"pivot_column":"Type","value_column":"Amount"}'
    """
    import json as json_mod

    params_dict = json_mod.loads(params) if params else {}
    cap = registry.get("transform_data")
    result = cap.execute(
        None,
        file=file,
        transform_type=transform_type,
        sheet=sheet,
        params=params_dict,
        output=output,
    )
    click.echo(json.dumps(result, ensure_ascii=False, indent=2))


# P1 新增命令
@main.command("insert-image")
@click.option("--file", "-f", required=True, help="Excel 文件路径（如 data.xlsx）")
@click.option("--sheet", "-s", required=True, help="工作表名称（如 Sheet1）")
@click.option("--cell", "-c", required=True, help="单元格位置（如 A1）")
@click.option("--image-path", required=True, help="图片文件路径（如 logo.png）")
@click.option("--width", type=int, help="图片宽度（像素，可选）")
@click.option("--height", type=int, help="图片高度（像素，可选）")
def insert_image(file, sheet, cell, image_path, width, height):
    """[商功章] 插入图片到单元格

    使用场景：
    - 在 Excel 中插入产品图片
    - 添加 Logo 或水印
    - 嵌入图表截图

    示例：
        abacus insert-image -f data.xlsx -s Sheet1 -c A1 --image-path logo.png
        abacus insert-image -f data.xlsx -s Sheet1 -c B5 --image-path photo.jpg --width 200 --height 150
    """
    cap = registry.get("insert_excel_image")
    result = cap.execute(
        None, file=file, sheet=sheet, cell=cell, image_path=image_path, width=width, height=height
    )
    click.echo(json.dumps(result, ensure_ascii=False, indent=2))


@main.command("group-rows")
@click.option("--file", "-f", required=True, help="Excel 文件路径（如 data.xlsx）")
@click.option("--sheet", "-s", required=True, help="工作表名称（如 Sheet1）")
@click.option("--start-row", required=True, type=int, help="起始行号（如 2）")
@click.option("--end-row", required=True, type=int, help="结束行号（如 10）")
@click.option("--level", default=1, type=int, help="分组层级（默认 1）")
def group_rows(file, sheet, start_row, end_row, level):
    """[商功章] 分组行（折叠/展开）

    使用场景：
    - 折叠明细数据
    - 创建层级结构
    - 报表分组展示

    示例：
        abacus group-rows -f data.xlsx -s Sheet1 --start-row 2 --end-row 10
        abacus group-rows -f data.xlsx -s Sheet1 --start-row 5 --end-row 20 --level 2
    """
    cap = registry.get("group_rows")
    result = cap.execute(
        None, file=file, sheet=sheet, start_row=start_row, end_row=end_row, level=level
    )
    click.echo(json.dumps(result, ensure_ascii=False, indent=2))


@main.command()
@click.option("--file", "-f", required=True, help="Excel 文件路径（如 data.xlsx）")
@click.option("--sheet", "-s", required=True, help="工作表名称（如 Sheet1）")
@click.option("--range", "-r", required=True, help="数据范围，Excel 格式（如 A1:D100）")
@click.option("--group-column", required=True, help="分组列名（如 Category）")
@click.option(
    "--function",
    default="sum",
    type=click.Choice(["sum", "mean", "count", "min", "max"]),
    help="聚合函数（默认 sum）",
)
def subtotal(file, sheet, range, group_column, function):
    """[衰分章] 分类汇总（按字段分组聚合）

    使用场景：
    - 按类别汇总销售额
    - 按部门统计平均工资
    - 按产品分组计算总量

    示例：
        abacus subtotal -f data.xlsx -s Sheet1 -r A1:D100 --group-column Category --function sum
        abacus subtotal -f data.xlsx -s Sales -r A1:E500 --group-column Region --function mean
    """
    cap = registry.get("subtotal")
    result = cap.execute(
        None, file=file, sheet=sheet, range=range, group_column=group_column, function=function
    )
    click.echo(json.dumps(result, ensure_ascii=False, indent=2))


@main.command()
@click.option("--file", "-f", required=True, help="Excel 文件路径（如 data.xlsx）")
@click.option("--sheet", "-s", required=True, help="工作表名称（如 Sheet1）")
@click.option("--range", "-r", required=True, help="数据范围，Excel 格式（如 A1:D10）")
@click.option("--output-sheet", help="输出工作表名称（可选，默认覆盖原工作表）")
def transpose(file, sheet, range, output_sheet):
    """[粟米章] 转置数据（行列互换）

    使用场景：
    - 将行数据转为列
    - 数据结构调整
    - 报表格式转换

    示例：
        abacus transpose -f data.xlsx -s Sheet1 -r A1:D10 --output-sheet Transposed
        abacus transpose -f data.xlsx -s Data -r A1:C5
    """
    cap = registry.get("transpose")
    result = cap.execute(None, file=file, sheet=sheet, range=range, output_sheet=output_sheet)
    click.echo(json.dumps(result, ensure_ascii=False, indent=2))


@main.command("text-to-columns")
@click.option("--file", "-f", required=True, help="Excel 文件路径（如 data.xlsx）")
@click.option("--sheet", "-s", required=True, help="工作表名称（如 Sheet1）")
@click.option("--column", "-c", required=True, help="列标识（如 A）")
@click.option("--delimiter", default=",", help="分隔符（默认逗号）")
def text_to_columns(file, sheet, column, delimiter):
    """[粟米章] 文本分列（按分隔符拆分）

    使用场景：
    - 将逗号分隔的数据拆分到多列
    - 处理 CSV 格式的数据
    - 拆分姓名、地址等复合字段

    示例：
        abacus text-to-columns -f data.xlsx -s Sheet1 -c A --delimiter ","
        abacus text-to-columns -f data.xlsx -s Sheet1 -c B --delimiter "|"
    """
    cap = registry.get("text_to_columns")
    result = cap.execute(None, file=file, sheet=sheet, column=column, delimiter=delimiter)
    click.echo(json.dumps(result, ensure_ascii=False, indent=2))


@main.command("transform-pipeline")
@click.option("--file", "-f", required=True, help="Excel 文件路径（如 data.xlsx）")
@click.option("--sheet", "-s", help="工作表名称（可选，默认活动工作表）")
@click.option(
    "--steps",
    "-p",
    required=True,
    help='转换步骤列表，JSON 格式（如 \'[{"type":"convert_type","range":"A1:A10","target_type":"float"}]\'）',
)
@click.option(
    "--stop-on-error/--continue-on-error", default=True, help="遇到错误是否停止（默认停止）"
)
def transform_pipeline(file, sheet, steps, stop_on_error):
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

    示例：
        abacus transform-pipeline -f data.xlsx -p '[{"type":"convert_type","range":"A1:A10","target_type":"float"},{"type":"fill_value","range":"B1:B10","value":0}]'
        abacus transform-pipeline -f data.xlsx -s Sheet1 -p '[{"type":"standardize","range":"C1:C10","text_case":"upper"}]'
    """
    import json as json_mod

    steps_list = json_mod.loads(steps)
    cap = registry.get("transform_pipeline")
    result = cap.execute(
        None, file=file, sheet=sheet, steps=steps_list, stop_on_error=stop_on_error
    )
    click.echo(json.dumps(result, ensure_ascii=False, indent=2))


@main.command("auto-sum")
@click.option("--file", "-f", required=True, help="Excel 文件路径（如 data.xlsx）")
@click.option("--sheet", "-s", required=True, help="工作表名称（如 Sheet1）")
@click.option("--range", "-r", required=True, help="数据范围，Excel 格式（如 A1:D10）")
@click.option(
    "--direction",
    default="down",
    type=click.Choice(["down", "right"]),
    help="求和方向：down=向下求和，right=向右求和",
)
def auto_sum(file, sheet, range, direction):
    """[方程章] 自动求和（在范围内设置 SUM 公式）

    使用场景：
    - 快速为数据列添加合计行
    - 为数据行添加合计列
    - 批量设置求和公式

    示例：
        abacus auto-sum -f data.xlsx -s Sheet1 -r A1:D10 --direction down
        abacus auto-sum -f data.xlsx -s Sheet1 -r A1:E5 --direction right
    """
    cap = registry.get("auto_sum")
    result = cap.execute(None, file=file, sheet=sheet, range=range, direction=direction)
    click.echo(json.dumps(result, ensure_ascii=False, indent=2))


@main.command("export-chart-image")
@click.option("--file", "-f", required=True, help="Excel 文件路径（如 data.xlsx）")
@click.option("--sheet", "-s", required=True, help="工作表名称（如 Sheet1）")
@click.option("--chart-index", required=True, type=int, help="图表索引（从 0 开始）")
@click.option("--output", "-o", required=True, help="输出图片路径（如 chart.png）")
def export_chart_image(file, sheet, chart_index, output):
    """[商功章] 导出图表为图片（使用 LibreOffice）

    使用场景：
    - 将 Excel 图表导出为 PNG 图片
    - 在报告中嵌入图表
    - 分享图表图片

    示例：
        abacus export-chart-image -f data.xlsx -s Sheet1 --chart-index 0 -o chart.png
        abacus export-chart-image -f report.xlsx -s Dashboard --chart-index 2 -o output.png
    """
    cap = registry.get("export_chart_as_image")
    result = cap.execute(None, file=file, sheet=sheet, chart_index=chart_index, output=output)
    click.echo(json.dumps(result, ensure_ascii=False, indent=2))


# ============================================================================
# 知识图谱 CLI 命令
# ============================================================================


@main.command("skill-search")
@click.argument("query")
@click.option("--limit", "-l", default=10, type=int, help="返回结果数量（默认 10）")
def skill_search(query, limit):
    """[知识图谱] 搜索 SKILL.md 和知识文件

    使用 FTS5 全文搜索，返回匹配的 skill 和知识文件。

    示例：
        abacus skill-search "公式验证"
        abacus skill-search "chart" --limit 5
    """
    from .skill.indexer import SkillIndexer

    indexer = SkillIndexer(Path("skills_index.db"))
    try:
        results = indexer.search(query, limit)
        click.echo(
            json.dumps(
                {"query": query, "results": results, "total": len(results)},
                ensure_ascii=False,
                indent=2,
            )
        )
    finally:
        indexer.close()


@main.command("skill-graph")
@click.argument("skill_name")
def skill_graph(skill_name):
    """[知识图谱] 获取 skill 关联图谱

    返回指定 skill 的完整关联信息：使用的能力、引用的知识文件、关联的其他 skill。

    示例：
        abacus skill-graph abacus-field
        abacus skill-graph abacus-work
    """
    from .skill.indexer import SkillIndexer

    indexer = SkillIndexer(Path("skills_index.db"))
    try:
        result = indexer.graph(skill_name)
        click.echo(json.dumps(result, ensure_ascii=False, indent=2))
    finally:
        indexer.close()


@main.command("skill-index-build")
def skill_index_build():
    """[知识图谱] 重建 SKILL.md 索引

    扫描 skills/ 目录，重建知识图谱索引。
    包括：9 个章节 SKILL.md + 29 个知识文件。

    示例：
        abacus skill-index-build
    """
    from .skill.indexer import SkillIndexer

    indexer = SkillIndexer(Path("skills_index.db"))
    try:
        indexer.scan(Path("skills"))
        stats = indexer.stats()
        click.echo(json.dumps({"status": "ok", "stats": stats}, ensure_ascii=False, indent=2))
    finally:
        indexer.close()


@main.command("skill-stats")
def skill_stats():
    """[知识图谱] 获取索引统计

    返回知识图谱的统计信息。

    示例：
        abacus skill-stats
    """
    from .skill.indexer import SkillIndexer

    indexer = SkillIndexer(Path("skills_index.db"))
    try:
        result = indexer.stats()
        click.echo(json.dumps(result, ensure_ascii=False, indent=2))
    finally:
        indexer.close()


# ============================================================================
# 商功章扩展命令
# ============================================================================


@main.command("pack-file")
@click.option("--file", "-f", required=True, help="Excel 文件路径（如 data.xlsx）")
@click.option("--output", "-o", help="输出 ZIP 文件路径（可选）")
def pack_file(file, output):
    """[商功章] 将 Excel 文件打包为 ZIP

    使用场景：
    - 需要将 Excel 文件打包为 ZIP
    - 用于调试和分析
    - 文件备份

    示例：
        abacus pack-file -f data.xlsx
        abacus pack-file -f data.xlsx -o output.zip
    """
    cap = registry.get("pack_file")
    result = cap.execute(None, file=file, output=output)
    click.echo(json.dumps(result, ensure_ascii=False, indent=2))


@main.command("unpack-file")
@click.option("--file", "-f", required=True, help="ZIP 文件路径（如 file.zip）")
@click.option("--output", "-o", help="输出目录（可选）")
def unpack_file(file, output):
    """[商功章] 将 ZIP 解包为 Excel 文件

    使用场景：
    - 需要将 ZIP 解包为 Excel 文件
    - 用于调试和分析
    - 文件恢复

    示例：
        abacus unpack-file -f file.zip
        abacus unpack-file -f file.zip -o output_dir
    """
    cap = registry.get("unpack_file")
    result = cap.execute(None, file=file, output=output)
    click.echo(json.dumps(result, ensure_ascii=False, indent=2))


@main.command("protect-workbook")
@click.option("--file", "-f", required=True, help="Excel 文件路径（如 data.xlsx）")
@click.option("--password", "-p", help="保护密码（可选）")
def protect_workbook(file, password):
    """[商功章] 保护工作簿

    使用场景：
    - 需要保护工作簿结构不被修改
    - 防止添加/删除/重命名工作表
    - 保护工作簿配置

    示例：
        abacus protect-workbook -f data.xlsx
        abacus protect-workbook -f data.xlsx -p mypassword
    """
    cap = registry.get("protect_workbook")
    result = cap.execute(None, file=file, password=password)
    click.echo(json.dumps(result, ensure_ascii=False, indent=2))


@main.command("protect-sheet")
@click.option("--file", "-f", required=True, help="Excel 文件路径（如 data.xlsx）")
@click.option("--sheet", "-s", required=True, help="工作表名称（如 Sheet1）")
@click.option("--password", "-p", help="保护密码（可选）")
def protect_sheet(file, sheet, password):
    """[商功章] 保护工作表

    使用场景：
    - 需要保护工作表不被修改
    - 防止意外编辑
    - 保护公式和数据

    示例：
        abacus protect-sheet -f data.xlsx -s Sheet1
        abacus protect-sheet -f data.xlsx -s Sheet1 -p mypassword
    """
    cap = registry.get("protect_sheet")
    result = cap.execute(None, file=file, sheet=sheet, password=password)
    click.echo(json.dumps(result, ensure_ascii=False, indent=2))


@main.command("set-array-formula")
@click.option("--file", "-f", required=True, help="Excel 文件路径（如 data.xlsx）")
@click.option("--sheet", "-s", required=True, help="工作表名称（如 Sheet1）")
@click.option("--range", "-r", required=True, help="数组公式范围（如 A1:A10）")
@click.option("--formula", required=True, help="数组公式内容")
def set_array_formula(file, sheet, range, formula):
    """[商功章] 设置数组公式

    使用场景：
    - 需要创建数组公式
    - 执行多单元格计算
    - 复杂的数据分析

    示例：
        abacus set-array-formula -f data.xlsx -s Sheet1 -r A1:A10 --formula "SUM(B1:B10*C1:C10)"
    """
    cap = registry.get("set_array_formula")
    result = cap.execute(None, file=file, sheet=sheet, range=range, formula=formula)
    click.echo(json.dumps(result, ensure_ascii=False, indent=2))


# ============================================================================
# P2 新增命令 - 智能数据匹配和关联
# ============================================================================


@main.command("fuzzy-match")
@click.option("--file", "-f", required=True, help="Excel 文件路径（如 data.xlsx）")
@click.option("--sheet", "-s", required=True, help="工作表名称（如 Sheet1）")
@click.option("--target-columns", "-t", required=True, help="目标列名列表（JSON 格式）")
@click.option("--threshold", type=float, default=0.6, help="相似度阈值（0-1，默认 0.6）")
def fuzzy_match(file, sheet, target_columns, threshold):
    """[粟米章] 模糊匹配列名（自动识别相似列名）

    使用场景：
    - 自动识别相似列名（如"销售额"和"销售金额"）
    - 多表合并时自动对齐列名
    - 数据整合时的列名匹配

    示例：
        abacus fuzzy-match -f data.xlsx -s Sheet1 -t '["销售额", "利润"]'
        abacus fuzzy-match -f data.xlsx -s Sheet1 -t '["Sales", "Profit"]' --threshold 0.7
    """
    import json as json_mod

    target_columns_list = json_mod.loads(target_columns)
    cap = registry.get("fuzzy_match_columns")
    result = cap.execute(
        None, file=file, sheet=sheet, target_columns=target_columns_list, threshold=threshold
    )
    click.echo(json.dumps(result, ensure_ascii=False, indent=2))


@main.command("quality-check")
@click.option("--file", "-f", required=True, help="Excel 文件路径（如 data.xlsx）")
@click.option("--sheet", "-s", required=True, help="工作表名称（如 Sheet1）")
@click.option("--range", "-r", help="数据范围（可选，默认全部）")
def quality_check(file, sheet, range):
    """[盈不足章] 数据质量检测（自动检测空值、异常值、重复数据）

    使用场景：
    - 自动检测空值、异常值、重复数据、格式不一致
    - 数据清洗前的质量评估
    - 数据质量报告

    示例：
        abacus quality-check -f data.xlsx -s Sheet1
        abacus quality-check -f data.xlsx -s Sheet1 -r A1:D100
    """
    cap = registry.get("data_quality_check")
    result = cap.execute(None, file=file, sheet=sheet, range=range)
    click.echo(json.dumps(result, ensure_ascii=False, indent=2))


@main.command("join-tables")
@click.option("--left-file", "-lf", required=True, help="左表 Excel 文件路径")
@click.option("--left-sheet", "-ls", required=True, help="左表工作表名称")
@click.option("--right-file", "-rf", required=True, help="右表 Excel 文件路径")
@click.option("--right-sheet", "-rs", required=True, help="右表工作表名称")
@click.option("--on", "-k", required=True, help="关联键（JSON 格式）")
@click.option("--how", "-h", default="inner", help="关联类型（left/right/inner/outer）")
@click.option("--output", "-o", help="输出文件路径（可选）")
def join_tables(left_file, left_sheet, right_file, right_sheet, on, how, output):
    """[均输章] SQL 风格关联（LEFT/RIGHT/INNER/OUTER JOIN）

    使用场景：
    - 替代 VLOOKUP，多表关联
    - LEFT/RIGHT/INNER/OUTER JOIN
    - 数据整合

    示例：
        abacus join-tables -lf left.xlsx -ls Sheet1 -rf right.xlsx -rs Sheet1 -k '["ID"]'
        abacus join-tables -lf left.xlsx -ls Sheet1 -rf right.xlsx -rs Sheet1 -k '["ID"]' --how left -o output.xlsx
    """
    import json as json_mod

    on_list = json_mod.loads(on)
    cap = registry.get("join_tables")
    result = cap.execute(
        None,
        left_file=left_file,
        left_sheet=left_sheet,
        right_file=right_file,
        right_sheet=right_sheet,
        on=on_list,
        how=how,
        output=output,
    )
    click.echo(json.dumps(result, ensure_ascii=False, indent=2))


@main.command("batch-merge")
@click.option("--folder", "-d", required=True, help="文件夹路径")
@click.option("--pattern", "-p", default="*.xlsx", help="文件匹配模式（默认 *.xlsx）")
@click.option("--sheet", "-s", help="工作表名称（可选）")
@click.option("--output", "-o", required=True, help="输出文件路径")
def batch_merge(folder, pattern, sheet, output):
    """[均输章] 多表批量合并（从文件夹批量合并多个 Excel 文件）

    使用场景：
    - 年终汇总、月度数据合并
    - 从文件夹批量合并多个 Excel 文件
    - 数据整合

    示例：
        abacus batch-merge -d /path/to/data -o merged.xlsx
        abacus batch-merge -d /path/to/data -p "*.xlsx" -s Sheet1 -o merged.xlsx
    """
    cap = registry.get("batch_merge")
    result = cap.execute(None, folder=folder, pattern=pattern, sheet=sheet, output=output)
    click.echo(json.dumps(result, ensure_ascii=False, indent=2))


@main.command("auto-type-infer")
@click.option("--file", "-f", required=True, help="Excel 文件路径（如 data.xlsx）")
@click.option("--sheet", "-s", required=True, help="工作表名称（如 Sheet1）")
@click.option("--range", "-r", help="数据范围（可选，默认全部）")
@click.option("--output", "-o", help="输出文件路径（可选）")
def auto_type_infer(file, sheet, range, output):
    """[粟米章] 自动类型推断（自动检测并转换数据类型）

    使用场景：
    - 自动检测并转换数据类型（文本→数字、日期等）
    - 导入数据时自动标准化
    - 数据清洗

    示例：
        abacus auto-type-infer -f data.xlsx -s Sheet1
        abacus auto-type-infer -f data.xlsx -s Sheet1 -o output.xlsx
    """
    cap = registry.get("auto_type_infer")
    result = cap.execute(None, file=file, sheet=sheet, range=range, output=output)
    click.echo(json.dumps(result, ensure_ascii=False, indent=2))


@main.command("standardize")
@click.option("--file", "-f", required=True, help="Excel 文件路径（如 data.xlsx）")
@click.option("--sheet", "-s", required=True, help="工作表名称（如 Sheet1）")
@click.option("--date-format", "-d", help="日期格式（如 %Y-%m-%d）")
@click.option("--number-format", "-n", help="数字格式（如 %.2f）")
@click.option("--text-case", "-t", help="文本大小写（lower/upper/title）")
@click.option("--output", "-o", help="输出文件路径（可选）")
def standardize(file, sheet, date_format, number_format, text_case, output):
    """[粟米章] 数据标准化（统一日期、数字、文本格式）

    使用场景：
    - 统一日期格式（如 %Y-%m-%d）
    - 统一数字格式（如 %.2f）
    - 统一文本大小写（lower/upper/title）

    示例：
        abacus standardize -f data.xlsx -s Sheet1 -d "%Y-%m-%d"
        abacus standardize -f data.xlsx -s Sheet1 -t lower -o output.xlsx
    """
    cap = registry.get("standardize_data")
    result = cap.execute(
        None,
        file=file,
        sheet=sheet,
        date_format=date_format,
        number_format=number_format,
        text_case=text_case,
        output=output,
    )
    click.echo(json.dumps(result, ensure_ascii=False, indent=2))


@main.command("summary-report")
@click.option("--file", "-f", required=True, help="Excel 文件路径（如 data.xlsx）")
@click.option("--sheet", "-s", required=True, help="工作表名称（如 Sheet1）")
def summary_report(file, sheet):
    """[商功章] 数据摘要报告（自动生成数据摘要）

    使用场景：
    - 快速了解数据全貌
    - 自动生成数据摘要（行列数、类型分布、质量问题）
    - 数据探索

    示例：
        abacus summary-report -f data.xlsx -s Sheet1
    """
    cap = registry.get("generate_summary_report")
    result = cap.execute(None, file=file, sheet=sheet)
    click.echo(json.dumps(result, ensure_ascii=False, indent=2))


@main.command("diff-report")
@click.option("--old-file", "-of", required=True, help="旧版本文件路径")
@click.option("--old-sheet", "-os", required=True, help="旧版本工作表名称")
@click.option("--new-file", "-nf", required=True, help="新版本文件路径")
@click.option("--new-sheet", "-ns", required=True, help="新版本工作表名称")
@click.option("--key-columns", "-k", help="用于匹配的键列（JSON 格式）")
def diff_report(old_file, old_sheet, new_file, new_sheet, key_columns):
    """[商功章] 变化检测报告（对比两个版本的数据，检测变化）

    使用场景：
    - 数据版本对比
    - 审计追踪
    - 检测数据变化

    示例：
        abacus diff-report -of old.xlsx -os Sheet1 -nf new.xlsx -ns Sheet1
        abacus diff-report -of old.xlsx -os Sheet1 -nf new.xlsx -ns Sheet1 -k '["ID"]'
    """
    import json as json_mod

    key_columns_list = json_mod.loads(key_columns) if key_columns else None
    cap = registry.get("generate_diff_report")
    result = cap.execute(
        None,
        old_file=old_file,
        old_sheet=old_sheet,
        new_file=new_file,
        new_sheet=new_sheet,
        key_columns=key_columns_list,
    )
    click.echo(json.dumps(result, ensure_ascii=False, indent=2))


@main.command("data-view")
@click.option("--file", "-f", required=True, help="Excel 文件路径（如 data.xlsx）")
@click.option("--sheet", "-s", required=True, help="工作表名称（如 Sheet1）")
@click.option("--action", "-a", required=True, help="操作（create/list/get/delete）")
@click.option("--view-name", "-v", help="视图名称")
@click.option("--columns", "-c", help="视图包含的列（JSON 格式）")
@click.option("--filters", "-fi", help="过滤条件（JSON 格式）")
def data_view(file, sheet, action, view_name, columns, filters):
    """[商功章] 数据视图管理（创建和管理不同角色的数据视图）

    使用场景：
    - 创建和管理不同角色的数据视图
    - 不同角色看不同数据子集
    - 数据权限管理

    示例：
        abacus data-view -f data.xlsx -s Sheet1 -a list
        abacus data-view -f data.xlsx -s Sheet1 -a create -v sales_view -c '["ID", "Name", "Sales"]'
        abacus data-view -f data.xlsx -s Sheet1 -a get -v sales_view
    """
    import json as json_mod

    columns_list = json_mod.loads(columns) if columns else None
    filters_dict = json_mod.loads(filters) if filters else None
    cap = registry.get("manage_data_view")
    result = cap.execute(
        None,
        file=file,
        sheet=sheet,
        action=action,
        view_name=view_name,
        columns=columns_list,
        filters=filters_dict,
    )
    click.echo(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
