"""Abacus 基础使用示例"""

from abacus.core import (
    MeasureRangeCapability,
    MeasureCellsCapability,
    MeasureStructureCapability,
    ConvertFormatCapability,
    GroupByCapability,
    AnalyzeStatsCapability,
)
import json


def example_read_range():
    """示例：读取 Excel 范围数据"""
    cap = MeasureRangeCapability()
    result = cap.execute(
        None,
        file="data.xlsx",
        sheet="Sales",
        range="A1:D10"
    )
    print(json.dumps(result, ensure_ascii=False, indent=2))


def example_read_cells():
    """示例：读取单元格详情"""
    cap = MeasureCellsCapability()
    result = cap.execute(
        None,
        file="data.xlsx",
        sheet="Sales",
        range="A1:B5"
    )
    print(json.dumps(result, ensure_ascii=False, indent=2))


def example_read_structure():
    """示例：查看工作表结构"""
    cap = MeasureStructureCapability()
    result = cap.execute(None, file="data.xlsx")
    print(json.dumps(result, ensure_ascii=False, indent=2))


def example_convert_format():
    """示例：转换数据格式"""
    cap = ConvertFormatCapability()
    result = cap.execute(
        None,
        file="data.xlsx",
        sheet="Sales",
        range="B2:B100",
        format_type="number"
    )
    print(json.dumps(result, ensure_ascii=False, indent=2))


def example_group_by():
    """示例：按字段分组"""
    cap = GroupByCapability()
    result = cap.execute(
        None,
        file="data.xlsx",
        sheet="Sales",
        range="A1:C100",
        group_columns=["Region"],
        value_field="Sales",
        agg_function="sum"
    )
    print(json.dumps(result, ensure_ascii=False, indent=2))


def example_analyze_stats():
    """示例：统计分析"""
    cap = AnalyzeStatsCapability()
    result = cap.execute(
        None,
        file="data.xlsx",
        sheet="Sales",
        range="B2:B100"
    )
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    print("=== Abacus 基础使用示例 ===\n")

    print("1. 读取范围数据:")
    # example_read_range()

    print("\n2. 读取单元格详情:")
    # example_read_cells()

    print("\n3. 查看工作表结构:")
    # example_read_structure()

    print("\n4. 转换数据格式:")
    # example_convert_format()

    print("\n5. 按字段分组:")
    # example_group_by()

    print("\n6. 统计分析:")
    # example_analyze_stats()

    print("\n注意：请先创建 data.xlsx 文件后再运行示例")
