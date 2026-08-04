"""Abacus 数据分析示例"""

from abacus.core import (
    AnalyzeStatsCapability,
    AnalyzeTrendCapability,
    AnalyzeCorrelationCapability,
    AnalyzeDataCapability,
    VisualizeDataCapability,
)
import json


def example_stats_analysis():
    """示例：统计分析"""
    cap = AnalyzeStatsCapability()
    result = cap.execute(
        None,
        file="data.xlsx",
        sheet="Sales",
        range="B2:D100"
    )
    print("统计分析结果：")
    print(json.dumps(result, ensure_ascii=False, indent=2))


def example_trend_analysis():
    """示例：趋势分析"""
    cap = AnalyzeTrendCapability()
    result = cap.execute(
        None,
        file="data.xlsx",
        sheet="Sales",
        range="A1:B100",
        value_column="Sales"
    )
    print("趋势分析结果：")
    print(json.dumps(result, ensure_ascii=False, indent=2))


def example_correlation_analysis():
    """示例：相关性分析"""
    cap = AnalyzeCorrelationCapability()
    result = cap.execute(
        None,
        file="data.xlsx",
        sheet="Sales",
        range="B1:C100",
        column1="Sales",
        column2="Profit"
    )
    print("相关性分析结果：")
    print(json.dumps(result, ensure_ascii=False, indent=2))


def example_smart_analysis():
    """示例：智能数据分析"""
    cap = AnalyzeDataCapability()
    result = cap.execute(
        None,
        file="data.xlsx",
        analysis_type="auto"
    )
    print("智能分析结果：")
    print(json.dumps(result, ensure_ascii=False, indent=2))


def example_visualization():
    """示例：数据可视化"""
    cap = VisualizeDataCapability()
    result = cap.execute(
        None,
        file="data.xlsx",
        output="visualization.xlsx",
        chart_type="bar",
        include_dashboard=True
    )
    print("可视化报告已创建：")
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    print("=== Abacus 数据分析示例 ===\n")

    print("1. 统计分析:")
    # example_stats_analysis()

    print("\n2. 趋势分析:")
    # example_trend_analysis()

    print("\n3. 相关性分析:")
    # example_correlation_analysis()

    print("\n4. 智能数据分析:")
    # example_smart_analysis()

    print("\n5. 数据可视化:")
    # example_visualization()

    print("\n注意：请先创建 data.xlsx 文件后再运行示例")
