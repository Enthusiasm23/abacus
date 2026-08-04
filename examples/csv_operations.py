"""Abacus CSV 操作示例"""

from abacus.core import (
    CSVMergeCapability,
    CSVVisualizeCapability,
    ExportDataCapability,
    ImportDataCapability,
)
import json


def example_merge_csv():
    """示例：合并多个 CSV 文件"""
    cap = CSVMergeCapability()
    result = cap.execute(
        None,
        files=["data1.csv", "data2.csv", "data3.csv"],
        output="merged.xlsx",
        merge_type="concat",
        dedup=True
    )
    print("CSV 合并结果：")
    print(json.dumps(result, ensure_ascii=False, indent=2))


def example_visualize_csv():
    """示例：CSV 数据可视化"""
    cap = CSVVisualizeCapability()
    result = cap.execute(
        None,
        file="data.csv",
        output="visualization.xlsx",
        chart_type="bar",
        include_dashboard=True
    )
    print("可视化报告已创建：")
    print(json.dumps(result, ensure_ascii=False, indent=2))


def example_export_to_csv():
    """示例：导出为 CSV"""
    cap = ExportDataCapability()
    result = cap.execute(
        None,
        file="data.xlsx",
        sheet="Sales",
        range="A1:D100",
        output="exported.csv",
        format="csv"
    )
    print("导出结果：")
    print(json.dumps(result, ensure_ascii=False, indent=2))


def example_import_from_csv():
    """示例：从 CSV 导入"""
    cap = ImportDataCapability()
    result = cap.execute(
        None,
        file="imported.xlsx",
        source="data.csv",
        source_type="csv",
        sheet="Imported"
    )
    print("导入结果：")
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    print("=== Abacus CSV 操作示例 ===\n")

    print("1. 合并 CSV 文件:")
    # example_merge_csv()

    print("\n2. CSV 数据可视化:")
    # example_visualize_csv()

    print("\n3. 导出为 CSV:")
    # example_export_to_csv()

    print("\n4. 从 CSV 导入:")
    # example_import_from_csv()

    print("\n注意：请先创建示例文件后再运行")
