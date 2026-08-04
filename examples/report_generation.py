"""Abacus 报表生成示例"""

from abacus.core import (
    BasicReportCapability,
    AdvancedReportCapability,
    TemplateReportCapability,
    BillPivotCapability,
)
import json


def example_basic_report():
    """示例：生成基础报表"""
    cap = BasicReportCapability()
    result = cap.execute(
        None,
        data_source="data.csv",
        output="basic_report.xlsx",
        title="销售报表",
        sheet_name="Sales"
    )
    print("基础报表已生成：")
    print(json.dumps(result, ensure_ascii=False, indent=2))


def example_advanced_report():
    """示例：生成高级报表"""
    cap = AdvancedReportCapability()
    result = cap.execute(
        None,
        data_source="data.csv",
        output="advanced_report.xlsx",
        chart_type="bar",
        include_dashboard=True
    )
    print("高级报表已生成：")
    print(json.dumps(result, ensure_ascii=False, indent=2))


def example_template_report():
    """示例：基于模板生成报表"""
    cap = TemplateReportCapability()
    result = cap.execute(
        None,
        template="template.xlsx",
        output="filled_report.xlsx",
        data={
            "A2": "2026-Q1",
            "B2": 1000000,
            "C2": 200000
        }
    )
    print("模板报表已生成：")
    print(json.dumps(result, ensure_ascii=False, indent=2))


def example_bill_pivot():
    """示例：生成账单透视表"""
    cap = BillPivotCapability()
    result = cap.execute(
        None,
        file="bill.xlsx",
        output="bill_pivot.xlsx",
        group_fields=["账务账期", "产品"],
        agg_fields=["原价", "折后价", "应付金额"]
    )
    print("账单透视表已生成：")
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    print("=== Abacus 报表生成示例 ===\n")

    print("1. 基础报表:")
    # example_basic_report()

    print("\n2. 高级报表:")
    # example_advanced_report()

    print("\n3. 模板报表:")
    # example_template_report()

    print("\n4. 账单透视表:")
    # example_bill_pivot()

    print("\n注意：请先创建示例文件后再运行")
