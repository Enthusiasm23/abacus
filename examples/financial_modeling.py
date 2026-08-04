"""Abacus 金融建模示例"""

from abacus.core import (
    DCFCapability,
    LBOCapability,
    VarianceCapability,
)
import json


def example_dcf_model():
    """示例：创建 DCF 估值模型"""
    cap = DCFCapability()
    result = cap.execute(
        None,
        output="dcf_model.xlsx",
        revenue=10000000,
        growth_rate=0.2,
        operating_margin=0.3,
        tax_rate=0.25,
        wacc=0.12,
        terminal_growth=0.03
    )
    print("DCF 模型已创建：")
    print(json.dumps(result, ensure_ascii=False, indent=2))


def example_lbo_model():
    """示例：创建 LBO 杠杆收购模型"""
    cap = LBOCapability()
    result = cap.execute(
        None,
        output="lbo_model.xlsx",
        ebitda=50000000,
        entry_multiple=8,
        exit_multiple=10,
        exit_year=5
    )
    print("LBO 模型已创建：")
    print(json.dumps(result, ensure_ascii=False, indent=2))


def example_variance_analysis():
    """示例：预算差异分析"""
    import pandas as pd

    # 创建测试数据
    budget = pd.DataFrame({
        "Item": ["Sales", "Marketing", "R&D"],
        "Budget": [1000000, 200000, 500000]
    })
    actual = pd.DataFrame({
        "Item": ["Sales", "Marketing", "R&D"],
        "Actual": [1200000, 180000, 550000]
    })

    # 保存到 Excel
    with pd.ExcelWriter("budget.xlsx") as writer:
        budget.to_excel(writer, sheet_name="Budget", index=False)
        actual.to_excel(writer, sheet_name="Actual", index=False)

    cap = VarianceCapability()
    result = cap.execute(
        None,
        file="budget.xlsx",
        budget_sheet="Budget",
        actual_sheet="Actual",
        output="variance_report.xlsx"
    )
    print("差异分析报告已创建：")
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    print("=== Abacus 金融建模示例 ===\n")

    print("1. DCF 估值模型:")
    # example_dcf_model()

    print("\n2. LBO 杠杆收购模型:")
    # example_lbo_model()

    print("\n3. 预算差异分析:")
    # example_variance_analysis()
