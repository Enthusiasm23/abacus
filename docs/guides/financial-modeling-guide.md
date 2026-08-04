# 金融建模指南

## 预算差异分析

### 概述

差异分析用于比较预算与实际数据。

### 使用方法

```python
from abacus.core import VarianceCapability

cap = VarianceCapability()
result = cap.execute(
    None,
    file="budget.xlsx",
    budget_sheet="Budget",
    actual_sheet="Actual",
    output="variance_report.xlsx"
)
```
