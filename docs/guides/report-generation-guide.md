# 报表生成指南

## 基础报表

### 概述

基础报表用于快速生成格式化的 Excel 文件。

### 使用方法

```python
from abacus.core import BasicReportCapability

cap = BasicReportCapability()
result = cap.execute(
    None,
    data_source="data.csv",
    output="report.xlsx",
    title="销售报表",
    sheet_name="Sales"
)
```

## 高级报表

### 概述

高级报表包含图表、条件格式、仪表板。

### 使用方法

```python
from abacus.core import AdvancedReportCapability

cap = AdvancedReportCapability()
result = cap.execute(
    None,
    data_source="data.csv",
    output="advanced_report.xlsx",
    chart_type="bar",
    include_dashboard=True
)
```

## 模板报表

### 概述

模板报表用于基于模板填充数据。

### 使用方法

```python
from abacus.core import TemplateReportCapability

cap = TemplateReportCapability()
result = cap.execute(
    None,
    template="template.xlsx",
    output="filled_report.xlsx",
    data={"A2": "2026-Q1", "B2": 1000000}
)
```
