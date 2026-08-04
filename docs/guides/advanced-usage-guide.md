# 高级使用指南

## 批量操作

### 批量执行多个操作
```python
from abacus.core import BatchExecuteCapability

cap = BatchExecuteCapability()
result = cap.execute(
    None,
    file="data.xlsx",
    operations=[
        {"type": "merge", "sheet": "Sheet1", "range": "A1:D1"},
        {"type": "write", "sheet": "Sheet1", "cell": "A1", "value": "Title"},
        {"type": "style", "sheet": "Sheet1", "cell": "A1", "font": {"bold": True}}
    ]
)
```

### 批量转换数据
```python
from abacus.core import BatchTransformCapability

cap = BatchTransformCapability()
result = cap.execute(
    None,
    file="data.xlsx",
    operations=[
        {"type": "replace", "sheet": "Sheet1", "old": "Widget", "new": "Item"},
        {"type": "fill_formula", "sheet": "Sheet1", "cell": "E1", "formula": "SUM(B:B)"}
    ]
)
```

## 数据验证

### 验证数据范围
```python
from abacus.core import ValidateRangeCapability

cap = ValidateRangeCapability()
result = cap.execute(
    None,
    file="data.xlsx",
    sheet="Sheet1",
    range="B2:B100",
    rules={"no_empty": True, "min_value": 0, "max_value": 1000}
)
```

### 验证公式
```python
from abacus.core import ValidateFormulaCapability

cap = ValidateFormulaCapability()
result = cap.execute(None, file="data.xlsx")
```

## 数据分析

### 统计分析
```python
from abacus.core import AnalyzeStatsCapability

cap = AnalyzeStatsCapability()
result = cap.execute(
    None,
    file="data.xlsx",
    sheet="Sheet1",
    range="A1:D100"
)
```

### 趋势分析
```python
from abacus.core import AnalyzeTrendCapability

cap = AnalyzeTrendCapability()
result = cap.execute(
    None,
    file="data.xlsx",
    sheet="Sheet1",
    range="A1:B100",
    value_column="Sales"
)
```

### 相关性分析
```python
from abacus.core import AnalyzeCorrelationCapability

cap = AnalyzeCorrelationCapability()
result = cap.execute(
    None,
    file="data.xlsx",
    sheet="Sheet1",
    range="B1:C100",
    column1="Sales",
    column2="Profit"
)
```
