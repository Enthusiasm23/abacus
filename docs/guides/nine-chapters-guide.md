# 九章使用指南

## 概述

Abacus 采用中国《九章算术》的分类体系，将 Excel 操作分为九个章节。

## 方田章 - 数据读取

**核心功能：** 读取、查看、管理 Excel 数据

### 常用操作

```python
# 读取范围数据
from abacus.core import MeasureRangeCapability
cap = MeasureRangeCapability()
result = cap.execute(None, file="data.xlsx", sheet="Sheet1", range="A1:D10")

# 查看工作表结构
from abacus.core import MeasureStructureCapability
cap = MeasureStructureCapability()
result = cap.execute(None, file="data.xlsx")
```

## 粟米章 - 格式转换

**核心功能：** 转换数据格式、单位、类型

### 常用操作

```python
# 转换数字格式
from abacus.core import ConvertFormatCapability
cap = ConvertFormatCapability()
result = cap.execute(None, file="data.xlsx", sheet="Sheet1", range="B2:B100", format_type="number")

# 转换数据类型
from abacus.core import ConvertTypeCapability
cap = ConvertTypeCapability()
result = cap.execute(None, file="data.xlsx", sheet="Sheet1", range="A2:A100", target_type="text")
```

## 衰分章 - 分组汇总

**核心功能：** 按字段分组、按比例分配、汇总统计

### 常用操作

```python
# 按字段分组
from abacus.core import GroupByCapability
cap = GroupByCapability()
result = cap.execute(None, file="data.xlsx", sheet="Sales", range="A1:C100", 
                    group_columns=["Region"], value_field="Sales", agg_function="sum")
```

## 商功章 - 批量操作

**核心功能：** 批量处理、表格管理、图表、透视表

### 常用操作

```python
# 创建透视表
from abacus.core import CreatePivotCapability
cap = CreatePivotCapability()
result = cap.execute(None, file="data.xlsx", sheet="Sales", range="A1:C100",
                    row_fields=["Region"], value_field="Sales", agg_function="sum")

# 创建图表
from abacus.core import CreateChartCapability
cap = CreateChartCapability()
result = cap.execute(None, file="data.xlsx", sheet="Sales", range="A1:B10",
                    chart_type="bar", title="Sales Chart")
```

## 方程章 - 公式计算

**核心功能：** 创建公式、解方程、执行计算

### 常用操作

```python
# 创建公式
from abacus.core import CreateFormulaCapability
cap = CreateFormulaCapability()
result = cap.execute(None, file="data.xlsx", sheet="Sheet1", cell="E1", formula="SUM(A1:D1)")

# 生成常用公式
from abacus.core import FormulaGeneratorCapability
cap = FormulaGeneratorCapability()
result = cap.execute(None, formula_type="vlookup", 
                    params={"lookup_value": "D2", "table_range": "A:B", "col_index": 2})
```

## 勾股章 - 数据分析

**核心功能：** 统计分析、趋势分析、相关性分析

### 常用操作

```python
# 统计分析
from abacus.core import AnalyzeStatsCapability
cap = AnalyzeStatsCapability()
result = cap.execute(None, file="data.xlsx", sheet="Sheet1", range="A1:D100")

# 趋势分析
from abacus.core import AnalyzeTrendCapability
cap = AnalyzeTrendCapability()
result = cap.execute(None, file="data.xlsx", sheet="Sheet1", range="A1:B100", value_column="Sales")
```
