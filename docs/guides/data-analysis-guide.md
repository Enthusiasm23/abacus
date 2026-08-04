# 数据分析指南

## 统计分析

### 概述

统计分析用于了解数据的基本特征。

### 使用方法

```python
from abacus.core import AnalyzeStatsCapability

cap = AnalyzeStatsCapability()
result = cap.execute(None, file="data.xlsx", sheet="Sheet1", range="A1:D100")
```

### 返回结果

- count: 数据数量
- mean: 平均值
- std: 标准差
- min: 最小值
- max: 最大值
- q1: 第一四分位数
- q3: 第三四分位数

## 趋势分析

### 概述

趋势分析用于识别数据的变化趋势。

### 使用方法

```python
from abacus.core import AnalyzeTrendCapability

cap = AnalyzeTrendCapability()
result = cap.execute(None, file="data.xlsx", sheet="Sheet1", range="A1:B100", value_column="Sales")
```

### 返回结果

- trend_direction: 趋势方向（上升/下降/平稳）
- slope: 斜率
- period_changes: 各期变化

## 相关性分析

### 概述

相关性分析用于识别变量之间的关系。

### 使用方法

```python
from abacus.core import AnalyzeCorrelationCapability

cap = AnalyzeCorrelationCapability()
result = cap.execute(None, file="data.xlsx", sheet="Sheet1", range="B1:C100", 
                    column1="Sales", column2="Profit")
```

### 返回结果

- correlation: 相关系数
- strength: 相关强度（强/中/弱）
- direction: 相关方向（正/负）
