---
name: abacus-share
description: "衰分章 - 分组汇总与数据分析。当需要按字段分组、按比例分配、汇总统计、透视分析时使用。"
version: 0.2.0
chapter: share
level: rod
tags: [excel, group, summarize, pivot, dashboard, analysis]
---

# 衰分章 - 分组汇总与数据分析

## CRITICAL RULES
1. 透视表不会自动刷新 - 数据变更后需重建
2. 使用 Excel 表格作为透视表源（自动扩展）
3. 字段名必须与源数据列名完全匹配
4. 多图表布局使用 targetRange 避免重叠

## 能力一览

| 能力 | CLI 命令 | 说明 |
|------|----------|------|
| group_by | `abacus group-by` | 按字段分组 |
| distribute | `abacus distribute` | 按比例分配 |
| summarize | `abacus summarize` | 分组汇总统计 |
| pivot_analysis | `abacus pivot-analysis` | 透视分析 |
| pivot_wizard | `abacus pivot-wizard` | 透视表向导 |
| subtotal | `abacus subtotal` | 分类汇总 |
| create_pivot | `abacus create-pivot` | 创建数据透视表 |

---

## 透视表参考

### 计算字段

| 特性 | 计算字段 | DAX 度量 |
|------|----------|----------|
| 单表公式 | 可用 (`=Qty*Price`) | 可用 |
| 跨表 | 不支持 | 完全支持 |
| 复杂逻辑 | 有限 | 完整 DAX |
| 可复用 | 仅限当前透视表 | 所有透视表可用 |

### 透视表数据源

| 数据源 | 创建操作 | 支持计算字段？ |
|--------|----------|----------------|
| 工作表范围 | `abacus_create_pivot` | 是 - 简单公式 |
| Excel 表格 | `abacus_create_pivot` | 是 - 结构化引用 |

### 字段配置

创建透视表时配置字段：

1. **行字段**：分组的类别
2. **值字段**：要聚合的数值数据
3. **聚合函数**：sum, avg, count, min, max

### 聚合函数选择

| 函数 | 使用场景 |
|------|----------|
| sum | 总计（收入、数量） |
| count | 记录数 |
| avg | 平均值 |
| min/max | 极值 |

### 透视表示例

```python
# 收入分析
abacus_create_pivot(
    file='data.xlsx',
    sheet='Sales',
    range='A1:D100',
    row_fields=['Region', 'Product'],
    value_field='Revenue',
    agg_function='sum'
)

# 交叉表
abacus_create_pivot(
    file='data.xlsx',
    sheet='Sales',
    range='A1:D100',
    row_fields=['Region'],
    value_field='Amount',
    agg_function='sum'
)

# 快速分析（不创建新工作表）
abacus_pivot_analysis(
    file='data.xlsx',
    group_by='Category',
    value_field='Sales',
    agg_function='sum',
    output='result.xlsx'
)
```

### 透视表常见错误

| 错误 | 原因 | 解决方案 |
|------|------|----------|
| "Field not found" | 字段名拼写错误 | 用 `abacus_measure_range` 检查列标题 |
| 数据不更新 | 源数据变更后未刷新 | 重建透视表 |
| 聚合错误 | agg_function 错误 | 指定正确的函数 (sum/avg/count) |
| 结果为空 | 范围包含空行 | 使用精确的 A1 表示法范围 |

---

## 仪表板最佳实践

### 专业报告工作流

```
1. 结构化数据 -> Excel 表格（永远不要用普通范围）
2. 格式化值 -> 按数据类型的数字格式
3. 添加视觉效果 -> 明确定位的图表
4. 验证布局 -> 截图确认无重叠
5. 保存并关闭 -> 持久化更改
```

### 常见仪表板布局

**摘要仪表板**
```
+------------------+
|   摘要表格       |
+------------------+
| 图表1  | 图表2   |
+------------------+
```

**分析仪表板**
```
+------------------+
| 图表1  | 图表2   |
+------------------+
| 图表3  | 图表4   |
+------------------+
```

**执行报告**
```
+------------------+
|   KPI 指标       |
+------------------+
|   趋势图表       |
+------------------+
|   详细数据       |
+------------------+
```

### 颜色方案

**专业配色**
- 主色：#2F5496（深蓝）
- 辅助色：#4472C4（蓝）
- 强调色：#ED7D31（橙）
- 背景色：F2F2F2（浅灰）

**语义颜色**
- 正面：#00B050（绿）
- 负面：#FF0000（红）
- 警告：#FFC000（黄）
- 中性：#A5A5A5（灰）

### 字体规范

- 标题：14-16pt，粗体
- 副标题：12-14pt，半粗体
- 正文：10-12pt，常规
- 数据标签：8-10pt，常规

### 数字格式速查

| 数据类型 | 格式代码 | 结果 |
|----------|----------|------|
| 货币 (USD) | `$#,##0.00` | $1,234.56 |
| 货币 (EUR) | `€#,##0.00` | €1,234.56 |
| 百分比 | `0.0%` | 12.3% |
| 日期 | `yyyy-mm-dd` | 2025-01-22 |
| 数字（千位） | `#,##0` | 1,235 |

### 图表定位规则

- 多图表布局使用 targetRange
- 图表间留 1-2 行/列间距
- 图表放在数据区域下方
- 保持图表大小一致

### 格式化清单

- 数据在 Excel 表格中（非普通范围）
- 已应用数字格式（货币、日期、百分比）
- 列宽适合内容
- 图表标题描述性强
- 图表轴标签已格式化
- 无图表与数据或其他图表重叠
- 仪表板中图表大小一致
- 已截图验证最终布局

---

## 最佳实践

1. **使用 Excel 表格作为数据源**：自动扩展范围简化刷新
2. **字段命名清晰**：使用描述性列标题
3. **选择正确的聚合**：金额用 sum，记录用 count，比率用 avg
4. **验证数据类型**：确保数值列是数字，不是文本
5. **先用小数据测试**：验证布局后再处理大数据集

## 示例

```bash
# 按字段分组
abacus group-by -f data.xlsx -s Sheet1 -r A1:D100 --group-columns Category Region

# 分组汇总
abacus summarize -f data.xlsx -s Sheet1 -r A1:D100 --group-by Category --agg-config {"Sales": "sum", "Profit": "mean"}

# 分类汇总
abacus subtotal -f data.xlsx -s Sheet1 -r A1:D100 --group-column Category --function sum

# 透视分析
abacus pivot-analysis -f data.xlsx --group-by Region --value-field Sales --agg-function sum

# 创建透视表
abacus create-pivot -f data.xlsx -s Sheet1 -r A1:D100 --row-fields Region Product --value-field Revenue --agg-function sum
```
