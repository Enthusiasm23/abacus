---
name: abacus-triangle
description: "勾股章 - 数据分析与可视化。当需要统计分析、趋势分析、相关性分析、数据可视化时使用。"
version: 0.2.0
chapter: triangle
level: rod
tags: [excel, analysis, statistics, trend, correlation, visualization, chart]
---

# 勾股章 - 数据分析与可视化

## CRITICAL RULES
1. 分析前先验证数据完整性 - 检查空值和异常值
2. 选择正确的分析方法 - 根据数据类型和目标选择
3. 结果必须包含统计指标 - 均值、中位数、标准差等
4. 可视化图表必须有标题和标签

## 能力一览

| 能力 | CLI 命令 | 说明 |
|------|----------|------|
| analyze_stats | `abacus analyze-stats` | 统计分析（均值、中位数、标准差） |
| analyze_trend | `abacus analyze-trend` | 趋势分析 |
| analyze_correlation | `abacus analyze-correlation` | 相关性分析 |
| analyze_data | `abacus analyze-data` | 智能数据分析 |
| visualize | `abacus visualize` | 数据可视化（生成图表） |
| visualize_data | `abacus visualize-data` | CSV 数据可视化 |
| variance_analysis | `abacus variance-analysis` | 预算与实际差异分析 |
| advanced_analysis | `abacus advanced-analysis` | 高级数据分析（回归、时间序列） |

---

## 统计分析参考

### analyze_stats 能力说明

analyze_stats 用于计算数据的统计指标，支持：
- 均值 (Mean)
- 中位数 (Median)
- 标准差 (Standard Deviation)
- 最小值 (Min)
- 最大值 (Max)
- 计数 (Count)

### 使用场景

```python
# 统计分析
abacus_analyze_stats(
    file='data.xlsx',
    sheet='Sheet1',
    range='A1:A100'
)
```

### 统计指标说明

| 指标 | 说明 | 用途 |
|------|------|------|
| mean | 均值 | 数据中心趋势 |
| median | 中位数 | 抗异常值的中心趋势 |
| std | 标准差 | 数据离散程度 |
| min | 最小值 | 数据下限 |
| max | 最大值 | 数据上限 |
| count | 计数 | 数据量 |

### 输出格式

```json
{
  "mean": 50.5,
  "median": 48.0,
  "std": 15.2,
  "min": 10,
  "max": 95,
  "count": 100
}
```

---

## 趋势分析参考

### analyze_trend 能力说明

analyze_trend 用于分析数据随时间的变化趋势，支持：
- 增长/下降趋势识别
- 时间序列分析
- 移动平均计算

### 使用场景

```python
# 趋势分析
abacus_analyze_trend(
    file='data.xlsx',
    sheet='Sheet1',
    range='A1:C100',
    value_column='Sales',
    time_column='Date'
)
```

### 趋势分析结果

```json
{
  "trend": "increasing",
  "growth_rate": 0.15,
  "moving_average": [45, 48, 52, 55, 58],
  "seasonality": false,
  "forecast": [60, 63, 66]
}
```

---

## 相关性分析参考

### analyze_correlation 能力说明

analyze_correlation 用于分析两个变量之间的相关性，支持：
- Pearson 相关系数
- 相关性强度判断
- 散点图数据

### 使用场景

```python
# 相关性分析
abacus_analyze_correlation(
    file='data.xlsx',
    sheet='Sheet1',
    range='A1:C100',
    column1='Sales',
    column2='Profit'
)
```

### 相关系数解释

| 系数值 | 相关性 |
|--------|--------|
| 0.8-1.0 | 强正相关 |
| 0.6-0.8 | 中等正相关 |
| 0.4-0.6 | 弱正相关 |
| 0.2-0.4 | 极弱正相关 |
| 0.0-0.2 | 无相关 |
| -0.2-0.0 | 极弱负相关 |
| -0.4--0.2 | 弱负相关 |
| -0.6--0.4 | 中等负相关 |
| -0.8--0.6 | 强负相关 |
| -1.0--0.8 | 强负相关 |

---

## 智能数据分析参考

### analyze_data 能力说明

analyze_data 用于智能数据分析，支持：
- 自动检测数据类型
- 生成统计摘要
- 发现数据间的相关性

### 使用场景

```python
# 智能分析
abacus_analyze_data(
    file='data.xlsx',
    sheet='Sheet1',
    analysis_type='auto'
)

# 摘要分析
abacus_analyze_data(
    file='data.xlsx',
    sheet='Sheet1',
    analysis_type='summary'
)

# 相关性分析
abacus_analyze_data(
    file='data.xlsx',
    sheet='Sheet1',
    analysis_type='correlation'
)
```

---

## 数据可视化参考

### visualize 能力说明

visualize 用于将 Excel/CSV 数据生成图片格式的图表，支持：
- 柱状图 (bar)
- 折线图 (line)
- 饼图 (pie)
- 散点图 (scatter)
- 热力图 (heatmap)

### 使用场景

```python
# 柱状图
abacus_visualize(
    file='data.xlsx',
    output='chart.png',
    chart_type='bar',
    x_column='Month',
    y_column='Sales',
    title='月度销售'
)

# 折线图
abacus_visualize(
    file='data.xlsx',
    output='trend.png',
    chart_type='line',
    x_column='Date',
    y_column='Value',
    title='趋势分析'
)

# 饼图
abacus_visualize(
    file='data.xlsx',
    output='pie.png',
    chart_type='pie',
    x_column='Category',
    y_column='Amount',
    title='占比分析'
)
```

### 图表类型选择

| 类型 | 使用场景 | 说明 |
|------|----------|------|
| bar | 比较类别 | 适合分类数据对比 |
| line | 趋势变化 | 适合时间序列数据 |
| pie | 占比构成 | 适合部分与整体关系 |
| scatter | 变量相关 | 适合两个变量关系 |
| heatmap | 密度分布 | 适合大量数据分布 |

---

## CSV 数据可视化参考

### visualize_data 能力说明

visualize_data 用于快速可视化 CSV 数据，支持：
- 自动生成图表
- 生成仪表板
- 生成统计摘要

### 使用场景

```python
# 快速可视化
abacus_visualize_data(
    file='data.csv',
    output='visualization.xlsx',
    chart_type='auto',
    include_dashboard=True,
    include_stats=True
)
```

---

## 预算与实际差异分析参考

### variance_analysis 能力说明

variance_analysis 用于预算与实际数据的差异分析，支持：
- 差异计算
- 差异原因分析
- 差异报告生成

### 使用场景

```python
# 差异分析
abacus_variance_analysis(
    file='data.xlsx',
    budget_sheet='Budget',
    actual_sheet='Actual',
    output='variance_report.xlsx',
    threshold=0.1
)
```

### 差异分析结果

```json
{
  "total_variance": 5000,
  "variance_percentage": 0.05,
  "significant_items": [
    {
      "item": "Marketing",
      "budget": 10000,
      "actual": 12000,
      "variance": 2000,
      "percentage": 0.20
    }
  ]
}
```

---

## 高级数据分析参考

### advanced_analysis 能力说明

advanced_analysis 用于高级数据分析，支持：
- 线性回归分析
- 时间序列分析
- 线性外推预测

### 使用场景

```python
# 回归分析
abacus_advanced_analysis(
    file='data.xlsx',
    sheet='Sheet1',
    analysis_type='regression',
    x_column='X',
    y_column='Y'
)

# 时间序列分析
abacus_advanced_analysis(
    file='data.xlsx',
    sheet='Sheet1',
    analysis_type='timeseries',
    y_column='Sales'
)

# 预测
abacus_advanced_analysis(
    file='data.xlsx',
    sheet='Sheet1',
    analysis_type='forecast',
    y_column='Sales',
    periods=10
)
```

### 回归分析结果

```json
{
  "slope": 2.5,
  "intercept": 10.0,
  "r_squared": 0.85,
  "equation": "y = 2.5x + 10.0",
  "prediction": [12.5, 15.0, 17.5]
}
```

---

## 场景驱动示例

### 财务报表场景

**利润表生成**
**用户说**：「帮我做一个利润表模板」

**处理步骤**：
1. 创建工作簿
2. 添加"利润表"工作表
3. 设置表头：项目、本期金额、上期金额、同比增长
4. 添加公式行：
   - 毛利润 = 营业收入 - 营业成本
   - 营业利润 = 毛利润 - 销售费用 - 管理费用 - 财务费用
   - 净利润 = 利润总额 - 所得税
5. 应用财务模型格式（蓝色=输入，黑色=公式）
6. 保存文件

### 人事管理场景

**考勤统计**
**用户说**：「统计这个月每个人的出勤天数」

**处理步骤**：
1. 识别关键词：统计 → COUNTIF
2. 生成公式：`=COUNTIF(B2:AF2,"√")`
3. 批量应用到所有行
4. 添加汇总行

### 销售分析场景

**业绩排名**
**用户说**：「按销售额从高到低排名」

**处理步骤**：
1. 识别关键词：排名 → RANK
2. 生成公式：`=RANK(B2,$B$2:$B$100,0)`
3. 批量应用
4. 条件格式：前10名高亮

**趋势分析**
**用户说**：「分析一下销售趋势」

**处理步骤**：
1. 识别关键词：趋势 → analyze_trend
2. 计算增长率、移动平均
3. 创建趋势图表
4. 输出分析报告

---

## 触发条件设计指南

### 设计原则

#### 1. 正面触发（Trigger When）
- 明确列出支持的文件格式
- 列出具体操作场景
- 包含隐式触发条件

#### 2. 负面排除（Do NOT Trigger When）
- 明确列出不触发的场景
- 基于"最终产出"判断
- 避免边界模糊

#### 3. 场景化描述
- 使用用户语言
- 包含具体示例
- 考虑边缘情况

### 常见触发词

#### 文件格式
- .xlsx, .xlsm, .xls, .csv, .tsv
- spreadsheet, workbook, worksheet

#### 操作动词
- create, make, generate
- open, read, load
- edit, modify, update
- fix, repair, debug
- convert, transform, export

#### 场景描述
- "帮我做一个报表"
- "把这个数据放到 Excel 里"
- "检查一下这个表格有没有问题"

---

## 中文关键词路由

### 分析场景

| 用户原话 | 关键词 | 推荐操作 |
|---------|--------|----------|
| 汇总/统计/分析 | 汇总/统计/分析 | pivot_table, summarize |
| 趋势/变化/增长 | 趋势/变化 | analyze_trend |
| 相关性/关联 | 相关/关联 | analyze_correlation |
| 图表/可视化/画图 | 图/画 | create_chart |

---

## 切片器概念

### 什么是切片器

切片器是 Excel 的可视化筛选器，用于：
- 交互式筛选数据
- 多个透视表/图表联动
- 直观的用户界面

### 切片器类型

#### 透视表切片器
- 连接到透视表字段
- 支持多选
- 支持清除筛选

#### 表格切片器
- 连接到 Excel 表格列
- 支持多选
- 支持搜索

### 工作流

#### 创建切片器
1. 选择数据源（透视表/表格）
2. 选择字段
3. 设置位置和大小
4. 格式化外观

#### 连接多个透视表
1. 创建主透视表
2. 创建切片器
3. 将切片器连接到其他透视表

### 限制

- openpyxl 不支持切片器操作
- 需要 Excel COM API
- 某些高级功能需要 Power BI

### 替代方案

对于 Python 环境，可以使用：
- 下拉列表替代切片器
- 数据验证实现筛选
- 自定义 Web 界面

---

## 最佳实践

1. **验证数据完整性** - 分析前检查空值和异常值
2. **选择正确方法** - 根据数据类型和目标选择
3. **包含统计指标** - 均值、中位数、标准差等
4. **图表有标题和标签** - 确保可视化清晰
5. **提供解释** - 不仅给出数字，还要解释含义

## 示例

```bash
# 统计分析
abacus analyze-stats -f data.xlsx -s Sheet1 -r A1:A100

# 趋势分析
abacus analyze-trend -f data.xlsx -s Sheet1 -r A1:C100 --value-column Sales --time-column Date

# 相关性分析
abacus analyze-correlation -f data.xlsx -s Sheet1 -r A1:C100 --column1 Sales --column2 Profit

# 可视化
abacus visualize -f data.xlsx --output chart.png --chart-type bar --x-column Month --y-column Sales --title "月度销售"

# 差异分析
abacus variance-analysis -f data.xlsx --budget-sheet Budget --actual-sheet Actual --output variance_report.xlsx

# 高级分析
abacus advanced-analysis -f data.xlsx -s Sheet1 --analysis-type regression --x-column X --y-column Y
```