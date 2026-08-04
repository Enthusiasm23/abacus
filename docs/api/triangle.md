# 勾股章 (Triangle) - 数据分析与可视化

> 勾股章负责统计分析、趋势分析、相关性分析、数据可视化和高级分析。

---

## 工具列表

| 工具 | 描述 |
|------|------|
| `analyze_stats` | 统计分析 |
| `analyze_trend` | 趋势分析 |
| `analyze_correlation` | 相关性分析 |
| `analyze_data` | 智能数据分析 |
| `visualize` | 数据可视化（生成 PNG/SVG/PDF 图表） |
| `visualize_data` | CSV 数据可视化 |
| `advanced_analysis` | 高级数据分析（回归、时间序列、预测） |
| `variance_analysis` | 预算差异分析 |

---

## analyze_stats

统计分析。

### 参数

| 参数 | 类型 | 必填 | 默认值 | 说明 |
|------|------|------|--------|------|
| `file` | string | 是 | - | Excel 文件路径 |
| `sheet` | string | 是 | - | 工作表名称 |
| `range` | string | 是 | - | 数据范围（A1 表示法） |

### 返回值

```python
{
    "mean": 5000.5,
    "median": 4500.0,
    "std": 2000.3,
    "min": 100,
    "max": 15000,
    "count": 100,
    "q1": 3000,
    "q3": 7000
}
```

### 示例

```python
analyze_stats(file="data.xlsx", sheet="Sheet1", range="A1:A100")
analyze_stats(file="data.xlsx", sheet="Sales", range="C1:C500")
```

---

## analyze_trend

趋势分析。

### 参数

| 参数 | 类型 | 必填 | 默认值 | 说明 |
|------|------|------|--------|------|
| `file` | string | 是 | - | Excel 文件路径 |
| `sheet` | string | 是 | - | 工作表名称 |
| `range` | string | 是 | - | 数据范围（A1 表示法） |
| `value_column` | string | 是 | - | 值列名 |
| `time_column` | string | 否 | None | 时间列名 |

### 返回值

```python
{
    "trend": "up",
    "slope": 150.5,
    "r_squared": 0.85,
    "forecast": [15000, 15150, 15300]
}
```

### 趋势方向

| 值 | 说明 |
|------|------|
| `up` | 上升趋势 |
| `down` | 下降趋势 |
| `stable` | 稳定趋势 |

### 示例

```python
analyze_trend(file="data.xlsx", sheet="Sheet1", range="A1:C100", value_column="Sales")
analyze_trend(file="data.xlsx", sheet="Sales", range="A1:D200", value_column="Revenue", time_column="Date")
```

---

## analyze_correlation

相关性分析。

### 参数

| 参数 | 类型 | 必填 | 默认值 | 说明 |
|------|------|------|--------|------|
| `file` | string | 是 | - | Excel 文件路径 |
| `sheet` | string | 是 | - | 工作表名称 |
| `range` | string | 是 | - | 数据范围（A1 表示法） |
| `column1` | string | 是 | - | 第一列名 |
| `column2` | string | 是 | - | 第二列名 |

### 返回值

```python
{
    "correlation": 0.85,
    "strength": "strong",
    "direction": "positive"
}
```

### 相关强度

| 系数范围 | 强度 |
|----------|------|
| 0.8 - 1.0 | strong（强） |
| 0.5 - 0.8 | moderate（中等） |
| 0.0 - 0.5 | weak（弱） |

### 示例

```python
analyze_correlation(file="data.xlsx", sheet="Sheet1", range="A1:C100", column1="Sales", column2="Profit")
analyze_correlation(file="data.xlsx", sheet="Data", range="B1:D200", column1="Price", column2="Quantity")
```

---

## analyze_data

智能数据分析（自动检测数据类型、统计摘要、相关性分析）。

### 参数

| 参数 | 类型 | 必填 | 默认值 | 说明 |
|------|------|------|--------|------|
| `file` | string | 是 | - | 文件路径（支持 Excel 和 CSV） |
| `sheet` | string | 否 | None | 工作表名称 |
| `analysis_type` | string | 否 | "auto" | 分析类型：`auto`/`summary`/`correlation` |

### 返回值

```python
{
    "data_types": {
        "numeric": ["Sales", "Profit", "Quantity"],
        "categorical": ["Category", "Region"],
        "datetime": ["Date"]
    },
    "statistics": {
        "Sales": {"mean": 5000, "std": 2000, "min": 100, "max": 15000},
        "Profit": {"mean": 1000, "std": 500, "min": 50, "max": 3000}
    },
    "correlations": {
        "Sales-Profit": 0.85,
        "Sales-Quantity": 0.72
    }
}
```

### 示例

```python
analyze_data(file="data.xlsx", sheet="Sheet1", analysis_type="auto")
analyze_data(file="data.csv", analysis_type="summary")
```

---

## visualize

数据可视化（生成 PNG/SVG/PDF 图表）。

### 参数

| 参数 | 类型 | 必填 | 默认值 | 说明 |
|------|------|------|--------|------|
| `file` | string | 是 | - | 数据文件路径 |
| `output` | string | 是 | - | 输出图片路径 |
| `chart_type` | string | 是 | - | 图表类型：`bar`/`line`/`pie`/`scatter`/`heatmap` |
| `x_column` | string | 否 | None | X轴列名 |
| `y_column` | string | 否 | None | Y轴列名 |
| `sheet` | string | 否 | None | 工作表名称 |
| `title` | string | 否 | None | 图表标题 |
| `width` | float | 否 | 10 | 图片宽度（英寸） |
| `height` | float | 否 | 6 | 图片高度（英寸） |

### 返回值

```python
{
    "file": "data.xlsx",
    "output": "chart.png",
    "chart_type": "bar",
    "created": True
}
```

### 示例

```python
visualize(file="data.xlsx", output="chart.png", chart_type="bar", title="销售趋势")
visualize(file="data.xlsx", output="scatter.png", chart_type="scatter", x_column="Price", y_column="Sales")
visualize(file="data.csv", output="pie.png", chart_type="pie", y_column="Category")
```

---

## visualize_data

CSV 数据可视化（自动生成图表、仪表板、统计摘要）。

### 参数

| 参数 | 类型 | 必填 | 默认值 | 说明 |
|------|------|------|--------|------|
| `file` | string | 是 | - | 文件路径（支持 CSV 和 Excel） |
| `output` | string | 是 | - | 输出文件路径 |
| `chart_type` | string | 否 | "auto" | 图表类型：`bar`/`line`/`pie`/`auto` |
| `include_dashboard` | bool | 否 | True | 是否包含仪表板 |
| `include_stats` | bool | 否 | True | 是否包含统计摘要 |

### 返回值

```python
{
    "success": True,
    "output": "visualization.xlsx",
    "charts_created": 5
}
```

### 示例

```python
visualize_data(file="data.csv", output="visualization.xlsx", chart_type="auto")
visualize_data(file="data.xlsx", output="report.xlsx", chart_type="bar", include_dashboard=True)
```

---

## advanced_analysis

高级数据分析（回归分析、时间序列、预测）。

### 参数

| 参数 | 类型 | 必填 | 默认值 | 说明 |
|------|------|------|--------|------|
| `file` | string | 是 | - | 文件路径 |
| `analysis_type` | string | 是 | - | 分析类型：`regression`/`timeseries`/`forecast` |
| `sheet` | string | 否 | None | 工作表名称 |
| `x_column` | string | 条件 | None | 自变量列名（regression 时必填） |
| `y_column` | string | 条件 | None | 因变量列名 |
| `periods` | int | 否 | 10 | 预测期数（forecast 时） |

### 返回值

```python
# regression
{
    "type": "regression",
    "regression": {
        "slope": 2.5,
        "intercept": 100,
        "r_squared": 0.92,
        "equation": "y = 2.5x + 100"
    }
}

# timeseries
{
    "type": "timeseries",
    "timeseries": {
        "mean": 5000,
        "std": 1500,
        "trend": "up",
        "seasonality": False
    }
}

# forecast
{
    "type": "forecast",
    "forecast": {
        "historical": [4500, 4800, 5000, 5200],
        "predicted": [5400, 5600, 5800, 6000, 6200]
    }
}
```

### 示例

```python
# 线性回归
advanced_analysis(file="data.xlsx", analysis_type="regression", x_column="Advertising", y_column="Sales")

# 时间序列
advanced_analysis(file="data.xlsx", analysis_type="timeseries", y_column="Revenue")

# 预测
advanced_analysis(file="data.xlsx", analysis_type="forecast", y_column="Sales", periods=12)
```

---

## variance_analysis

预算与实际差异分析。

### 参数

| 参数 | 类型 | 必填 | 默认值 | 说明 |
|------|------|------|--------|------|
| `file` | string | 是 | - | Excel 文件路径 |
| `budget_sheet` | string | 是 | - | 预算数据工作表 |
| `actual_sheet` | string | 是 | - | 实际数据工作表 |
| `output` | string | 否 | None | 输出文件路径 |
| `threshold` | float | 否 | 0.1 | 重要性阈值（10%） |

### 返回值

```python
{
    "success": True,
    "variances": [
        {"item": "Marketing", "budget": 50000, "actual": 55000, "variance": 0.10, "significant": True},
        {"item": "Rent", "budget": 30000, "actual": 30000, "variance": 0.0, "significant": False}
    ],
    "significant_variances": [
        {"item": "Marketing", "variance": 0.10, "direction": "over"}
    ]
}
```

### 示例

```python
variance_analysis(
    file="data.xlsx",
    budget_sheet="Budget",
    actual_sheet="Actual",
    threshold=0.1
)
variance_analysis(
    file="report.xlsx",
    budget_sheet="FY2024_Budget",
    actual_sheet="FY2024_Actual",
    output="variance_report.xlsx",
    threshold=0.05
)
```
