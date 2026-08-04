# 衰分章 (Share) - 分组汇总与数据分析

> 衰分章负责数据分组、比例分配、汇总统计和透视分析。

---

## 工具列表

| 工具 | 描述 |
|------|------|
| `group_by` | 按字段分组 |
| `distribute` | 按比例分配 |
| `summarize` | 分组汇总 |
| `pivot_analysis` | 数据透视分析 |
| `subtotal` | 分类汇总 |

---

## group_by

按字段分组。

### 参数

| 参数 | 类型 | 必填 | 默认值 | 说明 |
|------|------|------|--------|------|
| `file` | string | 是 | - | Excel 文件路径 |
| `sheet` | string | 是 | - | 工作表名称 |
| `range` | string | 是 | - | 数据范围（A1 表示法） |
| `group_columns` | list | 是 | - | 分组列名列表 |

### 返回值

```python
{
    "groups": {
        "East": [{"Name": "Product A", "Sales": 1000}, ...],
        "West": [{"Name": "Product B", "Sales": 2000}, ...]
    },
    "group_count": 2
}
```

### 示例

```python
group_by(file="data.xlsx", sheet="Sheet1", range="A1:D100", group_columns=["Region"])
group_by(file="data.xlsx", sheet="Sales", range="A1:E200", group_columns=["Category", "Region"])
```

---

## distribute

按比例分配。

### 参数

| 参数 | 类型 | 必填 | 默认值 | 说明 |
|------|------|------|--------|------|
| `file` | string | 是 | - | Excel 文件路径 |
| `sheet` | string | 是 | - | 工作表名称 |
| `range` | string | 是 | - | 数据范围（A1 表示法） |
| `total` | float | 是 | - | 待分配的总数 |
| `method` | string | 否 | "equal" | 分配方法：`equal`/`proportional` |

### 返回值

```python
{
    "allocations": [
        {"group": "A", "amount": 3333.33},
        {"group": "B", "amount": 3333.33},
        {"group": "C", "amount": 3333.34}
    ],
    "total_allocated": 10000
}
```

### 示例

```python
distribute(file="data.xlsx", sheet="Sheet1", range="A1:B10", total=10000, method="equal")
distribute(file="data.xlsx", sheet="Sheet1", range="A1:B10", total=10000, method="proportional")
```

---

## summarize

分组汇总。

### 参数

| 参数 | 类型 | 必填 | 默认值 | 说明 |
|------|------|------|--------|------|
| `file` | string | 是 | - | Excel 文件路径 |
| `sheet` | string | 是 | - | 工作表名称 |
| `range` | string | 是 | - | 数据范围（A1 表示法） |
| `group_by` | string | 是 | - | 分组列名 |
| `agg_config` | dict | 是 | - | 聚合配置 |

#### agg_config 示例

```python
{"Sales": "sum", "Profit": "mean", "Count": "count"}
```

### 聚合函数

| 函数 | 说明 |
|------|------|
| `sum` | 求和 |
| `mean` | 平均值 |
| `count` | 计数 |
| `min` | 最小值 |
| `max` | 最大值 |

### 返回值

```python
{
    "summary": {
        "East": {"Sales": 50000, "Profit": 10000},
        "West": {"Sales": 60000, "Profit": 12000}
    },
    "groups": 2
}
```

### 示例

```python
summarize(
    file="data.xlsx",
    sheet="Sales",
    range="A1:D100",
    group_by="Region",
    agg_config={"Sales": "sum", "Profit": "mean"}
)
```

---

## pivot_analysis

数据透视分析（分组汇总、交叉分析）。

### 参数

| 参数 | 类型 | 必填 | 默认值 | 说明 |
|------|------|------|--------|------|
| `file` | string | 是 | - | 文件路径 |
| `group_by` | string | 是 | - | 分组字段 |
| `value_field` | string | 是 | - | 值字段 |
| `sheet` | string | 否 | None | 工作表名称 |
| `agg_function` | string | 否 | "sum" | 聚合函数：`sum`/`mean`/`count`/`min`/`max` |
| `output` | string | 否 | None | 输出文件路径 |

### 返回值

```python
{
    "success": True,
    "pivot_table": {
        "Category A": 15000,
        "Category B": 25000,
        "Category C": 10000
    }
}
```

### 示例

```python
pivot_analysis(
    file="data.xlsx",
    group_by="Category",
    value_field="Sales",
    agg_function="sum"
)
pivot_analysis(
    file="data.xlsx",
    group_by="Region",
    value_field="Profit",
    agg_function="mean",
    output="pivot_result.xlsx"
)
```

---

## subtotal

分类汇总（按字段分组聚合）。

### 参数

| 参数 | 类型 | 必填 | 默认值 | 说明 |
|------|------|------|--------|------|
| `file` | string | 是 | - | Excel 文件路径 |
| `sheet` | string | 是 | - | 工作表名称 |
| `range` | string | 是 | - | 数据范围（A1 表示法） |
| `group_column` | string | 是 | - | 分组列名 |
| `function` | string | 否 | "sum" | 聚合函数：`sum`/`mean`/`count`/`min`/`max` |

### 返回值

```python
{
    "success": True,
    "group_column": "Category",
    "function": "sum",
    "groups_count": 5,
    "summary": {
        "Electronics": 150000,
        "Clothing": 80000,
        "Food": 60000
    }
}
```

### 示例

```python
subtotal(file="data.xlsx", sheet="Sheet1", range="A1:D100", group_column="Category", function="sum")
subtotal(file="data.xlsx", sheet="Sheet1", range="A1:D100", group_column="Department", function="mean")
```
