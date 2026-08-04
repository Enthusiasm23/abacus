# 方田章 (Field) - 数据读取与单元格操作

> 方田章负责 Excel 文件的数据读取、结构查看、单元格操作和工作表管理。

---

## 工具列表

| 工具 | 描述 |
|------|------|
| `measure_range` | 读取指定范围数据 |
| `measure_cells` | 读取单元格详细信息（值、公式、样式） |
| `measure_structure` | 读取工作表结构 |
| `list_sheets` | 返回工作表名称列表 |
| `peek_preview` | 快速预览前几行数据 |
| `detect_columns` | 检测列名和数据类型 |
| `search_content` | 搜索关键词 |
| `get_summary` | 获取文件摘要信息 |
| `get_sample_data` | 获取样本数据 |
| `manage_named_range` | 管理命名范围 |
| `manage_sheet_style` | 管理工作表样式 |
| `manage_sheet_visibility` | 管理工作表可见性 |

---

## measure_range

读取指定范围数据。

### 参数

| 参数 | 类型 | 必填 | 默认值 | 说明 |
|------|------|------|--------|------|
| `file` | string | 是 | - | Excel 文件路径 |
| `sheet` | string | 是 | - | 工作表名称 |
| `range` | string | 是 | - | 数据范围（A1 表示法） |

### 返回值

```python
{
    "range": "A1:D10",       # 实际读取的范围
    "sheet": "Sheet1",       # 工作表名称
    "data": [["Name", "Value"], ...],  # 二维数组数据
    "rows": 10,              # 行数
    "columns": 4             # 列数
}
```

### 示例

```python
measure_range(file="data.xlsx", sheet="Sales", range="A1:C10")
```

---

## measure_cells

读取单元格详细信息（值、公式、样式）。

### 参数

| 参数 | 类型 | 必填 | 默认值 | 说明 |
|------|------|------|--------|------|
| `file` | string | 是 | - | Excel 文件路径 |
| `sheet` | string | 是 | - | 工作表名称 |
| `range` | string | 是 | - | 数据范围（A1 表示法） |

### 返回值

```python
{
    "range": "A1:B5",
    "sheet": "Sheet1",
    "cells": [
        {
            "address": "A1",
            "value": "Name",
            "formula": None,
            "style": {...},
            "data_type": "string"
        },
        ...
    ],
    "count": 10
}
```

### 示例

```python
measure_cells(file="data.xlsx", sheet="Sheet1", range="A1:C5")
```

---

## measure_structure

读取工作表结构（行数、列数、合并单元格等）。

### 参数

| 参数 | 类型 | 必填 | 默认值 | 说明 |
|------|------|------|--------|------|
| `file` | string | 是 | - | Excel 文件路径 |
| `sheet` | string | 否 | None | 工作表名称（不填则返回所有工作表） |

### 返回值

```python
{
    "sheets": [
        {
            "name": "Sheet1",
            "max_row": 100,
            "max_column": 10,
            "merged_cells": []
        }
    ],
    "dimensions": {...},
    "merged_cells": []
}
```

### 示例

```python
measure_structure(file="data.xlsx")
measure_structure(file="data.xlsx", sheet="Sheet1")
```

---

## list_sheets

返回 Excel 文件中所有工作表名称列表。

### 参数

| 参数 | 类型 | 必填 | 默认值 | 说明 |
|------|------|------|--------|------|
| `file` | string | 是 | - | Excel 文件路径 |

### 返回值

```python
{
    "file": "data.xlsx",
    "sheets": ["Sheet1", "Sales", "Config"],
    "count": 3
}
```

### 示例

```python
list_sheets(file="data.xlsx")
```

---

## peek_preview

快速预览每个工作表的前几行数据。

### 参数

| 参数 | 类型 | 必填 | 默认值 | 说明 |
|------|------|------|--------|------|
| `file` | string | 是 | - | Excel 文件路径 |
| `rows` | int | 否 | 5 | 预览行数 |
| `sheet` | string | 否 | None | 工作表名称（不填则预览所有） |

### 返回值

```python
{
    "preview": {
        "Sheet1": {
            "headers": ["Name", "Value"],
            "rows": [["Product A", 100], ...]
        }
    }
}
```

### 示例

```python
peek_preview(file="data.xlsx", rows=3)
peek_preview(file="data.xlsx", sheet="Sheet1", rows=10)
```

---

## detect_columns

检测列名和数据类型。

### 参数

| 参数 | 类型 | 必填 | 默认值 | 说明 |
|------|------|------|--------|------|
| `file` | string | 是 | - | Excel 文件路径 |
| `sheet` | string | 是 | - | 工作表名称 |
| `sample_rows` | int | 否 | 100 | 采样行数 |

### 返回值

```python
{
    "columns": ["ID", "Name", "Sales", "Date"],
    "column_details": [
        {"name": "ID", "type": "int", "sample_count": 100},
        {"name": "Name", "type": "string", "sample_count": 100},
        ...
    ]
}
```

### 示例

```python
detect_columns(file="data.xlsx", sheet="Sheet1")
detect_columns(file="data.xlsx", sheet="Sheet1", sample_rows=500)
```

---

## search_content

在 Excel 文件中搜索关键词。

### 参数

| 参数 | 类型 | 必填 | 默认值 | 说明 |
|------|------|------|--------|------|
| `file` | string | 是 | - | Excel 文件路径 |
| `keyword` | string | 是 | - | 搜索关键词 |
| `sheet` | string | 否 | None | 工作表名称（不填则搜索所有） |
| `max_results` | int | 否 | 50 | 最大结果数 |

### 返回值

```python
{
    "results": [
        {"sheet": "Sheet1", "cell": "A1", "value": "销售"},
        {"sheet": "Sheet1", "cell": "C5", "value": "销售额"},
        ...
    ],
    "total_found": 15
}
```

### 示例

```python
search_content(file="data.xlsx", keyword="销售")
search_content(file="data.xlsx", keyword="100", sheet="Sheet1", max_results=10)
```

---

## get_summary

获取 Excel 文件摘要信息。

### 参数

| 参数 | 类型 | 必填 | 默认值 | 说明 |
|------|------|------|--------|------|
| `file` | string | 是 | - | Excel 文件路径 |

### 返回值

```python
{
    "sheet_count": 3,
    "total_rows": 1500,
    "sheets": [
        {"name": "Sheet1", "rows": 500, "columns": 10},
        {"name": "Sales", "rows": 800, "columns": 8},
        {"name": "Config", "rows": 200, "columns": 3}
    ]
}
```

### 示例

```python
get_summary(file="data.xlsx")
```

---

## get_sample_data

获取指定工作表的样本数据。

### 参数

| 参数 | 类型 | 必填 | 默认值 | 说明 |
|------|------|------|--------|------|
| `file` | string | 是 | - | Excel 文件路径 |
| `sheet` | string | 是 | - | 工作表名称 |
| `rows` | int | 否 | 10 | 样本行数 |

### 返回值

```python
{
    "columns": ["ID", "Name", "Sales"],
    "data": [
        [1, "Product A", 1000],
        [2, "Product B", 2000],
        ...
    ]
}
```

### 示例

```python
get_sample_data(file="data.xlsx", sheet="Sheet1", rows=5)
```

---

## manage_named_range

管理命名范围（创建/列出/读取/删除）。

### 参数

| 参数 | 类型 | 必填 | 默认值 | 说明 |
|------|------|------|--------|------|
| `file` | string | 是 | - | Excel 文件路径 |
| `action` | string | 是 | - | 操作：`create`/`list`/`read`/`delete` |
| `name` | string | 条件 | None | 命名范围名称（create/read/delete 时必填） |
| `refers_to` | string | 条件 | None | 引用位置（create 时必填） |

### 返回值

```python
# action=list
{
    "success": True,
    "named_ranges": [
        {"name": "SalesData", "refers_to": "Sheet1!$A$1:$D$10"}
    ]
}

# action=create/read/delete
{
    "success": True
}
```

### 示例

```python
manage_named_range(file="data.xlsx", action="create", name="SalesData", refers_to="Sheet1!$A$1:$D$10")
manage_named_range(file="data.xlsx", action="list")
manage_named_range(file="data.xlsx", action="read", name="SalesData")
manage_named_range(file="data.xlsx", action="delete", name="SalesData")
```

---

## manage_sheet_style

管理工作表样式（标签颜色）。

### 参数

| 参数 | 类型 | 必填 | 默认值 | 说明 |
|------|------|------|--------|------|
| `file` | string | 是 | - | Excel 文件路径 |
| `sheet` | string | 是 | - | 工作表名称 |
| `action` | string | 是 | - | 操作：`set`/`get`/`clear` |
| `color` | string | 条件 | None | 十六进制颜色代码（set 时必填） |

### 返回值

```python
{
    "success": True,
    "color": "FF0000"  # get 时返回
}
```

### 示例

```python
manage_sheet_style(file="data.xlsx", sheet="Sheet1", action="set", color="FF0000")
manage_sheet_style(file="data.xlsx", sheet="Sheet1", action="get")
manage_sheet_style(file="data.xlsx", sheet="Sheet1", action="clear")
```

---

## manage_sheet_visibility

管理工作表可见性（显示/隐藏/非常隐藏）。

### 参数

| 参数 | 类型 | 必填 | 默认值 | 说明 |
|------|------|------|--------|------|
| `file` | string | 是 | - | Excel 文件路径 |
| `sheet` | string | 是 | - | 工作表名称 |
| `action` | string | 是 | - | 操作：`show`/`hide`/`very-hide`/`get` |

### 返回值

```python
{
    "success": True,
    "visibility": "visible"  # get 时返回
}
```

### 示例

```python
manage_sheet_visibility(file="data.xlsx", sheet="Sheet1", action="hide")
manage_sheet_visibility(file="data.xlsx", sheet="Sheet1", action="show")
manage_sheet_visibility(file="data.xlsx", sheet="Sheet1", action="get")
```
