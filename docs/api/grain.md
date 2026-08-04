# 粟米章 (Grain) - 格式转换与数据处理

> 粟米章负责数据格式转换、类型转换、单位转换、数据清洗和高级数据转换。

---

## 工具列表

| 工具 | 描述 |
|------|------|
| `convert_format` | 转换数据格式（日期、数字、文本等） |
| `convert_type` | 转换数据类型（int/float/str/date） |
| `convert_unit` | 转换单位 |
| `transpose` | 转置数据（行列互换） |
| `text_to_columns` | 文本分列（按分隔符拆分） |
| `clean_data` | 数据清洗（去重、缺失值处理） |
| `transform_data` | 高级数据转换（透视、逆透视、合并、重塑） |
| `fuzzy_match_columns` | 模糊匹配列名 |
| `auto_type_infer` | 自动类型推断 |
| `standardize_data` | 数据标准化 |
| `transform_pipeline` | 数据转换管道 |

---

## convert_format

转换数据格式（日期、数字、文本等）。

### 参数

| 参数 | 类型 | 必填 | 默认值 | 说明 |
|------|------|------|--------|------|
| `file` | string | 是 | - | Excel 文件路径 |
| `sheet` | string | 是 | - | 工作表名称 |
| `range` | string | 是 | - | 数据范围（A1 表示法） |
| `format_type` | string | 是 | - | 目标格式：`date`/`number`/`text`/`percentage`/`currency` |

### 返回值

```python
{
    "success": True,
    "converted_count": 50
}
```

### 示例

```python
convert_format(file="data.xlsx", sheet="Sheet1", range="B2:B100", format_type="number")
convert_format(file="data.xlsx", sheet="Sheet1", range="A1:A50", format_type="date")
```

---

## convert_type

转换数据类型。

### 参数

| 参数 | 类型 | 必填 | 默认值 | 说明 |
|------|------|------|--------|------|
| `file` | string | 是 | - | Excel 文件路径 |
| `sheet` | string | 是 | - | 工作表名称 |
| `range` | string | 是 | - | 数据范围（A1 表示法） |
| `target_type` | string | 是 | - | 目标类型：`int`/`float`/`str`/`date` |

### 返回值

```python
{
    "success": True,
    "converted_count": 100
}
```

### 示例

```python
convert_type(file="data.xlsx", sheet="Sheet1", range="A2:A100", target_type="float")
convert_type(file="data.xlsx", sheet="Sheet1", range="B1:B50", target_type="str")
```

---

## convert_unit

转换单位。

### 参数

| 参数 | 类型 | 必填 | 默认值 | 说明 |
|------|------|------|--------|------|
| `file` | string | 是 | - | Excel 文件路径 |
| `sheet` | string | 是 | - | 工作表名称 |
| `range` | string | 是 | - | 数据范围（A1 表示法） |
| `from_unit` | string | 是 | - | 源单位 |
| `to_unit` | string | 是 | - | 目标单位 |

### 返回值

```python
{
    "success": True,
    "converted_count": 80
}
```

### 示例

```python
convert_unit(file="data.xlsx", sheet="Sheet1", range="B2:B100", from_unit="km", to_unit="m")
convert_unit(file="data.xlsx", sheet="Sheet1", range="C1:C50", from_unit="kg", to_unit="g")
```

---

## transpose

转置数据（行列互换）。

### 参数

| 参数 | 类型 | 必填 | 默认值 | 说明 |
|------|------|------|--------|------|
| `file` | string | 是 | - | Excel 文件路径 |
| `sheet` | string | 是 | - | 工作表名称 |
| `range` | string | 是 | - | 数据范围（A1 表示法） |
| `output_sheet` | string | 否 | None | 输出工作表名称（默认覆盖原工作表） |

### 返回值

```python
{
    "success": True,
    "source_range": "A1:D10",
    "source_rows": 10,
    "source_columns": 4,
    "output_sheet": "Transposed",
    "output_rows": 4,
    "output_columns": 10
}
```

### 示例

```python
transpose(file="data.xlsx", sheet="Sheet1", range="A1:D10", output_sheet="Transposed")
```

---

## text_to_columns

文本分列（按分隔符拆分）。

### 参数

| 参数 | 类型 | 必填 | 默认值 | 说明 |
|------|------|------|--------|------|
| `file` | string | 是 | - | Excel 文件路径 |
| `sheet` | string | 是 | - | 工作表名称 |
| `column` | string | 是 | - | 列标识（如 "A"） |
| `delimiter` | string | 否 | "," | 分隔符 |

### 返回值

```python
{
    "success": True,
    "column": "A",
    "delimiter": ",",
    "rows_split": 100,
    "columns_created": 3
}
```

### 示例

```python
text_to_columns(file="data.xlsx", sheet="Sheet1", column="A", delimiter=",")
text_to_columns(file="data.xlsx", sheet="Sheet1", column="B", delimiter="|")
```

---

## clean_data

数据清洗（去重、缺失值处理、格式化）。

### 参数

| 参数 | 类型 | 必填 | 默认值 | 说明 |
|------|------|------|--------|------|
| `file` | string | 是 | - | Excel 文件路径 |
| `sheet` | string | 否 | None | 工作表名称 |
| `output` | string | 否 | None | 输出文件路径（默认覆盖原文件） |
| `operations` | list | 否 | None | 清洗操作列表 |

#### operations 可选值

| 值 | 说明 |
|------|------|
| `remove_duplicates` | 去重 |
| `handle_missing` | 处理缺失值 |
| `strip_whitespace` | 去除空白 |
| `convert_types` | 类型转换 |

### 返回值

```python
{
    "success": True,
    "rows_affected": 25,
    "operations_performed": ["remove_duplicates", "handle_missing"]
}
```

### 示例

```python
clean_data(file="data.xlsx", operations=["remove_duplicates", "handle_missing"])
clean_data(file="data.xlsx", sheet="Sheet1", output="cleaned.xlsx")
```

---

## transform_data

高级数据转换（透视、转置、合并、重塑）。

### 参数

| 参数 | 类型 | 必填 | 默认值 | 说明 |
|------|------|------|--------|------|
| `file` | string | 是 | - | 文件路径 |
| `transform_type` | string | 是 | - | 转换类型：`pivot`/`melt`/`merge`/`reshape` |
| `sheet` | string | 否 | None | 工作表名称 |
| `params` | dict | 否 | None | 转换参数 |
| `output` | string | 否 | None | 输出文件路径 |

#### params 示例

```python
# pivot（透视）
{"index": "Category", "values": "Sales", "aggfunc": "sum"}

# melt（逆透视）
{"id_vars": ["ID"], "value_vars": ["A", "B"]}

# merge（合并）
{"other_file": "other.xlsx", "on": "ID"}

# reshape（重塑）
{"pivot_column": "Type", "value_column": "Amount"}
```

### 返回值

```python
{
    "transform_type": "pivot",
    "input_rows": 1000,
    "output_rows": 50,
    "output_columns": 5
}
```

### 示例

```python
transform_data(file="data.xlsx", transform_type="pivot", params={"index": "Category", "values": "Sales"})
transform_data(file="data.xlsx", transform_type="melt", params={"id_vars": ["ID"]})
```

---

## fuzzy_match_columns

模糊匹配列名（自动识别相似列名）。

### 参数

| 参数 | 类型 | 必填 | 默认值 | 说明 |
|------|------|------|--------|------|
| `file` | string | 是 | - | Excel 文件路径 |
| `sheet` | string | 是 | - | 工作表名称 |
| `target_columns` | list | 是 | - | 目标列名列表 |
| `threshold` | float | 否 | 0.6 | 相似度阈值 |

### 返回值

```python
{
    "source_columns": ["销售额", "利润", "日期"],
    "matches": [
        {"source": "销售金额", "target": "销售额", "score": 0.85}
    ],
    "match_count": 1
}
```

### 示例

```python
fuzzy_match_columns(file="data.xlsx", sheet="Sheet1", target_columns=["销售额", "利润"])
```

---

## auto_type_infer

自动类型推断（自动检测并转换数据类型）。

### 参数

| 参数 | 类型 | 必填 | 默认值 | 说明 |
|------|------|------|--------|------|
| `file` | string | 是 | - | Excel 文件路径 |
| `sheet` | string | 是 | - | 工作表名称 |
| `range` | string | 否 | None | 数据范围（默认全部） |
| `output` | string | 否 | None | 输出文件路径 |

### 返回值

```python
{
    "success": True,
    "inferred_types": {"A": "int", "B": "float", "C": "str"},
    "conversions": [...],
    "conversion_count": 50
}
```

### 示例

```python
auto_type_infer(file="data.xlsx", sheet="Sheet1")
auto_type_infer(file="data.xlsx", sheet="Sheet1", range="A1:D100", output="output.xlsx")
```

---

## standardize_data

数据标准化（统一日期、数字、文本格式）。

### 参数

| 参数 | 类型 | 必填 | 默认值 | 说明 |
|------|------|------|--------|------|
| `file` | string | 是 | - | Excel 文件路径 |
| `sheet` | string | 是 | - | 工作表名称 |
| `date_format` | string | 否 | None | 日期格式（如 `%Y-%m-%d`） |
| `number_format` | string | 否 | None | 数字格式（如 `%.2f`） |
| `text_case` | string | 否 | None | 文本大小写：`lower`/`upper`/`title` |
| `output` | string | 否 | None | 输出文件路径 |

### 返回值

```python
{
    "success": True,
    "operations": ["date_format", "text_case"],
    "operation_count": 2
}
```

### 示例

```python
standardize_data(file="data.xlsx", sheet="Sheet1", date_format="%Y-%m-%d", text_case="lower")
```

---

## transform_pipeline

数据转换管道（链式执行多个转换步骤）。

### 参数

| 参数 | 类型 | 必填 | 默认值 | 说明 |
|------|------|------|--------|------|
| `file` | string | 是 | - | Excel 文件路径 |
| `steps` | list | 是 | - | 转换步骤列表 |
| `sheet` | string | 否 | None | 工作表名称 |
| `stop_on_error` | bool | 否 | True | 遇到错误是否停止 |

#### steps 支持的类型

| 类型 | 参数 | 说明 |
|------|------|------|
| `convert_type` | `range`, `target_type` | 类型转换 |
| `convert_format` | `range`, `format_type` | 格式转换 |
| `convert_unit` | `range`, `factor` | 单位转换 |
| `standardize` | `text_case` | 标准化 |
| `fill_value` | `range`, `value` | 填充值 |
| `replace_value` | `old_value`, `new_value` | 替换值 |

### 返回值

```python
{
    "file": "data.xlsx",
    "sheet": "Sheet1",
    "steps_executed": 3,
    "steps_succeeded": 3,
    "results": [...]
}
```

### 示例

```python
transform_pipeline(
    file="data.xlsx",
    steps=[
        {"type": "convert_type", "range": "A1:A10", "target_type": "float"},
        {"type": "fill_value", "range": "B1:B10", "value": 0},
        {"type": "replace_value", "old_value": "N/A", "new_value": ""}
    ]
)
```
