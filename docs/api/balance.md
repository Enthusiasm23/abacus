# 盈不足章 (Balance) - 数据验证与审计

> 盈不足章负责数据验证、文件验证、质量检查和代码审计。

---

## 工具列表

| 工具 | 描述 |
|------|------|
| `validate_range` | 验证数据范围 |
| `validate_type` | 验证数据类型 |
| `validate_formula` | 验证公式正确性 |
| `validate_file` | 验证 Excel 文件结构 |
| `set_data_validation` | 设置单元格数据验证规则 |
| `data_quality_check` | 数据质量检测 |
| `excel_lint` | 代码审计 |
| `file_analyze` | 文件分析 |

---

## validate_range

验证数据范围。

### 参数

| 参数 | 类型 | 必填 | 默认值 | 说明 |
|------|------|------|--------|------|
| `file` | string | 是 | - | Excel 文件路径 |
| `sheet` | string | 是 | - | 工作表名称 |
| `range` | string | 是 | - | 数据范围（A1 表示法） |
| `min_value` | float | 否 | None | 最小值 |
| `max_value` | float | 否 | None | 最大值 |

### 返回值

```python
{
    "valid": True,
    "invalid_count": 3,
    "invalid_cells": [
        {"cell": "A5", "value": -10, "reason": "below_min"},
        {"cell": "A8", "value": 150, "reason": "above_max"},
        {"cell": "A12", "value": None, "reason": "empty"}
    ]
}
```

### 示例

```python
validate_range(file="data.xlsx", sheet="Sheet1", range="A1:A100", min_value=0, max_value=100)
validate_range(file="data.xlsx", sheet="Sheet1", range="B1:B50", min_value=0)
```

---

## validate_type

验证数据类型。

### 参数

| 参数 | 类型 | 必填 | 默认值 | 说明 |
|------|------|------|--------|------|
| `file` | string | 是 | - | Excel 文件路径 |
| `sheet` | string | 是 | - | 工作表名称 |
| `range` | string | 是 | - | 数据范围（A1 表示法） |
| `expected_type` | string | 是 | - | 期望类型：`int`/`float`/`str`/`date` |

### 返回值

```python
{
    "valid": True,
    "invalid_count": 5,
    "invalid_cells": [
        {"cell": "A3", "value": "abc", "reason": "type_mismatch"}
    ]
}
```

### 示例

```python
validate_type(file="data.xlsx", sheet="Sheet1", range="A1:A100", expected_type="float")
validate_type(file="data.xlsx", sheet="Sheet1", range="B1:B50", expected_type="date")
```

---

## validate_formula

验证公式正确性。

### 参数

| 参数 | 类型 | 必填 | 默认值 | 说明 |
|------|------|------|--------|------|
| `file` | string | 是 | - | Excel 文件路径 |
| `sheet` | string | 否 | None | 工作表名称（默认所有） |
| `cell` | string | 否 | None | 单元格位置（默认所有公式） |

### 返回值

```python
{
    "valid": False,
    "formula": "=SUM(A1:A10)/B1",
    "error": "#DIV/0!"
}
```

### 示例

```python
validate_formula(file="data.xlsx", sheet="Sheet1", cell="E1")
validate_formula(file="data.xlsx")
```

---

## validate_file

验证 Excel 文件结构（ZIP 格式、XML 结构、公式错误）。

### 参数

| 参数 | 类型 | 必填 | 默认值 | 说明 |
|------|------|------|--------|------|
| `file` | string | 是 | - | Excel 文件路径 |

### 返回值

```python
{
    "file": "data.xlsx",
    "valid": True,
    "checks": ["zip_structure", "xml_parse", "shared_strings"],
    "errors": [],
    "warnings": ["Sheet 'Config' has 0 rows"]
}
```

### 示例

```python
validate_file(file="data.xlsx")
```

---

## set_data_validation

设置单元格数据验证规则。

### 参数

| 参数 | 类型 | 必填 | 默认值 | 说明 |
|------|------|------|--------|------|
| `file` | string | 是 | - | Excel 文件路径 |
| `sheet` | string | 是 | - | 工作表名称 |
| `range` | string | 是 | - | 数据范围 |
| `validation_type` | string | 是 | - | 验证类型：`list`/`number`/`date`/`text_length` |
| `operator` | string | 否 | None | 运算符：`between`/`notBetween`/`equal`/`notEqual` 等 |
| `formula1` | string | 否 | None | 验证公式1 |
| `formula2` | string | 否 | None | 验证公式2（between 时需要） |
| `error_message` | string | 否 | None | 错误提示消息 |

### 返回值

```python
{
    "success": True,
    "sheet": "Sheet1",
    "range": "A1:A10",
    "validation_type": "list",
    "applied": True
}
```

### 示例

```python
# 下拉列表
set_data_validation(file="data.xlsx", sheet="Sheet1", range="A1:A10", validation_type="list", formula1="是,否")

# 数值范围
set_data_validation(
    file="data.xlsx", sheet="Sheet1", range="B1:B10",
    validation_type="number", operator="between",
    formula1="0", formula2="100"
)
```

---

## data_quality_check

数据质量检测（自动检测空值、异常值、重复数据）。

### 参数

| 参数 | 类型 | 必填 | 默认值 | 说明 |
|------|------|------|--------|------|
| `file` | string | 是 | - | Excel 文件路径 |
| `sheet` | string | 是 | - | 工作表名称 |
| `range` | string | 否 | None | 数据范围（默认全部） |

### 返回值

```python
{
    "quality_score": 85,
    "issues": [
        {"type": "null", "column": "Name", "count": 5},
        {"type": "duplicate", "column": "ID", "count": 3},
        {"type": "outlier", "column": "Sales", "count": 2}
    ],
    "issue_count": 10,
    "null_counts": {"Name": 5, "Sales": 2},
    "type_distribution": {"ID": "int", "Name": "str", "Sales": "float"}
}
```

### 示例

```python
data_quality_check(file="data.xlsx", sheet="Sheet1")
data_quality_check(file="data.xlsx", sheet="Sales", range="A1:D100")
```

---

## excel_lint

检查 openpyxl 代码的 10 类常见问题。

### 参数

| 参数 | 类型 | 必填 | 默认值 | 说明 |
|------|------|------|--------|------|
| `code` | string | 条件 | None | Python 代码内容（与 file 二选一） |
| `file` | string | 条件 | None | Python 文件路径（与 code 二选一） |

### 检查项

1. 未关闭工作簿
2. 未保存工作簿
3. 单元格赋值类型错误
4. 合并单元格操作错误
5. 样式对象重复使用
6. 无效的列名引用
7. 未处理的异常
8. 性能问题（逐行写入）
9. 日期格式问题
10. 内存泄漏风险

### 返回值

```python
{
    "issues": [
        {"line": 15, "severity": "warning", "message": "Workbook not closed", "rule": "unclosed_workbook"},
        {"line": 23, "severity": "error", "message": "Invalid column reference", "rule": "invalid_column"}
    ],
    "issue_count": 2,
    "severity": "warning"
}
```

### 示例

```python
excel_lint(code="import openpyxl\nwb = openpyxl.load_workbook('test.xlsx')")
excel_lint(file="/path/to/script.py")
```

---

## file_analyze

检查 Excel 文件的 10 类常见问题。

### 参数

| 参数 | 类型 | 必填 | 默认值 | 说明 |
|------|------|------|--------|------|
| `file` | string | 是 | - | Excel 文件路径 |

### 检查项

1. 文件大小异常
2. 工作表数量异常
3. 数据范围异常
4. 合并单元格问题
5. 空工作表
6. 公式错误
7. 格式问题
8. 数据类型不一致
9. 命名范围问题
10. 链接问题

### 返回值

```python
{
    "issues": [
        {"type": "empty_sheet", "sheet": "Config", "severity": "warning"},
        {"type": "formula_error", "cell": "E1", "error": "#REF!", "severity": "error"}
    ],
    "issue_count": 2,
    "severity": "warning"
}
```

### 示例

```python
file_analyze(file="data.xlsx")
```
