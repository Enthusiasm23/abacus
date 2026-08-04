# 方程章 (Equation) - 公式计算与诊断

> 方程章负责公式创建、诊断、重算、生成和数组公式。

---

## 工具列表

| 工具 | 描述 |
|------|------|
| `create_formula` | 在指定单元格创建公式 |
| `diagnose_formula` | 诊断公式错误 |
| `recalc_formulas` | 公式重算（使用 LibreOffice） |
| `generate_formula` | 生成常用 Excel 公式 |
| `set_array_formula` | 设置数组公式 |

---

## create_formula

在指定单元格创建公式。

### 参数

| 参数 | 类型 | 必填 | 默认值 | 说明 |
|------|------|------|--------|------|
| `file` | string | 是 | - | Excel 文件路径 |
| `sheet` | string | 是 | - | 工作表名称 |
| `cell` | string | 是 | - | 单元格位置 |
| `formula` | string | 是 | - | 公式内容 |

### 返回值

```python
{
    "success": True,
    "cell": "E1",
    "formula": "SUM(A1:D1)"
}
```

### 示例

```python
create_formula(file="data.xlsx", sheet="Sheet1", cell="E1", formula="SUM(A1:D1)")
create_formula(file="data.xlsx", sheet="Sheet1", cell="F1", formula="AVERAGE(B1:B10)")
create_formula(file="data.xlsx", sheet="Sheet1", cell="G1", formula='IF(A1>100,"High","Low")')
```

---

## diagnose_formula

诊断公式错误（分析 #REF!, #N/A, #VALUE!, #NAME?, #DIV/0! 等错误）。

### 参数

| 参数 | 类型 | 必填 | 默认值 | 说明 |
|------|------|------|--------|------|
| `file` | string | 是 | - | Excel 文件路径 |
| `sheet` | string | 否 | None | 工作表名称（默认检查所有） |
| `cell` | string | 否 | None | 单元格位置（默认检查所有公式） |

### 支持的错误类型

| 错误 | 说明 |
|------|------|
| `#REF!` | 引用错误 |
| `#N/A` | 值不可用 |
| `#VALUE!` | 值类型错误 |
| `#NAME?` | 名称错误 |
| `#DIV/0!` | 除零错误 |
| `#NULL!` | 空引用错误 |
| `#NUM!` | 数字错误 |

### 返回值

```python
{
    "file": "data.xlsx",
    "formulas_checked": 50,
    "errors_found": 3,
    "errors": [
        {
            "sheet": "Sheet1",
            "cell": "E1",
            "formula": "=SUM(A1:A10)/B1",
            "error": "#DIV/0!",
            "description": "Division by zero - B1 is empty or zero"
        },
        {
            "sheet": "Sheet1",
            "cell": "F5",
            "formula": "=VLOOKUP(D5,Sheet2!A:B,2,0)",
            "error": "#N/A",
            "description": "Value not found in lookup range"
        }
    ]
}
```

### 示例

```python
diagnose_formula(file="data.xlsx")
diagnose_formula(file="data.xlsx", sheet="Sheet1")
diagnose_formula(file="data.xlsx", sheet="Sheet1", cell="E1")
```

---

## recalc_formulas

公式重算：使用 LibreOffice 重算 Excel 公式（扫描所有错误）。

### 参数

| 参数 | 类型 | 必填 | 默认值 | 说明 |
|------|------|------|--------|------|
| `file` | string | 是 | - | Excel 文件路径 |
| `output` | string | 否 | None | 输出文件路径 |

### 依赖

需要安装 LibreOffice 才能使用此功能。

### 返回值

```python
{
    "file": "data.xlsx",
    "output": "recalc.xlsx",
    "errors_found": 2,
    "errors": [
        {"sheet": "Sheet1", "cell": "E1", "error": "#DIV/0!"}
    ],
    "recalculated": True
}
```

### 示例

```python
recalc_formulas(file="data.xlsx")
recalc_formulas(file="data.xlsx", output="recalc.xlsx")
```

---

## generate_formula

生成常用 Excel 公式（VLOOKUP、SUMIFS、IF 等）。

### 参数

| 参数 | 类型 | 必填 | 默认值 | 说明 |
|------|------|------|--------|------|
| `formula_type` | string | 是 | - | 公式类型 |
| `params` | dict | 是 | - | 公式参数 |
| `file` | string | 否 | None | Excel 文件路径（写入时使用） |
| `sheet` | string | 否 | None | 工作表名称 |
| `cell` | string | 否 | None | 单元格位置 |

### 支持的公式类型

| 类型 | 参数 | 说明 |
|------|------|------|
| `vlookup` | `lookup_value`, `table_array`, `col_index` | 垂直查找 |
| `sumifs` | `sum_range`, `criteria_range`, `criteria` | 条件求和 |
| `if` | `logical_test`, `value_true`, `value_false` | 条件判断 |
| `today` | - | 今日日期 |
| `npv` | `rate`, `values` | 净现值 |
| `pmt` | `rate`, `nper`, `pv` | 贷款支付 |
| `irr` | `values` | 内部收益率 |

### 返回值

```python
{
    "formula": "=VLOOKUP(D2,B:C,2,0)",
    "description": "在 B 列查找 D2 的值，返回 C 列对应值"
}
```

### 示例

```python
# VLOOKUP
generate_formula(
    formula_type="vlookup",
    params={"lookup_value": "D2", "table_array": "B:C", "col_index": 2}
)

# SUMIFS
generate_formula(
    formula_type="sumifs",
    params={"sum_range": "C:C", "criteria_range": "A:A", "criteria": "East"}
)

# IF
generate_formula(
    formula_type="if",
    params={"logical_test": "A1>100", "value_true": '"High"', "value_false": '"Low"'}
)

# 写入到单元格
generate_formula(
    formula_type="vlookup",
    params={"lookup_value": "D2", "table_array": "B:C", "col_index": 2},
    file="data.xlsx",
    sheet="Sheet1",
    cell="E1"
)
```

---

## set_array_formula

设置数组公式。

### 参数

| 参数 | 类型 | 必填 | 默认值 | 说明 |
|------|------|------|--------|------|
| `file` | string | 是 | - | Excel 文件路径 |
| `sheet` | string | 是 | - | 工作表名称 |
| `range` | string | 是 | - | 数组公式范围 |
| `formula` | string | 是 | - | 数组公式内容 |

### 返回值

```python
{
    "success": True,
    "action": "set_array_formula",
    "sheet": "Sheet1",
    "range": "A1:A10",
    "formula": "SUM(B1:B10*C1:C10)"
}
```

### 示例

```python
# 多条件求和
set_array_formula(
    file="data.xlsx",
    sheet="Sheet1",
    range="E1",
    formula="SUM((A1:A10=\"East\")*(B1:B10>100)*C1:C10)"
)

# 数组乘法求和
set_array_formula(
    file="data.xlsx",
    sheet="Sheet1",
    range="F1",
    formula="SUM(B1:B10*C1:C10)"
)
```
