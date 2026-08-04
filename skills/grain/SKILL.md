---
name: abacus-grain
description: "粟米章 - 格式转换与数据处理。当需要转换日期、数字、货币格式，或处理条件格式、公式时使用。"
version: 0.2.0
chapter: grain
level: rod
tags: [excel, format, convert, conditional, formula]
---

# 粟米章 - 格式转换与数据处理

## CRITICAL RULES
1. 格式转换前先备份数据
2. 条件格式使用绝对列引用（`$A1`）
3. 颜色值不带 `#` 前缀（用 `FFFF00` 不是 `#FFFF00`）
4. SUMIFS 条件语法：条件范围和条件成对出现

## 能力一览

| 能力 | CLI 命令 | 说明 |
|------|----------|------|
| convert_format | `abacus convert-format` | 转换数据格式（日期、数字、文本） |
| convert_type | `abacus convert-type` | 转换数据类型 |
| convert_unit | `abacus convert-unit` | 转换单位 |
| clean_data | `abacus clean-data` | 数据清洗（去重、缺失值、格式化） |
| transpose | `abacus transpose` | 转置数据（行列互换） |
| text_to_columns | `abacus text-to-columns` | 文本分列（按分隔符拆分） |
| transform_data | `abacus transform-data` | 高级数据转换（透视、转置、合并、重塑） |

---

## 条件格式参考

### 规则类型

| 类型 | 说明 | 参数 |
|------|------|------|
| `cell-value` | 基于单元格值比较 | operator + formula1 (+ formula2 用于 between) |
| `expression` | 基于公式结果 | 仅 formula |

### 运算符（cell-value 类型）

| 运算符 | 说明 | 所需公式 |
|--------|------|----------|
| `equal` | 等于 | formula1 |
| `not-equal` | 不等于 | formula1 |
| `greaterThan` | 大于 | formula1 |
| `lessThan` | 小于 | formula1 |
| `greaterThanOrEqual` | 大于等于 | formula1 |
| `lessThanOrEqual` | 小于等于 | formula1 |
| `between` | 在两者之间 | formula1 AND formula2 |
| `notBetween` | 不在两者之间 | formula1 AND formula2 |

### 格式选项

| 选项 | 类型 | 示例 |
|------|------|------|
| `interiorColor` | 十六进制颜色 | `FFFF00`（黄色填充） |
| `fontColor` | 十六进制颜色 | `FF0000`（红色文字） |
| `fontBold` | bool | `True` / `False` |
| `fontItalic` | bool | `True` / `False` |
| `borderStyle` | string | `thin`, `medium`, `thick` |
| `borderColor` | 十六进制颜色 | `000000` |

### 条件格式示例

```python
# 高亮大于 100 的单元格
abacus_format_range(
    file='data.xlsx',
    sheet='Sheet1',
    range='A1:A10',
    conditional={'type': 'cell', 'operator': 'greaterThan', 'value': 100}
)

# 高亮在 50-100 之间的单元格
abacus_format_range(
    file='data.xlsx',
    sheet='Sheet1',
    range='A1:A10',
    conditional={'type': 'cell', 'operator': 'between', 'value': 50, 'value2': 100}
)

# 如果 A 列是 "Active" 则高亮整行
abacus_format_range(
    file='data.xlsx',
    sheet='Sheet1',
    range='A1:D10',
    conditional={'type': 'expression', 'formula': '=$A1="Active"'}
)
```

### 清除条件格式

```python
abacus_clear_range(
    file='data.xlsx',
    sheet='Sheet1',
    range='A1:E100',
    clear_type='formats'
)
```

### 条件格式常见错误

| 错误 | 原因 | 解决方案 |
|------|------|----------|
| 规则不生效 | cell-value 类型缺少 operator | 始终指定 operator |
| 高亮错误 | 公式中缺少 `$` | 列用 `$A1`，固定单元格用 `$A$1` |
| 颜色不生效 | 缺少十六进制前缀 | 用 `FFFF00` 不是 `#FFFF00` |
| 规则应用到错误单元格 | 范围太大 | 使用精确的 A1 表示法 |

---

## 公式速查表

### 查找函数

**XLOOKUP（推荐）**
```
=XLOOKUP(查找值, 查找范围, 返回范围, [未找到时的值])
=XLOOKUP(D2, A:A, B:B, "未找到")
```

**INDEX + MATCH（兼容性好）**
```
=INDEX(返回范围, MATCH(查找值, 查找范围, 0))
=INDEX(B:B, MATCH(D2, A:A, 0))
```

**VLOOKUP（传统）**
```
=VLOOKUP(查找值, 表范围, 列号, FALSE)
=VLOOKUP(D2, A:B, 2, FALSE)
```

### 条件求和

**SUMIF / SUMIFS**
```
=SUMIF(条件范围, 条件, 求和范围)
=SUMIFS(求和范围, 条件范围1, 条件1, 条件范围2, 条件2)
=SUMIFS(C:C, A:A, "北京", B:B, ">1000")
```

**COUNTIF / COUNTIFS**
```
=COUNTIF(范围, 条件)
=COUNTIFS(A:A, "北京", B:B, ">1000")
```

**AVERAGEIF / AVERAGEIFS**
```
=AVERAGEIF(条件范围, 条件, 平均范围)
=AVERAGEIFS(C:C, A:A, "北京", B:B, ">1000")
```

### 日期函数

```
=TODAY()        # 当前日期
=NOW()          # 当前日期时间
=DATEDIF(开始日期, 结束日期, "Y")  # 年数
=DATEDIF(开始日期, 结束日期, "M")  # 月数
=DATEDIF(开始日期, 结束日期, "D")  # 天数
=EOMONTH(日期, 0)    # 本月月末
=NETWORKDAYS(开始日期, 结束日期)  # 工作日天数
=WORKDAY(开始日期, 天数)          # N个工作日后的日期
```

### 文本函数

```
=LEFT(文本, 字符数)      # 左侧提取
=RIGHT(文本, 字符数)     # 右侧提取
=MID(文本, 起始位置, 字符数)  # 中间提取
=FIND(查找文本, 文本)    # 查找位置（区分大小写）
=SEARCH(查找文本, 文本)  # 查找位置（不区分大小写）
=SUBSTITUTE(文本, 旧文本, 新文本)  # 替换
=TEXTJOIN(",", TRUE, 范围)  # 用逗号连接
```

### 逻辑函数

```
=IF(条件, 真值, 假值)
=IFS(条件1, 值1, 条件2, 值2, TRUE, 默认值)
=SWITCH(表达式, 值1, 结果1, 值2, 结果2, 默认值)
=IFERROR(公式, 错误时的值)
=IFNA(公式, #N/A时的值)
```

### 数组函数（Excel 365）

```
=UNIQUE(范围)           # 去重
=SORT(范围, 列号, 1)    # 排序
=FILTER(范围, 条件)     # 筛选
=SEQUENCE(行数, 列数, 起始值, 步长)
```

### AI 常见公式错误

| 错误 | 正确写法 |
|------|----------|
| `=QUARTER(A1)` | `=ROUNDUP(MONTH(A1)/3, 0)` |
| `=COUNTUNIQUE(A:A)` | `=COUNTA(UNIQUE(A:A))` |
| `=SUMIFS(C:C, A:A = "北京", B:B > 1000)` | `=SUMIFS(C:C, A:A, "北京", B:B, ">1000")` |
| `=AVERAGEIF(C:C, A:A, "北京")` | `=AVERAGEIF(A:A, "北京", C:C)` |

---

## 最佳实践

1. 测试表达式公式前先在 Excel 中验证逻辑
2. 替换时先清除现有规则再应用新规则
3. 行高亮时应用到完整范围（不只是单列）
4. 使用相对行引用（`$A1`）和绝对列引用
5. 保持条件格式简单 - 复杂规则会拖慢 Excel

## 示例

```bash
# 转换格式
abacus convert-format -f data.xlsx -s Sheet1 -r A1:A10 --format-type date

# 转换类型
abacus convert-type -f data.xlsx -s Sheet1 -r B1:B10 --target-type float

# 数据清洗
abacus clean-data -f data.xlsx --operations remove_duplicates handle_missing

# 文本分列
abacus text-to-columns -f data.xlsx -s Sheet1 -c A --delimiter ","
```
