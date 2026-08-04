---
name: abacus-equation
description: "方程章 - 公式计算与诊断。当需要创建公式、解方程、诊断错误、生成公式时使用。"
version: 0.2.0
chapter: equation
level: rod
tags: [excel, formula, calculate, diagnose, equation, solve]
---

# 方程章 - 公式计算与诊断

## CRITICAL RULES
1. 始终使用 Excel 公式，不要在 Python 中计算后硬编码值
2. 公式必须以 `=` 开头，使用 A1 表示法
3. SUMIFS 条件语法：条件范围和条件成对出现
4. 修改后必须保存 - openpyxl 变更仅在内存中

## 能力一览

| 能力 | CLI 命令 | 说明 |
|------|----------|------|
| create_formula | `abacus formula` | 在指定单元格创建公式 |
| solve_equation | `abacus solve-equation` | 解方程（求解未知数） |
| diagnose_formula | `abacus diagnose-formula` | 诊断公式错误 |
| recalc_formulas | `abacus recalc-formulas` | 重算 Excel 公式 |
| auto_sum | `abacus auto-sum` | 自动求和 |
| generate_formula | `abacus generate-formula` | 生成常用 Excel 公式 |
| set_array_formula | `abacus set-array-formula` | 设置数组公式 |

---

## 公式创建参考

### create_formula 能力说明

create_formula 用于在指定单元格创建公式，支持：
- 简单公式
- 复杂公式
- 跨工作表引用

### 使用场景

```python
# 创建简单公式
abacus_create_formula(
    file='data.xlsx',
    sheet='Sheet1',
    cell='E1',
    formula='SUM(A1:D1)'
)

# 创建复杂公式
abacus_create_formula(
    file='data.xlsx',
    sheet='Sheet1',
    cell='F1',
    formula='VLOOKUP(A1,Sheet2!A:B,2,FALSE)'
)
```

### 公式语法

| 元素 | 说明 | 示例 |
|------|------|------|
| 函数名 | 不区分大小写 | `SUM`, `sum`, `Sum` |
| 参数 | 用逗号分隔 | `SUM(A1,A2,A3)` |
| 范围 | 冒号连接 | `SUM(A1:A10)` |
| 引用 | 单元格地址 | `A1`, `$A$1`, `Sheet2!A1` |

---

## 方程求解参考

### solve_equation 能力说明

solve_equation 用于解数学方程，支持：
- 一元一次方程
- 一元二次方程
- 简单代数方程

### 使用场景

```python
# 解一元一次方程
abacus_solve_equation(equation="2x + 3 = 7")

# 解一元二次方程
abacus_solve_equation(equation="x^2 - 5x + 6 = 0")
```

### 方程语法

| 符号 | 说明 | 示例 |
|------|------|------|
| `+` | 加法 | `x + 3` |
| `-` | 减法 | `x - 2` |
| `*` | 乘法 | `2 * x` |
| `/` | 除法 | `x / 4` |
| `^` | 幂运算 | `x ^ 2` |
| `=` | 等号 | `2x + 3 = 7` |

---

## 公式诊断参考

### diagnose_formula 能力说明

diagnose_formula 用于诊断公式错误，支持：
- 单个公式诊断
- 整个工作表诊断
- 批量诊断

### 使用场景

```python
# 诊断单个公式
abacus_diagnose_formula(
    file='data.xlsx',
    sheet='Sheet1',
    cell='E1'
)

# 诊断整个工作表
abacus_diagnose_formula(
    file='data.xlsx',
    sheet='Sheet1'
)
```

### 常见错误类型

| 错误 | 原因 | 解决方案 |
|------|------|----------|
| #REF! | 引用的单元格不存在 | 检查引用的单元格是否被删除 |
| #N/A | 查找值未找到 | 检查查找值是否存在 |
| #VALUE! | 参数类型不匹配 | 检查函数参数类型 |
| #NAME? | 函数名或范围名不存在 | 检查函数名拼写 |
| #DIV/0! | 除数为零 | 添加 IFERROR 或检查除数 |
| #NUM! | 数值超出范围 | 检查数值是否在有效范围内 |
| #NULL! | 区域引用不相交 | 检查区域引用是否正确 |

### 诊断步骤

1. 定位错误单元格
2. 检查公式语法
3. 验证引用单元格
4. 检查数据类型
5. 测试公式逻辑

---

## 公式重算参考

### recalc_formulas 能力说明

recalc_formulas 用于使用 LibreOffice 重算 Excel 公式，支持：
- 扫描所有公式错误
- 重算整个工作簿
- 生成错误报告

### 使用场景

```python
# 重算工作簿
abacus_recalc_formulas(
    file='data.xlsx',
    output='recalculated.xlsx'
)
```

### 重算最佳实践

1. **批量写入后重算** - 确保公式计算正确
2. **验证错误** - 检查重算后的错误
3. **备份原文件** - 重算前创建备份

---

## 自动求和参考

### auto_sum 能力说明

auto_sum 用于在范围内设置 SUM 公式，支持：
- 向下求和（列）
- 向右求和（行）

### 使用场景

```python
# 向下求和
abacus_auto_sum(
    file='data.xlsx',
    sheet='Sheet1',
    range='A1:D10',
    direction='down'
)

# 向右求和
abacus_auto_sum(
    file='data.xlsx',
    sheet='Sheet1',
    range='A1:D10',
    direction='right'
)
```

---

## 公式生成参考

### generate_formula 能力说明

generate_formula 用于生成常用 Excel 公式，支持：
- VLOOKUP
- SUMIFS
- IF
- 日期函数
- 文本函数
- 财务函数

### 常用公式模板

**查找函数**

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

**条件求和**

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

**日期函数**
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

**文本函数**
```
=LEFT(文本, 字符数)      # 左侧提取
=RIGHT(文本, 字符数)     # 右侧提取
=MID(文本, 起始位置, 字符数)  # 中间提取
=FIND(查找文本, 文本)    # 查找位置（区分大小写）
=SEARCH(查找文本, 文本)  # 查找位置（不区分大小写）
=SUBSTITUTE(文本, 旧文本, 新文本)  # 替换
=TEXTJOIN(",", TRUE, 范围)  # 用逗号连接
```

**逻辑函数**
```
=IF(条件, 真值, 假值)
=IFS(条件1, 值1, 条件2, 值2, TRUE, 默认值)
=SWITCH(表达式, 值1, 结果1, 值2, 结果2, 默认值)
=IFERROR(公式, 错误时的值)
=IFNA(公式, #N/A时的值)
```

**数组函数（Excel 365）**
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

## 数组公式参考

### set_array_formula 能力说明

set_array_formula 用于创建数组公式，支持：
- 多单元格数组公式
- 动态数组公式

### 使用场景

```python
# 创建数组公式
abacus_set_array_formula(
    file='data.xlsx',
    sheet='Sheet1',
    range='C1:C10',
    formula='SUM(A1:A10*B1:B10)'
)
```

### 数组公式语法

- Excel 365：直接输入，自动溢出
- Excel 2019 及更早：按 Ctrl+Shift+Enter 输入

---

## 中文关键词路由

用户不需要知道工具名，只需描述需求，系统自动匹配。

### 公式场景

| 用户原话 | 关键词 | 推荐公式 |
|---------|--------|----------|
| 查价格/查数据/查找匹配 | 查 | VLOOKUP, XLOOKUP, INDEX-MATCH |
| 算总分/求和/加总 | 算/求/加 | SUM, SUMIF, SUMIFS |
| 看有几条/计数/统计数量 | 几条/计数/数量 | COUNT, COUNTIF, COUNTIFS |
| 算平均/平均值 | 平均 | AVERAGE, AVERAGEIF |
| 最大值/最高/最多 | 最 | MAX, MIN |
| 如果...就/判断/显示达标 | 如果/判断 | IF, IFS, SWITCH |
| 排名/排序 | 排 | RANK, SORT |
| 去重/不重复 | 去重 | UNIQUE |
| 提取/截取/取左边 | 取/提取 | LEFT, RIGHT, MID |
| 替换/改成 | 替换 | SUBSTITUTE, REPLACE |
| 合并/拼接/连接 | 合并/拼接 | CONCATENATE, TEXTJOIN, & |
| 日期差/相差几天 | 差/天数 | DATEDIF, NETWORKDAYS |
| 月末/月底 | 月末 | EOMONTH |
| 百分比/占比 | 百分比 | 公式: =B2/SUM(B:B) |

---

## 最佳实践

1. **使用 Excel 公式** - 始终使用公式而非硬编码值
2. **验证公式语法** - 使用 `abacus_diagnose_formula` 检查
3. **批量写入后重算** - 确保公式计算正确
4. **使用绝对引用** - 避免公式复制时引用错误
5. **添加错误处理** - 使用 IFERROR 处理异常情况

## 示例

```bash
# 创建公式
abacus formula -f data.xlsx -s Sheet1 --cell E1 --formula "SUM(A1:D1)"

# 解方程
abacus solve-equation --equation "2x + 3 = 7"

# 诊断公式
abacus diagnose-formula -f data.xlsx -s Sheet1 --cell E1

# 重算公式
abacus recalc-formulas -f data.xlsx --output recalculated.xlsx

# 自动求和
abacus auto-sum -f data.xlsx -s Sheet1 -r A1:D10 --direction down

# 生成公式
abacus generate-formula --formula-type vlookup --params '{"lookup_value": "A1", "table_array": "B:C", "col_index": 2}'
```