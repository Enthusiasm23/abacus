---
name: abacus-balance
description: "盈不足章 - 数据验证与审计。当需要验证数据范围、类型、公式，或检查文件质量时使用。"
version: 0.2.0
chapter: balance
level: rod
tags: [excel, validate, verify, audit, check, data-quality]
---

# 盈不足章 - 数据验证与审计

## CRITICAL RULES
1. 验证前先读取数据 - 不要假设数据存在
2. 使用结构化错误报告 - 包含具体位置和建议
3. 程序化验证优先于 Excel 验证 - 防止绕过
4. 验证结果必须包含通过/失败状态

## 能力一览

| 能力 | CLI 命令 | 说明 |
|------|----------|------|
| validate_range | `abacus validate-range` | 验证数据范围 |
| validate_type | `abacus validate-type` | 验证数据类型 |
| validate_formula | `abacus validate-formula` | 验证公式正确性 |
| set_data_validation | `abacus set-data-validation` | 设置数据验证规则 |
| file_validate | `abacus validate-file` | 验证 Excel 文件结构 |
| file_analyze | `abacus file-analyze` | 检查文件常见问题 |
| excel_lint | `abacus excel-lint` | 检查 openpyxl 代码质量 |

---

## 范围验证参考

### validate_range 能力说明

validate_range 用于检查数据是否在指定范围内，支持：
- 数值范围验证
- 日期范围验证
- 文本长度验证

### 使用场景

```python
# 验证数值范围
abacus_validate_range(
    file='data.xlsx',
    sheet='Sheet1',
    range='A1:A100',
    min_value=0,
    max_value=1000
)

# 验证日期范围
abacus_validate_range(
    file='data.xlsx',
    sheet='Sheet1',
    range='B1:B100',
    min_value='2024-01-01',
    max_value='2024-12-31'
)
```

### 验证结果格式

```json
{
  "valid": true,
  "total_cells": 100,
  "valid_cells": 100,
  "invalid_cells": 0,
  "invalid_positions": []
}
```

---

## 类型验证参考

### validate_type 能力说明

validate_type 用于检查数据类型是否符合预期，支持：
- 整数验证 (int)
- 浮点数验证 (float)
- 字符串验证 (str)
- 日期验证 (date)

### 使用场景

```python
# 验证整数类型
abacus_validate_type(
    file='data.xlsx',
    sheet='Sheet1',
    range='A1:A100',
    expected_type='int'
)

# 验证浮点数类型
abacus_validate_type(
    file='data.xlsx',
    sheet='Sheet1',
    range='B1:B100',
    expected_type='float'
)
```

### 类型验证最佳实践

1. **先清洗再验证** - 使用 `abacus_clean_data` 处理空值
2. **检查混合类型** - 范围内可能包含不同类型
3. **验证日期格式** - 日期可能存储为文本

---

## 公式验证参考

### validate_formula 能力说明

validate_formula 用于检查公式是否正确，支持：
- 公式语法检查
- 引用验证
- 错误检测

### 使用场景

```python
# 验证单个公式
abacus_validate_formula(
    file='data.xlsx',
    sheet='Sheet1',
    cell='E1'
)

# 批量验证公式
abacus_diagnose_formula(
    file='data.xlsx',
    sheet='Sheet1'
)
```

### 常见公式错误

| 错误 | 原因 | 解决方案 |
|------|------|----------|
| #REF! | 引用的单元格不存在 | 检查引用的单元格是否被删除 |
| #N/A | 查找值未找到 | 检查查找值是否存在 |
| #VALUE! | 参数类型不匹配 | 检查函数参数类型 |
| #NAME? | 函数名或范围名不存在 | 检查函数名拼写 |
| #DIV/0! | 除数为零 | 添加 IFERROR 或检查除数 |

---

## 数据验证规则参考

### set_data_validation 能力说明

set_data_validation 用于创建数据验证规则，支持：
- 下拉列表
- 数值范围限制
- 日期范围限制
- 文本长度限制

### 下拉列表

```python
abacus_set_data_validation(
    file='data.xlsx',
    sheet='Sheet1',
    range='A1:A100',
    validation_type='list',
    formula1='选项1,选项2,选项3'
)
```

### 数值范围

```python
abacus_set_data_validation(
    file='data.xlsx',
    sheet='Sheet1',
    range='B1:B100',
    validation_type='number',
    operator='between',
    formula1='0',
    formula2='1000'
)
```

### 日期范围

```python
abacus_set_data_validation(
    file='data.xlsx',
    sheet='Sheet1',
    range='C1:C100',
    validation_type='date',
    operator='between',
    formula1='2024-01-01',
    formula2='2024-12-31'
)
```

### 文本长度

```python
abacus_set_data_validation(
    file='data.xlsx',
    sheet='Sheet1',
    range='D1:D100',
    validation_type='text_length',
    operator='between',
    formula1='1',
    formula2='50'
)
```

### 数据验证最佳实践

1. **程序化验证优先** - Excel 验证可能被绕过
2. **在写入前验证** - 防止无效数据进入
3. **使用条件格式标记** - 高亮无效数据
4. **提供清晰的错误消息** - 告诉用户如何修复

---

## 文件验证参考

### file_validate 能力说明

file_validate 用于检查 Excel 文件结构，支持：
- ZIP 格式验证
- XML 结构验证
- 公式错误检测

### 使用场景

```python
# 验证文件结构
abacus_validate_file(file='data.xlsx')
```

### 验证结果

```json
{
  "valid": true,
  "zip_format": true,
  "xml_structure": true,
  "formula_errors": 0,
  "warnings": []
}
```

---

## 文件审计参考

### file_analyze 能力说明

file_analyze 用于检查 Excel 文件的 10 类常见问题：
- 工作表结构
- 合并单元格
- 数据类型
- 公式错误
- 格式问题

### 使用场景

```python
# 分析文件问题
abacus_file_analyze(file='data.xlsx')
```

### 常见问题类型

| 问题 | 说明 | 严重程度 |
|------|------|----------|
| 合并单元格 | 可能影响排序和筛选 | 警告 |
| 空行 | 影响数据分析 | 警告 |
| 公式错误 | 需要修复 | 错误 |
| 格式不一致 | 影响美观 | 建议 |

---

## Excel 代码审计参考

### excel_lint 能力说明

excel_lint 用于检查 openpyxl 代码的 10 类常见问题：
- 未关闭工作簿
- 未验证文件路径
- 未处理异常
- 性能问题

### 使用场景

```python
# 检查代码质量
abacus_excel_lint(code="import openpyxl\nwb = openpyxl.load_workbook('test.xlsx')")

# 检查文件中的代码
abacus_excel_lint(file='script.py')
```

### 常见代码问题

| 问题 | 说明 | 严重程度 |
|------|------|----------|
| 未关闭工作簿 | 资源泄漏 | 错误 |
| 未验证路径 | 可能导致崩溃 | 错误 |
| 未处理异常 | 用户体验差 | 警告 |
| 整列引用 | 性能问题 | 警告 |

---

## Excel 常见坑

### 1. 公式变字符串
**错误：** 写入公式后变成文本
```python
ws["A1"] = "SUM(B1:B10)"  # 错误：缺少 =
ws["A1"] = "=SUM(B1:B10)" # 正确
```

### 2. data_only=True 丢公式
**错误：** 读取后保存丢失公式
```python
wb = load_workbook("file.xlsx", data_only=True)
wb.save("file.xlsx")  # 公式永久丢失！
```

### 3. PatternFill 缺 fill_type
**错误：** 颜色不显示
```python
cell.fill = PatternFill(fgColor="FF0000")  # 错误
cell.fill = PatternFill(pattern_type="solid", fgColor="FF0000")  # 正确
```

### 4. 整列引用拖慢性能
**错误：** SUM(A:A) 引用百万行
**正确：** SUM(A1:A1000) 限定范围

### 5. 合并单元格排序失败
**错误：** 合并单元格导致排序异常
**正确：** 先取消合并再排序

---

## 财务模型规范

### 颜色编码标准

| 颜色 | RGB | 用途 |
|------|-----|------|
| 蓝色文本 | 0,0,255 | 硬编码输入、用户可修改的假设值 |
| 黑色文本 | 0,0,0 | 所有公式和计算 |
| 绿色文本 | 0,128,0 | 同一工作簿内跨表引用 |
| 红色文本 | 255,0,0 | 外部文件链接 |
| 黄色背景 | 255,255,0 | 需要关注的关键假设 |

### 数字格式标准

| 类型 | 格式 | 示例 |
|------|------|------|
| 年份 | 文本字符串 | "2024" 而非 "2,024" |
| 货币 | $#,##0 | "$1,234" |
| 零值 | 显示为 "-" | 使用数字格式 |
| 百分比 | 0.0% | "12.3%" |
| 倍数 | 0.0x | "15.2x" |
| 负数 | 括号 | "(123)" 而非 "-123" |

### 公式验证清单

- [ ] 测试 2-3 个样本引用
- [ ] 确认列映射正确
- [ ] 检查 NaN 值
- [ ] 检查除零错误
- [ ] 验证所有单元格引用指向正确位置

---

## 最佳实践

1. **验证前读取数据** - 不要假设数据存在
2. **使用结构化错误** - 包含位置和建议
3. **程序化验证优先** - Excel 验证可能被绕过
4. **提供修复建议** - 告诉用户如何解决问题
5. **验证结果包含状态** - 通过/失败/警告

## 示例

```bash
# 验证数据范围
abacus validate-range -f data.xlsx -s Sheet1 -r A1:A100 --min-value 0 --max-value 1000

# 验证数据类型
abacus validate-type -f data.xlsx -s Sheet1 -r A1:A100 --expected-type int

# 验证公式
abacus validate-formula -f data.xlsx -s Sheet1 --cell E1

# 设置数据验证
abacus set-data-validation -f data.xlsx -s Sheet1 -r A1:A100 --validation-type list --formula1 "选项1,选项2,选项3"

# 验证文件结构
abacus validate-file -f data.xlsx

# 分析文件问题
abacus file-analyze -f data.xlsx
```