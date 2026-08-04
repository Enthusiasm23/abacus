---
name: abacus-dimension
description: "少广章 - 反向计算与推导。当需要已知结果反推参数、执行数学计算、解方程时使用。"
version: 0.2.0
chapter: dimension
level: rod
tags: [excel, reverse, calculate, derive, equation]
---

# 少广章 - 反向计算与推导

## CRITICAL RULES
1. 始终使用 Excel 公式，不要在 Python 中计算后硬编码值
2. 公式必须以 `=` 开头，使用 A1 表示法
3. 反向推导前先验证目标值的合理性
4. 修改后必须保存 - openpyxl 变更仅在内存中

## 能力一览

| 能力 | CLI 命令 | 说明 |
|------|----------|------|
| find_dimension | `abacus find-dimension` | 已知面积求边长（几何计算） |
| derive | `abacus derive` | 反向推导（已知结果反推参数） |
| calculate | `abacus calculate` | 执行数学计算（表达式求值） |
| solve_equation | `abacus solve-equation` | 解方程（求解未知数） |

---

## 几何计算参考

### find_dimension 支持的形状

| 形状 | 参数 | 公式 |
|------|------|------|
| rectangle | area, known_side | side = area / known_side |
| circle | area | radius = √(area / π) |

### 使用场景

```bash
# 已知矩形面积和一边长，求另一边
abacus find-dimension --area 100 --known-side 10 --shape rectangle

# 已知圆面积求半径
abacus find-dimension --area 314.16 --shape circle
```

---

## 反向推导参考

### derive 能力说明

derive 用于已知目标值，反推需要的参数值。支持：
- 目标值求解
- 敏感性分析
- 参数调整建议

### 使用场景

```bash
# 已知目标利润，反推需要的销售额
abacus derive -f model.xlsx -s Sheet1 -c A1 --target-value 10000 --formula "A2*B2*(1-C2)"

# 已知目标结果，反推参数
abacus derive -f data.xlsx -s Sheet1 -c B10 --target-value 500 --formula "SUM(A1:A9)"
```

### 工作流程

1. 确定目标值和相关公式
2. 分析公式中的变量关系
3. 逐步调整参数值
4. 验证推导结果

---

## 数学计算参考

### calculate 能力说明

calculate 用于执行数学表达式求值，支持：
- 基本运算：`+`, `-`, `*`, `/`
- 函数调用：`SUM`, `AVERAGE`, `MAX`, `MIN` 等
- 变量替换

### 使用场景

```bash
# 简单计算
abacus calculate --expression "2 + 3 * 4"

# 带变量的计算
abacus calculate --expression "x * y + z" --variables '{"x": 10, "y": 20, "z": 5}'

# 复杂表达式
abacus calculate --expression "SUM(A1:A10) / COUNT(A1:A10)"
```

---

## 公式最佳实践

### 核心原则

**始终使用 Excel 公式，不要在 Python 中计算后硬编码值**

### 常用公式模板

**求和**
```
=SUM(A1:A100)
=SUMIF(B1:B100,"条件",A1:A100)
=SUMIFS(A1:A100,B1:B100,"条件1",C1:C100,"条件2")
```

**查找**
```
=VLOOKUP(查找值,范围,列号,FALSE)
=INDEX(范围,MATCH(查找值,查找范围,0))
=XLOOKUP(查找值,查找范围,返回范围)
```

**条件判断**
```
=IF(条件,真值,假值)
=IFS(条件1,值1,条件2,值2,TRUE,默认值)
=SWITCH(表达式,值1,结果1,值2,结果2,默认值)
```

**日期**
```
=TODAY()
=NOW()
=DATEDIF(开始日期,结束日期,"Y"/"M"/"D")
=EOMONTH(日期,月数)
```

**文本**
```
=LEFT(文本,字符数)
=RIGHT(文本,字符数)
=MID(文本,起始位置,字符数)
=SUBSTITUTE(文本,旧文本,新文本)
=CONCATENATE(文本1,文本2,...)
```

### AI 常见错误

| 错误 | 正确写法 |
|------|----------|
| `=QUARTER(A1)` | `=ROUNDUP(MONTH(A1)/3, 0)` |
| `=COUNTUNIQUE(A:A)` | `=COUNTA(UNIQUE(A:A))` |
| `=SUMIFS(C:C, A:A = "北京", B:B > 1000)` | `=SUMIFS(C:C, A:A, "北京", B:B, ">1000")` |
| `=AVERAGEIF(C:C, A:A, "北京")` | `=AVERAGEIF(A:A, "北京", C:C)` |

---

## 公式诊断参考

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

### 常见问题

**VLOOKUP 返回 #N/A**
- 检查查找值是否存在
- 检查查找范围是否正确
- 检查列号是否正确

**SUM 返回 #VALUE!**
- 检查范围是否包含文本
- 检查是否有错误值

**IF 返回 #NAME?**
- 检查函数名拼写
- 检查是否使用了中文括号

---

## 最佳实践

1. **使用 Excel 公式** - 始终使用公式而非硬编码值
2. **验证输入参数** - 确保目标值和公式正确
3. **测试小规模数据** - 先验证逻辑再处理大数据
4. **使用绝对引用** - 避免公式复制时引用错误
5. **添加错误处理** - 使用 IFERROR 处理异常情况

## 示例

```bash
# 已知面积求边长
abacus find-dimension --area 100 --known-side 10 --shape rectangle

# 反向推导
abacus derive -f model.xlsx -s Sheet1 -c A1 --target-value 10000 --formula "A2*B2"

# 数学计算
abacus calculate --expression "SQRT(144) + 2 * 3"

# 解方程
abacus solve-equation --equation "2x + 3 = 7"
```