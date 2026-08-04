# 少广章 (Dimension) - 反向计算与推导

> 少广章负责数学计算、反向推导、解方程和自动求和。

---

## 工具列表

| 工具 | 描述 |
|------|------|
| `find_dimension` | 已知面积求边长 |
| `derive` | 反向推导 |
| `calculate` | 执行计算 |
| `solve_equation` | 解方程 |
| `auto_sum` | 自动求和 |

---

## find_dimension

已知面积求边长。

### 参数

| 参数 | 类型 | 必填 | 默认值 | 说明 |
|------|------|------|--------|------|
| `area` | float | 是 | - | 面积值 |
| `shape` | string | 否 | "rectangle" | 形状：`rectangle`/`circle` |
| `known_side` | float | 条件 | None | 已知边长（矩形时必填） |

### 返回值

```python
# 矩形
{
    "shape": "rectangle",
    "area": 100,
    "width": 10,
    "height": 10
}

# 圆形
{
    "shape": "circle",
    "area": 100,
    "radius": 5.64
}
```

### 示例

```python
find_dimension(area=100, shape="rectangle", known_side=10)
find_dimension(area=100, shape="circle")
```

---

## derive

反向推导（已知结果反推参数）。

### 参数

| 参数 | 类型 | 必填 | 默认值 | 说明 |
|------|------|------|--------|------|
| `file` | string | 是 | - | Excel 文件路径 |
| `sheet` | string | 是 | - | 工作表名称 |
| `cell` | string | 是 | - | 单元格位置 |
| `target_value` | float | 是 | - | 目标值 |
| `formula` | string | 是 | - | 公式表达式 |

### 返回值

```python
{
    "result": 200.0,
    "iterations": 15
}
```

### 示例

```python
derive(file="data.xlsx", sheet="Sheet1", cell="A1", target_value=1000, formula="A2*B2")
```

---

## calculate

执行计算。

### 参数

| 参数 | 类型 | 必填 | 默认值 | 说明 |
|------|------|------|--------|------|
| `expression` | string | 是 | - | 计算表达式 |
| `variables` | dict | 否 | None | 变量值 |

### 支持的运算符

- 基本运算：`+`, `-`, `*`, `/`, `//`, `%`, `**`
- 函数：`sin`, `cos`, `tan`, `log`, `sqrt`, `abs`, `round`
- 常量：`pi`, `e`

### 返回值

```python
{
    "result": 14,
    "expression": "2 + 3 * 4"
}
```

### 示例

```python
calculate(expression="2 + 3 * 4")
calculate(expression="x + y", variables={"x": 10, "y": 20})
calculate(expression="sqrt(144) + pi")
```

---

## solve_equation

解方程。

### 参数

| 参数 | 类型 | 必填 | 默认值 | 说明 |
|------|------|------|--------|------|
| `equation` | string | 是 | - | 方程表达式 |

### 支持的方程类型

- 一元一次方程：`2x + 3 = 7`
- 一元二次方程：`x^2 - 5x + 6 = 0`
- 简单线性方程组

### 返回值

```python
{
    "solution": 2,
    "equation": "2x + 3 = 7",
    "type": "linear"
}
```

### 示例

```python
solve_equation(equation="2x + 3 = 7")
solve_equation(equation="x^2 - 5x + 6 = 0")
```

---

## auto_sum

自动求和（在范围内设置 SUM 公式）。

### 参数

| 参数 | 类型 | 必填 | 默认值 | 说明 |
|------|------|------|--------|------|
| `file` | string | 是 | - | Excel 文件路径 |
| `sheet` | string | 是 | - | 工作表名称 |
| `range` | string | 是 | - | 数据范围（A1 表示法） |
| `direction` | string | 否 | "down" | 求和方向：`down`/`right` |

### 返回值

```python
{
    "success": True,
    "range": "A1:D10",
    "direction": "down",
    "formulas_set": 4
}
```

### 示例

```python
auto_sum(file="data.xlsx", sheet="Sheet1", range="A1:D10", direction="down")
auto_sum(file="data.xlsx", sheet="Sheet1", range="A1:D10", direction="right")
```
