---
name: abacus-field
description: "方田章 - Excel 数据读取与单元格操作。当需要读取 Excel 数据、查看工作表结构、管理单元格时使用。"
version: 0.2.0
chapter: field
level: rod
tags: [excel, read, data, cell, worksheet, range, table]
---

# 方田章 - 数据读取与单元格操作

## CRITICAL RULES
1. 始终指定完整文件路径
2. 工作表名称区分大小写
3. 范围使用 A1 表示法（如 `A1:D10`，不是 `A:D`）
4. 写入前检查合并单元格
5. 修改后必须保存 - openpyxl 变更仅在内存中

## 能力一览

| 能力 | CLI 命令 | 说明 |
|------|----------|------|
| measure_range | `abacus read` | 读取范围数据 |
| measure_cells | `abacus cells` | 读取单元格详情（值、公式、样式） |
| measure_structure | `abacus structure` | 查看工作表结构 |
| manage_named_range | `abacus manage-named-range` | 管理命名范围 |

---

## Range 引用参考

### 格式化工具选择

| 用途 | 工具 | 参数 | 场景 |
|------|------|------|------|
| 数字显示格式 | `abacus_convert_format` | format_type | 日期、货币、百分比 |
| 类型转换 | `abacus_convert_type` | target_type | 文本转数值等 |
| 单元格格式 | `abacus_format_range` | font/fill/border | 加粗、颜色、边框 |
| 自动列宽 | `abacus_manage_size` | auto | 适应内容 |
| 语义样式 | `abacus_manage_style` | apply_header/apply_kpi | 表头、KPI |
| 查找替换 | `abacus_find_replace` | find/replace | 文本替换 |

### 快速模式：写入 → 格式化 → 自适应

```
1. abacus_batch_execute (写入数据)
2. abacus_convert_format (应用数字格式)
3. abacus_manage_size (自动列宽)
```

### 快速模式：带填充色的表头行

```python
abacus_format_range(
    file='data.xlsx',
    sheet='Sheet1',
    range='A1:D1',
    font={'bold': True, 'color': 'FFFFFF'},
    fill={'color': '4472C4', 'pattern_type': 'solid'},
    alignment={'horizontal': 'center'}
)
```

### 格式代码速查

| 类型 | 代码 | 示例 |
|------|------|------|
| 数字 | `#,##0.00` | 1,234.56 |
| 美元 | `$#,##0.00` | $1,234.56 |
| 欧元 | `€#,##0.00` | €1,234.56 |
| 人民币 | `¥#,##0` | ¥1,235 |
| 百分比 | `0.00%` | 12.34% |
| 日期(ISO) | `yyyy-mm-dd` | 2023-03-15 |
| 日期(美式) | `mm/dd/yyyy` | 03/15/2023 |
| 日期(欧式) | `dd/mm/yyyy` | 15/03/2023 |
| 时间 | `h:mm AM/PM` | 2:30 PM |
| 时间(24h) | `hh:mm:ss` | 14:30:00 |
| 文本 | `@` | 原样显示 |

### format-range 属性速查

| 属性 | 类型 | 示例 |
|------|------|------|
| `bold` | bool | `True` |
| `italic` | bool | `True` |
| `underline` | bool | `True` |
| `size` | number | `14` |
| `name` | string | `"Calibri"` |
| `color` (font) | hex | `"FFFFFF"` |
| `color` (fill) | hex | `"4472C4"` |
| `horizontal` | string | `"center"`, `"left"`, `"right"` |
| `vertical` | string | `"middle"`, `"top"`, `"bottom"` |
| `wrap_text` | bool | `True` |
| `style` (border) | string | `"thin"`, `"medium"`, `"thick"` |

---

## Worksheet 工作表操作

### 工作表生命周期

| 操作 | 工具 | 关键参数 |
|------|------|----------|
| 创建工作表 | `abacus_batch_execute` | 写入新工作表名称 |
| 读取结构 | `abacus_measure_structure` | file, sheet |
| 复制范围 | `abacus_copy_range` | source, target |
| 删除内容 | `abacus_clear_range` | range, clear_type |

### 工作表可见性

| 操作 | 工具 | 说明 |
|------|------|------|
| 显示 | `abacus_manage_sheet_visibility` | action='show' |
| 隐藏 | `abacus_manage_sheet_visibility` | action='hide' |
| 非常隐藏 | `abacus_manage_sheet_visibility` | action='very-hide' |
| 获取状态 | `abacus_manage_sheet_visibility` | action='get' |

### 工作表样式

| 操作 | 工具 | 说明 |
|------|------|------|
| 设置标签颜色 | `abacus_manage_sheet_style` | action='set', color='FF0000' |
| 获取标签颜色 | `abacus_manage_sheet_style` | action='get' |
| 清除标签颜色 | `abacus_manage_sheet_style` | action='clear' |

### 冻结窗格

```python
abacus_freeze_panes(
    file='data.xlsx',
    sheet='Sheet1',
    rows=1,           # 冻结首行
    columns=1         # 冻结首列
)
```

### 常用操作示例

```
# 读取工作表结构
abacus_measure_structure(file='data.xlsx', sheet='Sheet1')

# 跨工作表复制范围
abacus_copy_range(
    file='data.xlsx',
    sheet='Source',
    range='A1:D10',
    target='Destination!A1'
)

# 清除工作表内容
abacus_clear_range(
    file='data.xlsx',
    sheet='Sheet1',
    range='A1:Z100',
    clear_type='contents'  # 或 'formats' 或 'all'
)
```

### 常见错误

| 错误 | 原因 | 解决方案 |
|------|------|----------|
| "Sheet not found" | 工作表名拼写错误 | 用 `abacus_measure_structure` 查看可用工作表 |
| "Range invalid" | A1 表示法错误 | 用 `abacus_measure_range` 检查 |
| "File not found" | 路径错误 | 验证文件路径 |
| "Merged cell conflict" | 范围与合并单元格冲突 | 用 `abacus_measure_structure` 检查合并单元格 |

---

## Table 表格操作

### 表格操作一览

| 操作 | 说明 | 工具 |
|------|------|------|
| create | 从范围创建新表格 | `abacus_manage_table` |
| list | 列出所有表格 | `abacus_manage_table` |
| delete | 删除表格（保留数据） | `abacus_manage_table` |
| append | 向表格追加行 | `abacus_manage_table` |

### 创建表格

```python
abacus_manage_table(
    file='data.xlsx',
    sheet='Sheet1',
    action='create',
    table_name='SalesData',
    range='A1:D100',
    style='TableStyleMedium2'
)
```

### 表格样式

| 样式 | 说明 |
|------|------|
| TableStyleLight1 | 最小边框，无表头填充 |
| TableStyleLight2-21 | 各种浅色主题 |
| TableStyleMedium2 | 标准蓝色，最常用 |
| TableStyleMedium9 | 橙色强调 |
| TableStyleDark1-11 | 深色表头，白色文字 |

### 追加数据到表格

```python
abacus_manage_table(
    file='data.xlsx',
    sheet='Sheet1',
    action='append',
    table_name='SalesData',
    data=[
        {'Name': 'Product A', 'Sales': 100},
        {'Name': 'Product B', 'Sales': 200}
    ]
)
```

### 为什么使用表格而非普通范围

- 结构化引用：`=SUM(Sales[Amount])` 而非 `=SUM(B2:B100)`
- 添加行时自动扩展
- 内置筛选、排序和隔行色
- Power Query 命名引用：`Excel.CurrentWorkbook(){[Name="SalesData"]}`

### 何时不使用表格

- 单元格参数（使用命名范围）
- 带合并单元格的布局区域
- 特定间距的打印格式报表

---

## 最佳实践

1. **操作前验证工作表存在**
2. **使用 A1 表示法**（如 `A1:D10`，不是 `A:D`）
3. **写入前检查合并单元格**
4. **修改后保存** - openpyxl 变更仅在内存中
5. **大数据集使用冻结窗格** 保持表头可见

## 示例

```bash
# 读取范围数据
abacus read -f data.xlsx -s Sheet1 -r A1:D10

# 查看工作表结构
abacus structure -f data.xlsx

# 读取单元格详情
abacus cells -f data.xlsx -s Sheet1 -r A1:B5

# 管理命名范围
abacus manage-named-range -f data.xlsx --action create --name SalesData --ref Sheet1!$A$1:$D$100
```
