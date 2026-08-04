---
name: abacus-work
description: "商功章 - 批量操作与图表管理。当需要批量合并、写入、验证，或创建图表、管理格式时使用。"
version: 0.2.0
chapter: work
level: rod
tags: [excel, batch, chart, format, pivot, table, workflow]
---

# 商功章 - 批量操作与图表管理

## CRITICAL RULES
1. 批量写入时使用手动计算模式，避免重复计算
2. 多图表布局必须使用 position 参数避免重叠
3. 删除前验证对象存在 - openpyxl 删除操作不可逆
4. 修改后必须保存 - openpyxl 变更仅在内存中

## 能力一览

| 能力 | CLI 命令 | 说明 |
|------|----------|------|
| batch_execute | `abacus batch` | 批量执行多个操作 |
| batch_transform | `abacus batch-transform` | 批量转换 |
| batch_validate | `abacus batch-validate` | 批量验证 |
| create_chart | `abacus create-chart` | 创建图表（柱形、折线、饼图等） |
| create_advanced_chart | `abacus create-advanced-chart` | 创建高级图表（组合图、双轴图） |
| update_chart | `abacus update-chart` | 更新图表 |
| delete_chart | `abacus delete-chart` | 删除图表 |
| list_charts | `abacus list-charts` | 列出所有图表 |
| create_pivot | `abacus create-pivot` | 创建数据透视表 |
| manage_table | `abacus manage-table` | 管理 Excel 表格 |
| format_range | `abacus format` | 格式化单元格 |
| manage_comment | `abacus comment` | 批注管理 |
| freeze_panes | `abacus freeze` | 冻结窗格 |
| set_auto_filter | `abacus set-auto-filter` | 设置自动筛选 |
| manage_row_column_visibility | `abacus manage-visibility` | 管理行列可见性 |
| group_rows | `abacus group-rows` | 分组行 |
| protect_workbook | `abacus protect-workbook` | 保护工作簿 |
| protect_sheet | `abacus protect-sheet` | 保护工作表 |
| unprotect_sheet | `abacus unprotect-sheet` | 解除工作表保护 |
| set_array_formula | `abacus set-array-formula` | 设置数组公式 |
| insert_excel_image | `abacus insert-image` | 插入图片 |
| export_chart_as_image | `abacus export-chart-image` | 导出图表为图片 |
| pack_file | `abacus pack-file` | 打包为 ZIP |
| unpack_file | `abacus unpack-file` | 解包 ZIP |
| set_print_area | `abacus set-print-area` | 设置打印区域 |
| set_zoom | `abacus set-zoom` | 控制缩放 |
| create_basic_report | `abacus create-basic-report` | 生成基础报表 |
| create_advanced_report | `abacus create-advanced-report` | 生成高级报表 |
| fill_template | `abacus fill-template` | 基于模板填充数据 |

---

## 批量操作参考

### batch_execute 操作类型

| 类型 | 说明 | 参数 |
|------|------|------|
| merge | 合并单元格 | sheet, range |
| unmerge | 取消合并 | sheet, range |
| write | 写入单元格 | sheet, cell, value |
| style | 应用样式 | sheet, range, font/fill/border |

### batch_transform 操作类型

| 类型 | 说明 | 参数 |
|------|------|------|
| convert_type | 类型转换 | range, target |
| convert_format | 格式转换 | range, target |

### batch_validate 操作类型

| 类型 | 说明 | 参数 |
|------|------|------|
| validate_range | 范围验证 | range, min, max |
| validate_range | 类型验证 | range, expected |

### 使用场景

```bash
# 批量执行多个操作
abacus batch -f data.xlsx --operations '[
  {"type": "write", "sheet": "Sheet1", "cell": "A1", "value": "Hello"},
  {"type": "format", "sheet": "Sheet1", "range": "A1:D1", "font": {"bold": true}}
]'

# 批量转换
abacus batch-transform -f data.xlsx --operations '[
  {"type": "convert_type", "range": "A1:A10", "target": "float"},
  {"type": "convert_format", "range": "B1:B10", "target": "number"}
]'

# 批量验证
abacus batch-validate -f data.xlsx --operations '[
  {"type": "validate_range", "range": "A1:A10", "min": 0, "max": 100},
  {"type": "validate_type", "range": "B1:B10", "expected": "float"}
]'
```

---

## 图表参考

### 图表类型

| 类型 | 值 | 使用场景 |
|------|-----|----------|
| Column Clustered | `bar` | 比较类别 |
| Line | `line` | 趋势变化 |
| Pie | `pie` | 占比构成 |
| Area | `area` | 累计趋势 |
| Scatter | `scatter` | 变量相关性 |

### 图表创建

```python
# 从范围创建
abacus_create_chart(
    file='data.xlsx',
    sheet='Sheet1',
    range='A1:C10',
    chart_type='bar',
    title='月度销售',
    x_axis='月份',
    y_axis='销售额'
)

# 从透视表创建
abacus_create_chart(
    file='data.xlsx',
    sheet='PivotResult',
    range='A1:D20',
    chart_type='line'
)
```

### 多图表布局

```python
# 数据在 A1:D10，在下方放置 4 个图表（2×2 网格）
abacus_create_chart(..., position='A12')   # 左上
abacus_create_chart(..., position='G12')   # 右上
abacus_create_chart(..., position='A27')   # 左下
abacus_create_chart(..., position='G27')   # 右下
```

### 图表管理

```python
# 列出所有图表
abacus_list_charts(file='data.xlsx', sheet='Sheet1')

# 更新图表标题
abacus_update_chart(file='data.xlsx', sheet='Sheet1', chart_index=0, title='新标题')

# 删除图表
abacus_delete_chart(file='data.xlsx', sheet='Sheet1', chart_index=0)
```

### 高级图表

```python
# 组合图（柱形图+折线图）
abacus_create_advanced_chart(
    file='output.xlsx',
    data={'headers': ['月份', '销售额', '利润'], 'rows': [['1月', 100, 20], ['2月', 120, 25]]},
    chart_type='combo',
    title='销售趋势'
)

# 双轴图
abacus_create_advanced_chart(
    file='output.xlsx',
    data={'headers': ['月份', '销售额', '增长率'], 'rows': [['1月', 100, 0.1], ['2月', 120, 0.2]]},
    chart_type='dual_axis',
    title='销售与增长'
)
```

---

## 格式化参考

### 格式化工具选择

| 用途 | 工具 | 场景 |
|------|------|------|
| 数字显示格式 | `abacus_convert_format` | 日期、货币、百分比 |
| 类型转换 | `abacus_convert_type` | 文本转数值等 |
| 单元格格式 | `abacus_format_range` | 加粗、颜色、边框 |
| 自动列宽 | `abacus_manage_size` | 适应内容 |
| 语义样式 | `abacus_manage_style` | 表头、KPI |

### 快速模式：写入 → 格式化 → 自适应

```
1. abacus_batch_execute (写入数据)
2. abacus_convert_format (应用数字格式)
3. abacus_manage_size (自动列宽)
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

## 表格操作参考

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

### 为什么使用表格而非普通范围

- 结构化引用：`=SUM(Sales[Amount])` 而非 `=SUM(B2:B100)`
- 添加行时自动扩展
- 内置筛选、排序和隔行色
- Power Query 命名引用：`Excel.CurrentWorkbook(){[Name="SalesData"]}`

---

## 已知限制和陷阱

### 透视表格式丢失
**问题：** 透视表刷新后格式可能丢失
**解决方案：** 使用透视表内置样式（不手动格式化）

### 隐藏对象
**问题：** 隐藏的工作表/行/列可能影响操作
**解决方案：** 操作前检查对象可见性，使用 `skip_hidden` 参数

### 超时问题
**问题：** 大数据集操作可能超时
**解决方案：** 分批处理大数据集，使用手动计算模式

### 公式计算顺序
**问题：** 公式依赖关系可能导致计算错误
**解决方案：** 使用手动计算模式，按依赖顺序写入公式

### 合并单元格陷阱
**问题：** 合并单元格可能导致排序/筛选失败
**解决方案：** 避免在数据区域使用合并单元格，操作前取消合并

---

## 最佳实践

1. **批量操作替代逐单元格** - 提高性能，减少错误
2. **使用 Excel 表格** - 自动扩展，结构化引用
3. **明确图表定位** - 多图表布局使用 position 参数
4. **验证后再删除** - openpyxl 删除操作不可逆
5. **使用手动计算模式** - 批量写入时避免重复计算

## 示例

```bash
# 批量执行操作
abacus batch -f data.xlsx --operations '[{"type": "write", "sheet": "Sheet1", "cell": "A1", "value": "Hello"}]'

# 创建图表
abacus create-chart -f data.xlsx -s Sheet1 -r A1:C10 --chart-type bar --title "销售趋势"

# 格式化范围
abacus format -f data.xlsx -s Sheet1 -r A1:D1 --font '{"bold": true}' --fill '{"color": "4472C4"}'

# 冻结窗格
abacus freeze -f data.xlsx -s Sheet1 --rows 1

# 创建透视表
abacus create-pivot -f data.xlsx -s Sheet1 -r A1:D100 --row-fields Region --value-field Sales --agg-function sum
```