# 商功章 (Work) - 批量操作与图表管理

> 商功章负责批量操作、图表管理、格式化、筛选、保护、报告生成和模板填充。

---

## 工具列表

### 批量操作

| 工具 | 描述 |
|------|------|
| `batch_execute` | 批量执行多个操作 |
| `batch_transform` | 批量转换 |
| `batch_validate` | 批量验证 |

### 图表管理

| 工具 | 描述 |
|------|------|
| `create_chart` | 创建图表（柱形图、折线图、饼图、面积图、散点图） |
| `update_chart` | 更新图表标题 |
| `list_charts` | 列出所有图表 |
| `delete_chart` | 删除图表 |
| `create_advanced_chart` | 创建高级图表（组合图、双轴图、瀑布图、甘特图） |
| `export_chart_as_image` | 导出图表为图片 |

### 格式化

| 工具 | 描述 |
|------|------|
| `format_range` | 格式化单元格（字体、颜色、边框、条件格式） |
| `manage_style` | 管理样式（行业品牌色、表头、KPI 格式） |

### 筛选

| 工具 | 描述 |
|------|------|
| `set_auto_filter` | 设置自动筛选 |
| `advanced_filter` | 高级筛选（支持复杂条件） |

### 表格与视图

| 工具 | 描述 |
|------|------|
| `manage_table` | 管理 Excel 表格 |
| `manage_comment` | 批注管理 |
| `manage_data_view` | 数据视图管理 |

### 布局

| 工具 | 描述 |
|------|------|
| `freeze_panes` | 冻结窗格 |
| `manage_row_column_visibility` | 管理行列可见性 |
| `group_rows` | 分组行（折叠/展开） |
| `set_print_area` | 设置打印区域 |
| `set_zoom` | 控制缩放 |

### 报告生成

| 工具 | 描述 |
|------|------|
| `create_basic_report` | 生成基础报表 |
| `create_advanced_report` | 生成高级报表 |
| `fill_template` | 基于模板填充数据 |
| `create_mapping_template` | 创建数据映射模板 |
| `generate_summary_report` | 数据摘要报告 |
| `generate_diff_report` | 变化检测报告 |

### 拆分与保护

| 工具 | 描述 |
|------|------|
| `split_sheet` | 拆分工作表 |
| `protect_workbook` | 保护工作簿 |
| `protect_sheet` | 保护工作表 |
| `unprotect_sheet` | 解除工作表保护 |

### 其他

| 工具 | 描述 |
|------|------|
| `insert_excel_image` | 插入图片 |
| `pack_file` | 打包为 ZIP |
| `unpack_file` | 解包 ZIP |

---

## batch_execute

批量执行多个操作。

### 参数

| 参数 | 类型 | 必填 | 默认值 | 说明 |
|------|------|------|--------|------|
| `file` | string | 是 | - | Excel 文件路径 |
| `operations` | list | 是 | - | 操作列表 |

#### operations 示例

```python
[
    {"type": "write", "sheet": "Sheet1", "cell": "A1", "value": "Hello"},
    {"type": "format", "sheet": "Sheet1", "range": "A1:D1", "font": {"bold": true}},
    {"type": "merge", "sheet": "Sheet1", "range": "A1:D1"}
]
```

### 返回值

```python
{
    "success": True,
    "results": [
        {"type": "write", "success": True},
        {"type": "format", "success": True}
    ]
}
```

### 示例

```python
batch_execute(
    file="data.xlsx",
    operations=[
        {"type": "write", "sheet": "Sheet1", "cell": "A1", "value": "Test"},
        {"type": "format", "sheet": "Sheet1", "range": "A1:D1", "font": {"bold": True}}
    ]
)
```

---

## create_chart

创建图表。

### 参数

| 参数 | 类型 | 必填 | 默认值 | 说明 |
|------|------|------|--------|------|
| `file` | string | 是 | - | Excel 文件路径 |
| `sheet` | string | 是 | - | 工作表名称 |
| `range` | string | 是 | - | 数据范围（A1 表示法） |
| `chart_type` | string | 是 | - | 图表类型：`bar`/`line`/`pie`/`area`/`scatter` |
| `title` | string | 否 | None | 图表标题 |
| `x_axis` | string | 否 | None | X 轴标题 |
| `y_axis` | string | 否 | None | Y 轴标题 |
| `output_sheet` | string | 否 | None | 输出工作表名称 |
| `position` | string | 否 | "A1" | 图表位置 |
| `width` | float | 否 | 15 | 图表宽度 |
| `height` | float | 否 | 10 | 图表高度 |

### 返回值

```python
{
    "success": True,
    "chart_index": 0,
    "chart_type": "bar"
}
```

### 示例

```python
create_chart(file="data.xlsx", sheet="Sales", range="A1:C10", chart_type="bar", title="销售趋势")
create_chart(file="data.xlsx", sheet="Data", range="A1:B20", chart_type="line", y_axis="数量")
```

---

## create_advanced_chart

创建高级图表（组合图、双轴图、瀑布图、甘特图）。

### 参数

| 参数 | 类型 | 必填 | 默认值 | 说明 |
|------|------|------|--------|------|
| `file` | string | 是 | - | 输出文件路径 |
| `data` | dict | 是 | - | 图表数据（包含 headers 和 rows） |
| `chart_type` | string | 是 | - | 图表类型：`combo`/`dual_axis`/`waterfall`/`gantt` |
| `title` | string | 否 | None | 图表标题 |
| `x_axis` | string | 否 | None | X轴标题 |
| `y_axis` | string | 否 | None | Y轴标题 |

### 返回值

```python
{
    "file": "chart.xlsx",
    "chart_type": "combo",
    "title": "销售趋势",
    "created": True
}
```

### 示例

```python
create_advanced_chart(
    file="chart.xlsx",
    data={
        "headers": ["月份", "销售额", "利润"],
        "rows": [["1月", 100, 20], ["2月", 120, 25]]
    },
    chart_type="combo",
    title="销售趋势"
)
```

---

## format_range

格式化单元格（字体、颜色、边框、条件格式等）。

### 参数

| 参数 | 类型 | 必填 | 默认值 | 说明 |
|------|------|------|--------|------|
| `file` | string | 是 | - | Excel 文件路径 |
| `sheet` | string | 是 | - | 工作表名称 |
| `range` | string | 是 | - | 数据范围（A1 表示法） |
| `font` | dict | 否 | None | 字体设置 |
| `fill` | dict | 否 | None | 填充设置 |
| `border` | dict | 否 | None | 边框设置 |
| `alignment` | dict | 否 | None | 对齐设置 |
| `number_format` | string | 否 | None | 数字格式 |
| `conditional` | dict | 否 | None | 条件格式设置 |

#### font 示例

```python
{"name": "Arial", "size": 12, "bold": True, "color": "000000"}
```

#### fill 示例

```python
{"color": "FFFF00", "pattern_type": "solid"}
```

#### border 示例

```python
{"style": "thin", "color": "000000"}
```

#### alignment 示例

```python
{"horizontal": "center", "vertical": "center", "wrap_text": True}
```

#### conditional 示例

```python
{"type": "cell", "operator": "greaterThan", "value": 100}
```

### 返回值

```python
{
    "success": True,
    "formatted_range": "A1:D10"
}
```

### 示例

```python
format_range(
    file="data.xlsx",
    sheet="Sheet1",
    range="A1:D1",
    font={"bold": True, "size": 14},
    fill={"color": "4472C4"},
    alignment={"horizontal": "center"}
)
```

---

## create_pivot

创建数据透视表。

### 参数

| 参数 | 类型 | 必填 | 默认值 | 说明 |
|------|------|------|--------|------|
| `file` | string | 是 | - | Excel 文件路径 |
| `sheet` | string | 是 | - | 源数据工作表 |
| `range` | string | 是 | - | 源数据范围（A1 表示法） |
| `row_fields` | list | 是 | - | 行字段列表 |
| `value_field` | string | 是 | - | 值字段 |
| `agg_function` | string | 否 | "sum" | 聚合函数：`sum`/`avg`/`count`/`min`/`max` |
| `output_sheet` | string | 否 | None | 输出工作表名称 |

### 返回值

```python
{
    "success": True,
    "output_sheet": "PivotResult",
    "rows": 50
}
```

### 示例

```python
create_pivot(
    file="data.xlsx",
    sheet="Sheet1",
    range="A1:D100",
    row_fields=["Category", "Region"],
    value_field="Sales",
    agg_function="sum"
)
```

---

## freeze_panes

冻结窗格。

### 参数

| 参数 | 类型 | 必填 | 默认值 | 说明 |
|------|------|------|--------|------|
| `file` | string | 是 | - | Excel 文件路径 |
| `sheet` | string | 是 | - | 工作表名称 |
| `rows` | int | 否 | None | 冻结行数 |
| `columns` | int | 否 | None | 冻结列数 |
| `cell` | string | 否 | None | 冻结位置（如 "B2"） |

### 返回值

```python
{
    "action": "freeze",
    "rows": 1
}
```

### 示例

```python
freeze_panes(file="data.xlsx", sheet="Sheet1", rows=1)
freeze_panes(file="data.xlsx", sheet="Sheet1", cell="B2")
freeze_panes(file="data.xlsx", sheet="Sheet1")  # 解除冻结
```

---

## set_auto_filter

设置自动筛选。

### 参数

| 参数 | 类型 | 必填 | 默认值 | 说明 |
|------|------|------|--------|------|
| `file` | string | 是 | - | Excel 文件路径 |
| `sheet` | string | 是 | - | 工作表名称 |
| `action` | string | 是 | - | 操作：`set`/`remove`/`get` |
| `range` | string | 条件 | None | 筛选范围（set 时必填） |
| `column` | string | 否 | None | 筛选列 |
| `criteria` | string | 否 | None | 筛选条件 |

### 返回值

```python
{
    "action": "set",
    "range": "A1:D100"
}
```

### 示例

```python
set_auto_filter(file="data.xlsx", sheet="Sheet1", action="set", range="A1:D100")
set_auto_filter(file="data.xlsx", sheet="Sheet1", action="get")
set_auto_filter(file="data.xlsx", sheet="Sheet1", action="remove")
```

---

## manage_table

管理 Excel 表格（创建/列出/删除/追加）。

### 参数

| 参数 | 类型 | 必填 | 默认值 | 说明 |
|------|------|------|--------|------|
| `file` | string | 是 | - | Excel 文件路径 |
| `sheet` | string | 是 | - | 工作表名称 |
| `action` | string | 是 | - | 操作：`create`/`list`/`delete`/`append` |
| `table_name` | string | 条件 | None | 表格名称 |
| `range` | string | 条件 | None | 数据范围（create 时必填） |
| `style` | string | 否 | None | 表格样式 |
| `data` | list | 条件 | None | 追加数据（append 时必填） |

### 返回值

```python
# action=list
{
    "tables": [
        {"name": "SalesTable", "range": "A1:D100"}
    ]
}
```

### 示例

```python
manage_table(file="data.xlsx", sheet="Sheet1", action="create", table_name="SalesTable", range="A1:D100")
manage_table(file="data.xlsx", sheet="Sheet1", action="list")
manage_table(file="data.xlsx", sheet="Sheet1", action="append", table_name="SalesTable", data=[{"Name": "A", "Sales": 100}])
```

---

## create_basic_report

从数据生成基础 Excel 报表。

### 参数

| 参数 | 类型 | 必填 | 默认值 | 说明 |
|------|------|------|--------|------|
| `data_source` | string | 是 | - | CSV 文件路径 |
| `output` | string | 是 | - | 输出文件路径 |
| `sheet_name` | string | 否 | "Data" | 工作表名称 |
| `title` | string | 否 | None | 报表标题 |

### 返回值

```python
{
    "success": True,
    "output": "report.xlsx",
    "rows": 100
}
```

### 示例

```python
create_basic_report(data_source="data.csv", output="report.xlsx", title="月度销售报表")
```

---

## create_advanced_report

生成高级 Excel 报表。

### 参数

| 参数 | 类型 | 必填 | 默认值 | 说明 |
|------|------|------|--------|------|
| `data_source` | string | 是 | - | CSV 文件路径 |
| `output` | string | 是 | - | 输出文件路径 |
| `chart_type` | string | 否 | "bar" | 图表类型：`bar`/`line`/`pie` |
| `include_dashboard` | bool | 否 | True | 是否包含仪表板 |

### 返回值

```python
{
    "success": True,
    "output": "advanced_report.xlsx",
    "charts_created": 3
}
```

### 示例

```python
create_advanced_report(data_source="data.csv", output="report.xlsx", chart_type="bar", include_dashboard=True)
```

---

## fill_template

基于模板填充数据生成报表。

### 参数

| 参数 | 类型 | 必填 | 默认值 | 说明 |
|------|------|------|--------|------|
| `template` | string | 是 | - | 模板文件路径 |
| `output` | string | 是 | - | 输出文件路径 |
| `data_source` | string | 否 | None | CSV 数据源 |
| `data` | dict | 否 | None | 填充数据（字典格式） |
| `sheet_name` | string | 否 | None | 工作表名称 |
| `start_cell` | string | 否 | "A1" | 起始单元格 |

### 返回值

```python
{
    "success": True,
    "output": "output.xlsx",
    "cells_filled": 15
}
```

### 示例

```python
fill_template(
    template="template.xlsx",
    output="output.xlsx",
    data={"Name": "John", "Sales": 1000}
)
fill_template(
    template="template.xlsx",
    output="output.xlsx",
    data_source="data.csv"
)
```

---

## protect_workbook

保护工作簿。

### 参数

| 参数 | 类型 | 必填 | 默认值 | 说明 |
|------|------|------|--------|------|
| `file` | string | 是 | - | Excel 文件路径 |
| `password` | string | 否 | None | 保护密码 |

### 返回值

```python
{
    "success": True,
    "action": "protect_workbook",
    "file": "data.xlsx"
}
```

### 示例

```python
protect_workbook(file="data.xlsx")
protect_workbook(file="data.xlsx", password="mypassword")
```

---

## protect_sheet

保护工作表。

### 参数

| 参数 | 类型 | 必填 | 默认值 | 说明 |
|------|------|------|--------|------|
| `file` | string | 是 | - | Excel 文件路径 |
| `sheet` | string | 是 | - | 工作表名称 |
| `password` | string | 否 | None | 保护密码 |

### 返回值

```python
{
    "success": True,
    "action": "protect_sheet",
    "sheet": "Sheet1"
}
```

### 示例

```python
protect_sheet(file="data.xlsx", sheet="Sheet1")
protect_sheet(file="data.xlsx", sheet="Sheet1", password="mypassword")
```

---

## unprotect_sheet

解除工作表保护。

### 参数

| 参数 | 类型 | 必填 | 默认值 | 说明 |
|------|------|------|--------|------|
| `file` | string | 是 | - | Excel 文件路径 |
| `sheet` | string | 是 | - | 工作表名称 |
| `password` | string | 否 | None | 保护密码 |

### 返回值

```python
{
    "success": True,
    "action": "unprotect_sheet",
    "sheet": "Sheet1"
}
```

### 示例

```python
unprotect_sheet(file="data.xlsx", sheet="Sheet1")
unprotect_sheet(file="data.xlsx", sheet="Sheet1", password="mypassword")
```

---

## split_sheet

将 Excel 工作表按条件拆分为多个文件。

### 参数

| 参数 | 类型 | 必填 | 默认值 | 说明 |
|------|------|------|--------|------|
| `file` | string | 是 | - | Excel 文件路径 |
| `sheet` | string | 是 | - | 工作表名称 |
| `output_dir` | string | 是 | - | 输出目录 |
| `split_by` | string | 是 | - | 拆分方式：`column`/`row_count`/`range` |
| `split_column` | string | 条件 | None | 拆分列名（column 时必填） |
| `row_count` | int | 条件 | None | 每文件行数（row_count 时必填） |
| `prefix` | string | 否 | "split" | 输出文件前缀 |

### 返回值

```python
{
    "success": True,
    "files_created": 5,
    "output_dir": "output"
}
```

### 示例

```python
split_sheet(
    file="data.xlsx",
    sheet="Sheet1",
    output_dir="output",
    split_by="column",
    split_column="Category"
)
split_sheet(
    file="data.xlsx",
    sheet="Sheet1",
    output_dir="output",
    split_by="row_count",
    row_count=1000
)
```

---

## manage_comment

批注管理（添加、删除、获取批注）。

### 参数

| 参数 | 类型 | 必填 | 默认值 | 说明 |
|------|------|------|--------|------|
| `file` | string | 是 | - | Excel 文件路径 |
| `sheet` | string | 是 | - | 工作表名称 |
| `action` | string | 是 | - | 操作：`add`/`delete`/`get`/`list` |
| `cell` | string | 条件 | None | 单元格位置 |
| `text` | string | 条件 | None | 批注内容（add 时必填） |
| `author` | string | 否 | "Abacus" | 批注作者 |

### 返回值

```python
# action=list
{
    "comments": [
        {"cell": "A1", "text": "重要数据", "author": "张三"}
    ]
}
```

### 示例

```python
manage_comment(file="data.xlsx", sheet="Sheet1", action="add", cell="A1", text="重要数据")
manage_comment(file="data.xlsx", sheet="Sheet1", action="list")
manage_comment(file="data.xlsx", sheet="Sheet1", action="delete", cell="A1")
```

---

## generate_summary_report

数据摘要报告（自动生成数据摘要）。

### 参数

| 参数 | 类型 | 必填 | 默认值 | 说明 |
|------|------|------|--------|------|
| `file` | string | 是 | - | Excel 文件路径 |
| `sheet` | string | 是 | - | 工作表名称 |

### 返回值

```python
{
    "total_rows": 1000,
    "total_columns": 10,
    "columns": ["ID", "Name", "Sales", ...],
    "dtypes": {"ID": "int64", "Name": "object", "Sales": "float64"},
    "null_counts": {"ID": 0, "Name": 2, "Sales": 5},
    "numeric_stats": {"Sales": {"mean": 5000, "std": 2000, "min": 100, "max": 15000}},
    "categorical_stats": {"Name": {"unique": 500, "top": "Product A"}}
}
```

### 示例

```python
generate_summary_report(file="data.xlsx", sheet="Sheet1")
```

---

## generate_diff_report

变化检测报告（对比两个版本的数据）。

### 参数

| 参数 | 类型 | 必填 | 默认值 | 说明 |
|------|------|------|--------|------|
| `old_file` | string | 是 | - | 旧版本文件路径 |
| `old_sheet` | string | 是 | - | 旧版本工作表名称 |
| `new_file` | string | 是 | - | 新版本文件路径 |
| `new_sheet` | string | 是 | - | 新版本工作表名称 |
| `key_columns` | list | 否 | None | 用于匹配的键列 |

### 返回值

```python
{
    "old_rows": 100,
    "new_rows": 120,
    "row_diff": 20,
    "added_columns": ["NewCol"],
    "removed_columns": [],
    "changes": [...]
}
```

### 示例

```python
generate_diff_report(
    old_file="old.xlsx", old_sheet="Sheet1",
    new_file="new.xlsx", new_sheet="Sheet1",
    key_columns=["ID"]
)
```
