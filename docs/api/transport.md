# 均输章 (Transport) - 导入导出与数据迁移

> 均输章负责数据导入导出、文件迁移、合并和格式转换。

---

## 工具列表

| 工具 | 描述 |
|------|------|
| `import_data` | 导入数据（CSV/JSON → Excel） |
| `export_data` | 导出数据（Excel → CSV/JSON） |
| `migrate` | 数据迁移 |
| `merge_files` | 合并多个文件 |
| `join_tables` | SQL 风格关联 |
| `batch_merge` | 多表批量合并 |
| `excel_to_markdown` | Excel 转 Markdown |

---

## import_data

导入数据。

### 参数

| 参数 | 类型 | 必填 | 默认值 | 说明 |
|------|------|------|--------|------|
| `file` | string | 是 | - | 目标 Excel 文件路径 |
| `source` | string | 是 | - | 源文件路径 |
| `source_type` | string | 否 | "csv" | 源文件类型：`csv`/`json` |
| `sheet` | string | 否 | "Sheet1" | 目标工作表 |

### 返回值

```python
{
    "success": True,
    "rows_imported": 1000,
    "columns_imported": 10
}
```

### 示例

```python
import_data(file="output.xlsx", source="data.csv", source_type="csv", sheet="Sales")
import_data(file="output.xlsx", source="data.json", source_type="json")
```

---

## export_data

导出数据。

### 参数

| 参数 | 类型 | 必填 | 默认值 | 说明 |
|------|------|------|--------|------|
| `file` | string | 是 | - | 源 Excel 文件路径 |
| `sheet` | string | 是 | - | 工作表名称 |
| `range` | string | 是 | - | 数据范围（A1 表示法） |
| `output` | string | 是 | - | 输出文件路径 |
| `format` | string | 否 | "csv" | 输出格式：`csv`/`json` |

### 返回值

```python
{
    "success": True,
    "rows_exported": 100,
    "columns_exported": 8
}
```

### 示例

```python
export_data(file="data.xlsx", sheet="Sheet1", range="A1:D100", output="output.csv")
export_data(file="data.xlsx", sheet="Sales", range="A1:C50", output="data.json", format="json")
```

---

## migrate

数据迁移。

### 参数

| 参数 | 类型 | 必填 | 默认值 | 说明 |
|------|------|------|--------|------|
| `source` | string | 是 | - | 源文件路径 |
| `target` | string | 是 | - | 目标文件路径 |
| `sheets` | list | 否 | None | 工作表列表（默认迁移所有） |

### 返回值

```python
{
    "success": True,
    "migrated_sheets": ["Sheet1", "Sheet2", "Sheet3"]
}
```

### 示例

```python
migrate(source="source.xlsx", target="target.xlsx")
migrate(source="source.xlsx", target="target.xlsx", sheets=["Sheet1", "Sheet2"])
```

---

## merge_files

合并多个 CSV/Excel 文件。

### 参数

| 参数 | 类型 | 必填 | 默认值 | 说明 |
|------|------|------|--------|------|
| `files` | list | 是 | - | 文件路径列表 |
| `output` | string | 是 | - | 输出文件路径 |
| `merge_type` | string | 否 | "concat" | 合并类型：`concat`/`merge`/`join` |
| `on` | string | 条件 | None | 合并键（merge/join 时必填） |
| `dedup` | bool | 否 | False | 是否去重 |
| `dedup_columns` | list | 否 | None | 去重列 |

### 返回值

```python
{
    "success": True,
    "rows_merged": 5000,
    "output": "merged.csv"
}
```

### 示例

```python
# 纵向合并
merge_files(files=["file1.csv", "file2.csv"], output="merged.csv", merge_type="concat")

# 按键合并
merge_files(files=["left.csv", "right.csv"], output="merged.csv", merge_type="merge", on="ID")

# 去重合并
merge_files(files=["f1.csv", "f2.csv"], output="merged.csv", dedup=True, dedup_columns=["ID"])
```

---

## join_tables

SQL 风格关联（LEFT/RIGHT/INNER/OUTER JOIN）。

### 参数

| 参数 | 类型 | 必填 | 默认值 | 说明 |
|------|------|------|--------|------|
| `left_file` | string | 是 | - | 左表 Excel 文件路径 |
| `left_sheet` | string | 是 | - | 左表工作表名称 |
| `right_file` | string | 是 | - | 右表 Excel 文件路径 |
| `right_sheet` | string | 是 | - | 右表工作表名称 |
| `on` | list | 是 | - | 关联键 |
| `how` | string | 否 | "inner" | 关联类型：`left`/`right`/`inner`/`outer` |
| `output` | string | 否 | None | 输出文件路径 |

### 返回值

```python
{
    "success": True,
    "left_rows": 1000,
    "right_rows": 500,
    "result_rows": 800,
    "result_columns": ["ID", "Name", "Sales", "Region"]
}
```

### 示例

```python
join_tables(
    left_file="customers.xlsx", left_sheet="Sheet1",
    right_file="orders.xlsx", right_sheet="Sheet1",
    on=["CustomerID"],
    how="left"
)
join_tables(
    left_file="left.xlsx", left_sheet="Sheet1",
    right_file="right.xlsx", right_sheet="Sheet1",
    on=["ID", "Name"],
    how="inner",
    output="joined.xlsx"
)
```

---

## batch_merge

多表批量合并（从文件夹批量合并多个 Excel 文件）。

### 参数

| 参数 | 类型 | 必填 | 默认值 | 说明 |
|------|------|------|--------|------|
| `folder` | string | 是 | - | 文件夹路径 |
| `pattern` | string | 否 | "*.xlsx" | 文件匹配模式 |
| `sheet` | string | 否 | None | 工作表名称 |
| `output` | string | 是 | - | 输出文件路径 |

### 返回值

```python
{
    "success": True,
    "file_count": 12,
    "success_count": 12,
    "total_rows": 50000,
    "columns": ["Date", "Product", "Sales"]
}
```

### 示例

```python
batch_merge(folder="/path/to/monthly_data", pattern="*.xlsx", output="annual.xlsx")
batch_merge(folder="/path/to/data", pattern="sales_*.xlsx", sheet="Sheet1", output="merged.xlsx")
```

---

## excel_to_markdown

将 Excel 表格转换为 Markdown 格式。

### 参数

| 参数 | 类型 | 必填 | 默认值 | 说明 |
|------|------|------|--------|------|
| `file` | string | 是 | - | Excel 文件路径 |
| `sheet` | string | 否 | None | 工作表名称（默认全部） |
| `output` | string | 否 | None | 输出文件路径 |
| `merge_mode` | string | 否 | "tl" | 合并单元格处理：`tl`（左上角值）/`fill`（填充） |
| `include_styles` | bool | 否 | True | 是否包含样式 |

### 返回值

```python
{
    "success": True,
    "markdown": "| Name | Sales |\n|------|-------|\n| A | 100 |",
    "output": "output.md"
}
```

### 示例

```python
excel_to_markdown(file="data.xlsx", sheet="Sheet1")
excel_to_markdown(file="data.xlsx", output="table.md", include_styles=False)
```
