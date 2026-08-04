---
name: abacus-transport
description: "均输章 - 导入导出与数据迁移。当需要导入CSV/JSON、导出数据、迁移工作表、合并文件时使用。"
version: 0.2.0
chapter: transport
level: rod
tags: [excel, import, export, csv, json, migrate, merge, markdown]
---

# 均输章 - 导入导出与数据迁移

## CRITICAL RULES
1. 导入前验证源文件格式和编码
2. 导出时指定明确的数据范围
3. 迁移前备份目标文件
4. 合并文件时检查列名一致性

## 能力一览

| 能力 | CLI 命令 | 说明 |
|------|----------|------|
| import_data | `abacus import-data` | 导入数据（CSV、JSON等） |
| export_data | `abacus export-data` | 导出数据（CSV、JSON） |
| migrate | `abacus migrate` | 数据迁移（工作表复制） |
| merge_files | `abacus merge-files` | 合并多个文件 |
| excel_to_markdown | `abacus excel-to-markdown` | 转换为 Markdown 格式 |

---

## 导入数据参考

### 支持的格式

| 格式 | 说明 | 参数 |
|------|------|------|
| CSV | 逗号分隔值 | source_type='csv' |
| JSON | JavaScript 对象表示法 | source_type='json' |

### 导入 CSV

```python
abacus_import_data(
    file='output.xlsx',
    source='data.csv',
    source_type='csv',
    sheet='ImportedData'
)
```

### 导入 JSON

```python
abacus_import_data(
    file='output.xlsx',
    source='data.json',
    source_type='json',
    sheet='ImportedData'
)
```

### 常见问题

| 问题 | 原因 | 解决方案 |
|------|------|----------|
| 中文乱码 | 编码错误 | 指定 UTF-8 编码 |
| 列数不匹配 | 源文件格式问题 | 检查 CSV 分隔符 |
| 数据类型错误 | 自动检测失败 | 手动指定列类型 |

---

## 导出数据参考

### 支持的格式

| 格式 | 说明 | 参数 |
|------|------|------|
| CSV | 逗号分隔值 | format='csv' |
| JSON | JavaScript 对象表示法 | format='json' |

### 导出 CSV

```python
abacus_export_data(
    file='data.xlsx',
    sheet='Sheet1',
    range='A1:D100',
    output='export.csv',
    format='csv'
)
```

### 导出 JSON

```python
abacus_export_data(
    file='data.xlsx',
    sheet='Sheet1',
    range='A1:D100',
    output='export.json',
    format='json'
)
```

### 导出最佳实践

1. **指定明确范围** - 避免导出空行
2. **验证数据类型** - 确保导出格式正确
3. **检查编码** - 中文数据使用 UTF-8

---

## 数据迁移参考

### migrate 能力说明

migrate 用于将数据从一个 Excel 文件迁移到另一个文件，支持：
- 工作表复制
- 选择性迁移
- 格式保留

### 使用场景

```python
# 迁移所有工作表
abacus_migrate(
    source='source.xlsx',
    target='target.xlsx'
)

# 迁移指定工作表
abacus_migrate(
    source='source.xlsx',
    target='target.xlsx',
    sheets=['Sheet1', 'Sheet2']
)
```

### 迁移最佳实践

1. **备份目标文件** - 迁移前创建备份
2. **验证工作表名** - 确保目标文件无重名工作表
3. **检查数据格式** - 迁移后验证格式保留

---

## 文件合并参考

### 合并类型

| 类型 | 说明 | 参数 |
|------|------|------|
| concat | 纵向合并（追加行） | merge_type='concat' |
| merge | 横向合并（连接列） | merge_type='merge' |
| join | 按键连接 | merge_type='join' |

### 纵向合并

```python
abacus_merge_files(
    files=['file1.csv', 'file2.csv', 'file3.csv'],
    output='merged.csv',
    merge_type='concat',
    dedup=True
)
```

### 横向合并

```python
abacus_merge_files(
    files=['file1.csv', 'file2.csv'],
    output='merged.csv',
    merge_type='merge'
)
```

### 按键连接

```python
abacus_merge_files(
    files=['file1.csv', 'file2.csv'],
    output='merged.csv',
    merge_type='join',
    on='ID'
)
```

### 合并最佳实践

1. **检查列名一致性** - 确保合并列名匹配
2. **处理重复数据** - 使用 dedup 参数去重
3. **验证数据完整性** - 合并后检查行数和内容

---

## Markdown 转换参考

### 能力说明

excel_to_markdown 将 Excel 表格转换为 Markdown 格式，适合：
- 技术文档嵌入
- README 文件
- 数据展示

### 使用场景

```python
# 转换整个工作簿
abacus_excel_to_markdown(
    file='data.xlsx',
    output='output.md'
)

# 转换指定工作表
abacus_excel_to_markdown(
    file='data.xlsx',
    sheet='Sheet1',
    output='output.md'
)
```

### 转换选项

| 选项 | 说明 | 参数 |
|------|------|------|
| include_styles | 包含样式信息 | include_styles=True |
| merge_mode | 合并单元格处理 | merge_mode='tl' (使用左上角值) |

---

## Power Query 概念

### 什么是 Power Query

Power Query 是 Excel 的数据连接和转换引擎，用于：
- 从多种数据源导入数据
- 清洗和转换数据
- 合并多个数据源
- 自动刷新数据

### M 代码语法基础

```m
let
    // 步骤定义
    Source = Excel.Workbook(File.Contents("data.xlsx")),
    Sheet1 = Source{[Name="Sheet1"]}[Data],
    // 更多步骤...
in
    // 最终结果
    Sheet1
```

### 常用函数

- `Excel.Workbook()`: 读取 Excel 文件
- `Table.SelectRows()`: 筛选行
- `Table.SelectColumns()`: 选择列
- `Table.TransformColumnTypes()`: 转换列类型
- `Table.Group()`: 分组
- `Table.Sort()`: 排序

### 工作流

1. 使用 `evaluate` 测试 M 代码
2. 验证结果正确
3. 创建永久查询
4. 设置刷新计划

### 限制

- openpyxl 不支持 Power Query 操作
- 需要 Excel COM API 或 Microsoft Graph API
- 某些高级功能需要 Power BI Desktop

---

## M 代码语法参考

### 列名引用

- 空格：`#"Column Name"`
- 特殊字符：`#"Column (With Parentheses)"`

### 命名范围

- 读取：`Excel.CurrentWorkbook(){[Name="MyRange"]}[Content]`

### 查询链接

- 引用其他查询：`Source = PreviousQuery`

### 常用函数

**数据源**
- `Excel.Workbook(File.Contents("file.xlsx"))` - 读取 Excel
- `Csv.Document(File.Contents("file.csv"))` - 读取 CSV

**表操作**
- `Table.SelectRows` - 筛选行
- `Table.SelectColumns` - 选择列
- `Table.RenameColumns` - 重命名列
- `Table.TransformColumnTypes` - 转换列类型

**聚合**
- `Table.Group` - 分组
- `Table.Aggregate` - 聚合

---

## 最佳实践

1. **验证源文件** - 导入前检查文件格式和编码
2. **指定明确范围** - 导出时避免导出空数据
3. **备份重要文件** - 迁移和合并前创建备份
4. **检查数据一致性** - 合并后验证数据完整性
5. **使用 UTF-8 编码** - 中文数据避免乱码

## 示例

```bash
# 导入 CSV
abacus import-data -f output.xlsx --source data.csv --source-type csv --sheet ImportedData

# 导出 CSV
abacus export-data -f data.xlsx -s Sheet1 -r A1:D100 --output export.csv --format csv

# 迁移工作表
abacus migrate --source source.xlsx --target target.xlsx --sheets Sheet1 Sheet2

# 合并文件
abacus merge-files --files file1.csv file2.csv --output merged.csv --merge-type concat

# 转换为 Markdown
abacus excel-to-markdown -f data.xlsx --output output.md
```