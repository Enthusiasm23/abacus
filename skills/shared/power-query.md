# Power Query 概念指南

## 什么是 Power Query

Power Query 是 Excel 的数据连接和转换引擎，用于：
- 从多种数据源导入数据
- 清洗和转换数据
- 合并多个数据源
- 自动刷新数据

## M 代码语法基础

### 基本结构
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

## 工作流

### 测试优先模式
1. 使用 `evaluate` 测试 M 代码
2. 验证结果正确
3. 创建永久查询
4. 设置刷新计划

### 常见模式
```m
// 读取 Excel 文件
let
    Source = Excel.Workbook(File.Contents("data.xlsx")),
    Sheet1 = Source{[Name="Sheet1"]}[Data],
    PromotedHeaders = Table.PromoteHeaders(Sheet1),
    FilteredRows = Table.SelectRows(PromotedHeaders, each [Status] = "Active")
in
    FilteredRows
```

## 限制

- openpyxl 不支持 Power Query 操作
- 需要 Excel COM API 或 Microsoft Graph API
- 某些高级功能需要 Power BI Desktop

## 替代方案

对于 Python 环境，可以使用：
- `pandas` 进行数据转换
- `openpyxl` 读取/写入 Excel
- 自定义 ETL 脚本替代 Power Query
```