# M 代码语法参考

## 基本语法

### 列名引用
- 空格：`#"Column Name"`
- 特殊字符：`#"Column (With Parentheses)"`

### 命名范围
- 读取：`Excel.CurrentWorkbook(){[Name="MyRange"]}[Content]`

### 查询链接
- 引用其他查询：`Source = PreviousQuery`

## 常用函数

### 数据源
- `Excel.Workbook(File.Contents("file.xlsx"))` - 读取 Excel
- `Csv.Document(File.Contents("file.csv"))` - 读取 CSV

### 表操作
- `Table.SelectRows` - 筛选行
- `Table.SelectColumns` - 选择列
- `Table.RenameColumns` - 重命名列
- `Table.TransformColumnTypes` - 转换列类型

### 聚合
- `Table.Group` - 分组
- `Table.Aggregate` - 聚合