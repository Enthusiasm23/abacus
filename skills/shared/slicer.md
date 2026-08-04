# 切片器概念指南

## 什么是切片器

切片器是 Excel 的可视化筛选器，用于：
- 交互式筛选数据
- 多个透视表/图表联动
- 直观的用户界面

## 切片器类型

### 透视表切片器
- 连接到透视表字段
- 支持多选
- 支持清除筛选

### 表格切片器
- 连接到 Excel 表格列
- 支持多选
- 支持搜索

## 工作流

### 创建切片器
1. 选择数据源（透视表/表格）
2. 选择字段
3. 设置位置和大小
4. 格式化外观

### 连接多个透视表
1. 创建主透视表
2. 创建切片器
3. 将切片器连接到其他透视表

## 样式设置

- 颜色方案
- 按钮大小
- 列数
- 高度/宽度

## 限制

- openpyxl 不支持切片器操作
- 需要 Excel COM API
- 某些高级功能需要 Power BI

## 替代方案

对于 Python 环境，可以使用：
- 下拉列表替代切片器
- 数据验证实现筛选
- 自定义 Web 界面

---

## 服务器特定注意事项

### 切片器类型

两种不同的切片器类型：
- **透视表切片器**：筛选透视表（可控制多个透视表）
- **表格切片器**：筛选 Excel 表格（仅单个表格）

### 操作

| 操作 | 说明 | 必需参数 |
|------|------|----------|
| `create-slicer` | 创建透视表切片器 | pivotTableName, fieldName |
| `list-slicers` | 列出所有透视表切片器 | （无） |
| `set-slicer-selection` | 设置透视表切片器筛选 | slicerName, selectedItems |
| `delete-slicer` | 删除透视表切片器 | slicerName |
| `create-table-slicer` | 创建表格切片器 | tableName, columnName |
| `list-table-slicers` | 列出所有表格切片器 | （无） |
| `set-table-slicer-selection` | 设置表格切片器筛选 | slicerName, selectedItems |
| `delete-table-slicer` | 删除表格切片器 | slicerName |

### 命名约定

- 如果未提供 `slicerName`，自动生成 `{FieldName}Slicer` 或 `{ColumnName}Slicer`
- 切片器名称在工作簿中必须唯一
- 使用 `list-slicers` 或 `list-table-slicers` 检查现有名称

### 选择行为

- `selectedItems` 是字符串列表：`["Value1", "Value2"]`
- 空列表 `[]` 清除所有筛选（显示所有项）
- 值必须精确匹配（区分大小写）
- 无效值被静默忽略

### 定位

- `destinationSheet` 指定哪个工作表承载切片器
- `position` 是左上角的单元格地址（例如 `'E1'`、`'G5'`）
- 未指定位置时：Excel 自动选择

### 常见错误

- 为不在透视表中的字段创建切片器 → 错误
- 为不在表格中的列创建表格切片器 → 错误
- 使用错误大小写设置选择 → 值被忽略（筛选显示为空）
- 删除不存在的切片器 → 错误

### 最佳实践

1. 创建前调用 `list-slicers` 避免名称冲突
2. 使用 `list-slicers` 获取精确的切片器名称
3. 多透视表筛选：创建一个切片器，在 Excel UI 中连接到多个透视表