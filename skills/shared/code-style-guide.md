# 代码风格指南

## Python 代码规范

### 原则
- 简洁、无冗余注释
- 避免冗长变量名
- 避免不必要的 print 语句

### 示例

```python
# ❌ 错误：冗余注释
# 打开工作簿
wb = load_workbook("file.xlsx")

# ✅ 正确：简洁代码
wb = load_workbook("file.xlsx")
```

```python
# ❌ 错误：冗余变量名
the_workbook_that_we_are_loading = load_workbook("file.xlsx")

# ✅ 正确：简洁变量名
wb = load_workbook("file.xlsx")
```

## Excel 文件规范

### 原则
- 复杂公式添加注释
- 重要假设注明来源
- 关键计算添加说明

### 示例

```python
# 添加公式注释
ws["B10"] = "=SUM(B2:B9)"
ws["B10"].comment = Comment("汇总所有月份销售额", "System")

# 来源文档
ws["B6"] = 0.05
ws["B6"].comment = Comment("假设增长率 5%", "财务部, 2024-01-15, 预算文件")
```

## 错误消息规范

### 原则
- 结构化（code + message + details）
- 可操作（告诉用户怎么修复）
- 机器可解析（JSON 格式）

### 示例

```json
{
  "code": "FORMULA_ERROR",
  "message": "公式包含错误",
  "details": {
    "cell": "B10",
    "error": "#REF!",
    "suggestion": "检查引用的单元格是否存在"
  }
}
```
