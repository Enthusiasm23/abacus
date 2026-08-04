# 常见坑与最佳实践

## openpyxl 8 大常见坑

### 1. data_only=True 后 save() 导致公式永久丢失
```python
# 错误：data_only=True 读取后保存会丢失公式
wb = load_workbook("file.xlsx", data_only=True)
wb.save("file.xlsx")  # 公式永久丢失！

# 正确：需要保留公式时不要用 data_only=True
wb = load_workbook("file.xlsx", data_only=False)
```

### 2. 公式只是字符串，不自动计算
```python
# openpyxl 只写入公式字符串，不计算结果
ws["A1"] = "=SUM(B1:B10)"  # 写入的是字符串

# 需要用 LibreOffice 或 Excel 打开才会计算
```

### 3. PatternFill 必须传 fill_type
```python
# 错误：颜色不显示
cell.fill = PatternFill(fgColor="FF0000")

# 正确：必须指定 pattern_type
cell.fill = PatternFill(pattern_type="solid", fgColor="FF0000")
```

### 4. 行列号是 1-based
```python
# openpyxl 行列号从 1 开始，不是 0
cell = ws.cell(row=1, column=1)  # A1
```

### 5. Sheet 名禁止字符
```python
# 不能包含 : \ / ? * [ ]
# 长度不超过 31 字符
ws.title = "Sheet:1"  # 错误！
ws.title = "Sheet1"   # 正确
```

### 6. 跨 sheet 引用需要单引号
```python
# 如果 sheet 名包含空格或特殊字符
ws["A1"] = "='Sheet Name'!B1"
```

### 7. 中文字符字体回退
```python
# 中文字符需要指定中文字体
cell.font = Font(name="Microsoft YaHei")  # 微软雅黑
```

### 8. 图表 Reference 忘了 min_col
```python
# 错误：图表数据引用错误
ref = Reference(ws, min_row=1, max_row=10, min_col=1)

# 正确：需要指定 max_col
ref = Reference(ws, min_row=1, max_row=10, min_col=1, max_col=2)
```

## 格式化最佳实践

### 使用 NamedStyle 复用样式
```python
from openpyxl.styles import NamedStyle

# 创建样式
header_style = NamedStyle(name="header")
header_style.font = Font(bold=True, size=12)
header_style.fill = PatternFill(pattern_type="solid", fgColor="4472C4")
header_style.alignment = Alignment(horizontal="center")

# 应用样式
cell.style = header_style
```

### 自动列宽计算
```python
from openpyxl.utils import get_column_letter

for col_idx, col in enumerate(ws.columns, 1):
    max_len = max((len(str(c.value)) if c.value else 0) for c in col)
    ws.column_dimensions[get_column_letter(col_idx)].width = max_len + 2
```

### 颜色处理
```python
# 确保颜色有 FF 前缀（不透明度）
color = "FF0000"
if not color.startswith('FF'):
    color = f'FF{color}'
font_color = Color(rgb=color)
```

## 性能优化

### 批量写入
```python
# 错误：逐单元格写入
for row in range(1, 1000):
    for col in range(1, 10):
        ws.cell(row=row, column=col, value=f"data_{row}_{col}")

# 正确：使用 append 批量写入
for row in range(1, 1000):
    ws.append([f"data_{row}_{col}" for col in range(1, 10)])
```

### 使用 read_only 模式
```python
# 只读模式更快
wb = load_workbook("file.xlsx", read_only=True)
```

### 使用 write_only 模式
```python
# 写入大量数据时使用 write_only 模式
wb = Workbook(write_only=True)
ws = wb.create_sheet()
for row in data:
    ws.append(row)
wb.save("file.xlsx")
```

## 更多常见坑

### 9. 整列引用拖慢性能
**错误：** SUM(A:A) 引用百万行
**正确：** SUM(A1:A1000) 限定范围

### 10. 合并单元格排序失败
**错误：** 合并单元格导致排序异常
**正确：** 先取消合并再排序

### 11. 隐藏行被复制
**错误：** 复制时包含隐藏行
**正确：** 检查行可见性

### 12. 条件格式覆盖
**错误：** 多个条件格式冲突
**正确：** 设置优先级
