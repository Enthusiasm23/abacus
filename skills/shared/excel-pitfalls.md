# Excel 常见坑（15 个）

## 1. 公式变字符串
**错误：** 写入公式后变成文本
```python
ws["A1"] = "SUM(B1:B10)"  # 错误：缺少 =
ws["A1"] = "=SUM(B1:B10)" # 正确
```

## 2. data_only=True 丢公式
**错误：** 读取后保存丢失公式
```python
wb = load_workbook("file.xlsx", data_only=True)
wb.save("file.xlsx")  # 公式永久丢失！
```

## 3. pandas 不能写公式
**错误：** 用 pandas 写入公式
```python
df.to_excel("file.xlsx")  # 公式变成字符串
```

## 4. sheet 名禁止字符
**错误：** sheet 名包含 : \ / ? * [ ]
```python
ws.title = "Sheet:1"  # 错误！
ws.title = "Sheet1"   # 正确
```

## 5. 日期显示为数字
**错误：** 日期显示为 44197 而非 2021-01-01
```python
cell.number_format = "YYYY-MM-DD"  # 设置日期格式
```

## 6. 百分比字符串化
**错误：** 写入 "30%" 变成文本
```python
cell.value = 0.3
cell.number_format = "0.0%"  # 正确
```

## 7. 列宽不适配
**错误：** 列宽太窄显示 ######
```python
ws.column_dimensions["A"].width = 15  # 手动设置
```

## 8. PatternFill 缺 fill_type
**错误：** 颜色不显示
```python
cell.fill = PatternFill(fgColor="FF0000")  # 错误
cell.fill = PatternFill(pattern_type="solid", fgColor="FF0000")  # 正确
```

## 9. 跨 sheet 引用缺引号
**错误：** sheet 名有空格时引用失败
```python
ws["A1"] = "=Sheet Name!B1"   # 错误
ws["A1"] = "='Sheet Name'!B1" # 正确
```

## 10. 中文字体回退
**错误：** 中文显示为方块
```python
cell.font = Font(name="Microsoft YaHei")  # 设置中文字体
```

## 11. 整列引用拖慢性能
**错误：** SUM(A:A) 引用百万行
```python
ws["A1"] = "=SUM(A1:A1000)"  # 限定范围
```

## 12. 合并单元格排序失败
**错误：** 合并单元格导致排序异常
```python
ws.unmerge_cells("A1:B1")  # 先取消合并
```

## 13. 隐藏行被复制
**错误：** 复制时包含隐藏行
```python
# 需要手动检查行可见性
```

## 14. 条件格式覆盖
**错误：** 多个条件格式冲突
```python
# 需要设置优先级
```

## 15. 数据验证绕过
**错误：** 程序写入绕过验证
```python
# 需要在写入前验证
```
