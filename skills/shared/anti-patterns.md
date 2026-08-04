# 反模式指南

## 1. 硬编码计算值反模式
**Bad:** 在 Python 中计算后写入静态值
```python
result = sum(values)  # Python 计算
ws["A1"] = result     # 写入静态值
```

**Good:** 使用 Excel 公式
```python
ws["A1"] = "=SUM(A2:A100)"  # 写入公式
```

## 2. 不检查工作表存在性反模式
**Bad:** 直接访问工作表
```python
ws = wb["Sheet1"]  # 可能不存在
```

**Good:** 先检查再访问
```python
if "Sheet1" in wb.sheetnames:
    ws = wb["Sheet1"]
else:
    raise SheetNotFoundError("Sheet1 not found")
```

## 3. 不关闭工作簿反模式
**Bad:** 不关闭工作簿
```python
wb = load_workbook("file.xlsx")
# 操作...
# 忘记关闭
```

**Good:** 使用 with 或 try/finally
```python
wb = load_workbook("file.xlsx")
try:
    # 操作...
finally:
    wb.close()
```

## 4. 忽略数据验证反模式
**Bad:** 不验证输入参数
```python
def execute(self, context, **params):
    file = params.get("file")  # 可能为 None
```

**Good:** 验证输入参数
```python
def execute(self, context, **params):
    file = params.get("file")
    if not file:
        raise DataError("file parameter is required")
```

## 5. 冗余格式化反模式

### 问题

对同一范围重复应用相同格式：

```
错误：重复应用粗体

range_format(action: 'format-range', rangeAddress: 'A1:D1', bold: true)
range_format(action: 'format-range', rangeAddress: 'A1:D1', bold: true, fillColor: '#4472C4')
// 粗体已应用，第二次调用不必要地重复应用
```

错误：为每个属性单独调用：

```
range_format(action: 'format-range', rangeAddress: 'A1:D1', bold: true)
range_format(action: 'format-range', rangeAddress: 'A1:D1', fillColor: '#4472C4')
range_format(action: 'format-range', rangeAddress: 'A1:D1', fontColor: '#FFFFFF')
range_format(action: 'format-range', rangeAddress: 'A1:D1', horizontalAlignment: 'center')
```

### 解决方案

**一次**调用设置所有格式属性：

```
正确：每次范围一次调用

range_format(action: 'format-range', rangeAddress: 'A1:D1',
    bold: true, fillColor: '#4472C4', fontColor: '#FFFFFF', horizontalAlignment: 'center')
```

对多个不连续范围使用共享格式：

```
正确：一次共享多范围格式化调用

range_format(action: 'format-ranges',
    rangeAddresses: ['A1:D1', 'A12:D12', 'A24:D24'],
    bold: true, fillColor: '#4472C4', fontColor: '#FFFFFF')
```

## 6. 错误样式系统反模式

### 问题

对有自己样式系统的对象使用 `range_format`：

```
错误：格式化表格标题行

table(action: 'create', tableName: 'Sales', rangeAddress: 'A1:D10')
range_format(action: 'format-range', rangeAddress: 'A1:D1', bold: true, fillColor: '#4472C4')
// 表格样式已控制标题外观 - 这会产生不一致的覆盖
```

### 解决方案

使用每种对象类型所属的样式系统：

| 对象 | 正确的样式方法 | 不要使用 |
|------|----------------|----------|
| Excel 表格 | `table(action:'set-style')` 或创建时的 `tableStyle` | `range_format` |
| 透视表 | 不支持 - 保留默认 | `range_format`（刷新时清除） |
| 图表 | `chart_config(action:'set-style', styleNumber: 1-48)` | `range_format` |
| 普通单元格/范围 | `range_format` | — |

## 7. 删除并重建反模式

### 问题

为做小改动而删除整个结构：

```
错误：用户想更新单元格 B5

table(action: 'delete', tableName: 'SalesData')
range(action: 'set-values', values: [[整个数据集，B5 已修复]])
table(action: 'create', tableName: 'SalesData', ...)
```

这会破坏：
- 单元格格式
- 条件格式规则
- 数据验证
- 指向表格的命名范围
- 透视表连接
- 引用表格的 DAX 度量

### 解决方案

使用有针对性的修改：

```
正确：仅更新更改的单元格

range(action: 'set-values', rangeAddress: 'B5', values: [[newValue]])
```

## 8. 发现循环反模式

### 问题

重复执行 `file(list)`、`worksheet(list)` 或 `table(list)` 而不采取行动：

```
错误：错误后循环发现

worksheet(action: 'list')           → 获取工作表列表
worksheet(action: 'list')           → 再次获取相同列表
file(action: 'list')                → 获取会话列表
worksheet(action: 'list')           → 再次获取相同列表
...（数十次重复）
```

### 解决方案

如果已有 sessionId，直接使用：

```
正确：使用已有的 sessionId

错误："session expired"
→ file(action: 'open', path: original_path)  ← 重新打开一次
→ 立即使用新 sessionId 继续
```

**规则：**
- 最多重试 **2 次**
- 2 次失败后：停止重试，报告错误
- **永远不要连续调用 `list` 超过两次**

## 9. 确认循环反模式

### 问题

每个操作都请求确认：

```
错误：

用户："创建销售报告"
AI："要我为销售报告创建新的 Excel 文件吗？"
用户："是"
AI："您希望文件命名为什么？"
用户："sales_report.xlsx"
AI："要放在文档文件夹吗？"
...（10 个更多问题）
```

### 解决方案

使用合理默认值执行，报告结果：

```
正确：

用户："创建销售报告"
AI："已创建销售报告 C:\Users\You\Documents\sales_report.xlsx：
- 工作表 'Summary' 包含标题：Date, Product, Region, Sales
- 准备输入数据

您想添加什么数据？"
```

## 10. 单元格更新反模式

### 问题

读取整个范围，在内存中修改，写回整个范围：

```
错误：通过重写数千个单元格来更新一个单元格

data = range(action: 'get-values', rangeAddress: 'A1:Z1000')
data[4][1] = "new value"
range(action: 'set-values', rangeAddress: 'A1', values: data)
```

这会：
- 不必要地传输兆字节数据
- 中断时有数据损坏风险
- 破坏公式（仅值，非公式）
- 丢失单元格格式

### 解决方案

仅写入更改的单元格：

```
正确：直接单元格更新

range(action: 'set-values', rangeAddress: 'B5', values: [["new value"]])
```

## 11. 会话泄漏反模式

### 问题

打开文件而不关闭：

```
错误：会话累积

file(action: 'open', filePath: 'file1.xlsx')  // 会话 1
file(action: 'open', filePath: 'file2.xlsx')  // 会话 2
file(action: 'open', filePath: 'file3.xlsx')  // 会话 3
// ... 从不关闭
```

### 解决方案

始终关闭会话：

```
正确：正确的生命周期

session1 = file(action: 'open', path: 'file1.xlsx')
// ... 处理 file1 ...
file(action: 'close', sessionId: session1, save: true)
```

## 12. 忽略错误上下文反模式

### 问题

不读取错误就重试失败操作：

```
错误：盲目重试

datamodel(action: 'create-measure', ...) → 错误：表不在数据模型中
datamodel(action: 'create-measure', ...) → 错误：表不在数据模型中
```

### 解决方案

读取并根据错误上下文行动：

```
正确：基于错误的修正

datamodel(action: 'create-measure', ...) 
→ 错误：表 'Sales' 不在数据模型中
→ 建议：table(action: 'add-to-data-model', tableName: 'Sales')

table(action: 'add-to-data-model', tableName: 'Sales')  // 修复前置条件
datamodel(action: 'create-measure', ...)  // 现在成功
```

## 13. 数字格式区域反模式

### 问题

使用区域特定的格式代码：

```
错误：德语/欧洲格式

range(action: 'set-number-format', formatCode: '#.##0,00')  // 德语
range(action: 'set-number-format', formatCode: '# ##0,00')  // 法语
```

### 解决方案

始终使用美国格式代码（Excel 自动翻译）：

```
正确：美国格式代码（通用）

range(action: 'set-number-format', formatCode: '#,##0.00')
```