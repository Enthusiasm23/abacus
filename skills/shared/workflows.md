# 关键约束和工作流

## 会话生命周期

**始终遵循：打开 → 操作 → 保存关闭**

```python
wb = load_workbook(filepath)
try:
    # 操作...
    wb.save(filepath)
finally:
    wb.close()
```

## 计算模式优化

**批量写入时禁用自动计算：**

```python
# 1. 读取当前计算模式
# 2. 设置为手动模式
# 3. 执行所有写入
# 4. 手动触发计算
# 5. 恢复自动模式
```

**为什么重要：**
- 每次写入都会触发重新计算
- 100次写入 = 100次计算
- 手动模式：100次写入 + 1次计算

## Power Query 测试优先

**始终先测试后创建：**

```python
# 1. 使用 evaluate 测试 M 代码
result = evaluate_m_code(m_code)

# 2. 验证结果
if result_is_valid(result):
    # 3. 创建永久查询
    create_query(name, m_code)
```

**为什么重要：**
- 错误的 M 代码会损坏工作簿
- 测试可以提前发现问题
- 修复成本远低于事后修复

## 数据模型刷新策略

**工作表修改后必须显式刷新：**

```python
# 1. 修改工作表数据
modify_table_data()

# 2. 刷新数据模型
refresh_data_model()

# 3. 刷新透视表
refresh_pivot_tables()
```

**为什么重要：**
- 工作表表格和数据模型是独立副本
- 修改不会自动同步
- 必须显式刷新

## 图表定位策略

**多图表布局必须使用 targetRange：**

```python
# 好：明确指定位置
create_chart(range="A1:B10", position="D1")

# 差：自动定位（可能重叠）
create_chart(range="A1:B10")
```

**布局规则：**
- 每个图表都需要明确的位置
- 留出 1-2 行/列间距
- 检查重叠警告

## 错误处理模式

**错误驱动修正：**

```python
try:
    result = execute_operation()
except Error as e:
    # 1. 读取错误信息
    error_message = str(e)
    
    # 2. 分析错误原因
    if "not found" in error_message:
        # 3. 创建缺失的对象
        create_missing_object()
    
    # 4. 重试操作
    retry_operation()
```

**重试限制：**
- 最多重试 2 次
- 超过后报告错误
- 不要盲目重试
