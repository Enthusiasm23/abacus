# 数据模型概念指南

## 什么是数据模型

Excel 数据模型是：
- 内置的内存数据库
- 支持多表关系
- 支持 DAX 公式
- 支持大型数据集（百万行）

## DAX 基础

### 度量公式
```dax
// 求和
Total Sales = SUM(Sales[Amount])

// 计数
Order Count = COUNTROWS(Sales)

// 平均
Avg Sales = AVERAGE(Sales[Amount])

// 条件
High Sales = CALCULATE(SUM(Sales[Amount]), Sales[Amount] > 1000)
```

### 常用函数
- `SUM()`: 求和
- `AVERAGE()`: 平均值
- `COUNTROWS()`: 计数
- `CALCULATE()`: 条件计算
- `RELATED()`: 关联表数据
- `DIVIDE()`: 安全除法

## 关系管理

### 星型架构
```
事实表 (Sales)
├── 维度表 (Products)
├── 维度表 (Customers)
├── 维度表 (Dates)
└── 维度表 (Regions)
```

### 关系类型
- 一对多 (1:N): 最常见
- 多对一 (N:1): 反向关系
- 多对多 (N:N): 需要桥接表

## 工作流

### 创建数据模型
1. 创建 Excel 表格
2. 添加到数据模型
3. 定义关系
4. 创建 DAX 度量
5. 创建透视表

### 刷新策略
1. 修改源数据
2. 刷新数据模型
3. 刷新透视表
4. 刷新图表

## 限制

- openpyxl 不支持数据模型操作
- 需要 Excel COM API 或 OLE DB
- 某些功能需要 Power Pivot

## 替代方案

对于 Python 环境，可以使用：
- `pandas` 进行多表关联
- `sqlite3` 作为内存数据库
- 自定义聚合逻辑替代 DAX

---

## 服务器特定注意事项

### 前置条件：表必须先添加到数据模型

数据模型（Power Pivot）只包含明确添加的表。**不能**对不在数据模型中的表创建 DAX 度量。

### MSOLAP 前置条件（用于 evaluate/execute-dmv）

`evaluate` 和 `execute-dmv` 操作需要 Microsoft Analysis Services OLE DB Provider (MSOLAP)。

如果看到 "Class not registered" (0x80040154) 错误，请安装：
1. **Power BI Desktop**（推荐 - 包含 MSOLAP）
2. **Microsoft OLE DB Driver for Analysis Services**

### 关键：数据模型同步

**工作表表格和数据模型表是独立的副本！**

追加/修改工作表表格时，数据模型**不会**自动更新。必须显式刷新数据模型。

```
# 错误：数据仍显示旧值
table(append, tableName="Sales", csvData="...")  // 工作表已更新
datamodel(evaluate, daxQuery="...")               // 返回旧值！

# 正确：工作表更改后刷新数据模型
table(append, tableName="Sales", csvData="...")  // 工作表已更新
datamodel(refresh)                                 // 同步到数据模型
datamodel(evaluate, daxQuery="...")               // 返回新值！
```

**自动刷新时：**
- `powerquery(refresh)` 同时刷新 Power Query 和数据模型

**需要手动刷新时：**
- `table(append)` 后
- `range(set-values)` 修改表格数据后
- 任何手动/直接工作表编辑后

### Excel Power Pivot 限制（对比 SSAS/Power BI）

| 功能 | Power BI/SSAS | Excel Power Pivot | 解决方案 |
|------|---------------|-------------------|----------|
| 计算表 | DAX: `MyTable = FILTER(...)` | 不支持 | 使用 Power Query 创建表 |
| 计算列 | DAX: `Table[Col] = ...` | 无 COM API 访问 | 使用 Power Query 或 DAX 度量 |
| 度量 | 完全支持 | 完全支持 | - |
| 关系 | 完全支持 | 完全支持 | - |

### 如何将表添加到数据模型

| 来源 | 方法 |
|------|------|
| 工作表 Excel 表格 | table 的 add-to-data-model 操作 |
| 外部文件（CSV 等） | powerquery 的 loadDestination='data-model' |
| 数据库/Web 源 | powerquery 的 loadDestination='data-model' |

### DAX 度量创建

- tableName：度量所属的表（用于组织）
- measureName：度量显示名称
- daxFormula：DAX 表达式（例如 "SUM(Sales[Revenue])"）
- formatString：可选数字格式（#,##0.00, 0%, $#,##0 等）

### 常见错误

- 在将源表添加到数据模型之前创建度量 → 错误
- 使用工作表表格名称而非数据模型表名称
- 忘记 delete-table 会删除该表的所有度量
- 创建度量时未指定 tableName（必需）

### 显示数据模型数据 - 选择正确的输出

| 目标 | 最佳工具 | 原因 |
|------|----------|------|
| **平面查询结果** | `table create-from-dax` | 清晰表格显示，无透视表 UI |
| **静态报告/快照** | `table create-from-dax` | DAX 做聚合，表格只显示 |
| **公式数据** | `table create-from-dax` | 使用结构化引用 |
| **交互式钻取** | `pivottable` | 用户可重新分组、筛选、展开/折叠 |
| **交叉表（行 × 列）** | `pivottable` | 带行/列字段的矩阵布局 |

### 星型架构

**为什么用 DAX 而非 Power Query 做计算？**
- DAX 在刷新时重新计算，无需重新运行 Power Query
- 当查找/利率表经常变化时有用