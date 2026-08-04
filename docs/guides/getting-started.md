# 快速开始

## 安装

```bash
pip install abacus
```

## 基本使用

### 1. 读取 Excel 数据

```python
from abacus.core import MeasureRangeCapability

cap = MeasureRangeCapability()
result = cap.execute(None, file="data.xlsx", sheet="Sheet1", range="A1:D10")
print(result)
```

### 2. 创建公式

```python
from abacus.core import CreateFormulaCapability

cap = CreateFormulaCapability()
result = cap.execute(None, file="data.xlsx", sheet="Sheet1", cell="E1", formula="SUM(A1:D1)")
print(result)
```

### 3. 数据分析

```python
from abacus.core import AnalyzeStatsCapability

cap = AnalyzeStatsCapability()
result = cap.execute(None, file="data.xlsx", sheet="Sheet1", range="A1:D100")
print(result)
```

## CLI 使用

```bash
# 查看帮助
abacus --help

# 读取数据
abacus read -f data.xlsx -s Sheet1 -r A1:D10

# 查看结构
abacus structure -f data.xlsx

# 列出所有能力
abacus capabilities
```

## MCP Server 使用

```bash
# 启动 MCP Server
python -m abacus.mcp_server
```

## 下一步

- [查看使用示例](../../examples/)
- [查看 API 参考](../api/reference.md)
