# 🧮 Abacus（算盘）

> 🎯 本地 Excel 自动化框架，以中国算学文化为命名体系

---

## 📊 项目状态

| 指标 | 数值 |
|------|------|
| 🎯 核心能力 | **103** 个 |
| 🧪 测试用例 | **1136** 个（全部通过） |
| 🔧 MCP 工具 | **107** 个 |
| 📝 CLI 命令 | **100+** 个 |
| 📚 知识库文件 | **38** 个 |
| 📈 测试覆盖率 | **88%** |

---

## ✨ 特性

- 🏛️ **九章分类体系** - 方田、粟米、衰分、少广、商功、均输、盈不足、方程、勾股
- 🔧 **三入口** - CLI + MCP Server + Python 库
- 🤖 **AI 友好** - 完整的 SKILL.md 文档
- 🌐 **中英文支持** - i18n 国际化
- 📝 **结构化日志** - 可配置日志级别
- 🔒 **安全可靠** - 路径遍历检查、ReDoS 防护、公式注入防护

---

## 🚀 快速开始

### 安装

```bash
pip install abacus-excel
```

### CLI 使用

```bash
# 查看帮助
abacus --help

# 读取数据
abacus read -f data.xlsx -s Sheet1 -r A1:D10

# 查看结构
abacus structure -f data.xlsx

# 转换格式
abacus convert-format -f data.xlsx -s Sheet1 -r A1:A100 -t number
```

### Python 库使用

```python
from abacus import MeasureRangeCapability, CapabilityRegistry

# 读取数据
cap = MeasureRangeCapability()
result = cap.execute(None, file="data.xlsx", sheet="Sheet1", range="A1:D10")
print(result)
```

### MCP Server 使用

```bash
# 启动 MCP Server
python -m abacus.mcp_server
```

---

## 🏛️ 九章分类

| 章节 | 能力数 | 说明 |
|------|--------|------|
| 📖 方田章 | 10 | 数据读取、单元格操作、工作表管理 |
| 🔄 粟米章 | 10 | 格式转换、数据清洗、转置、分列 |
| 📊 衰分章 | 4 | 分组、分配、汇总、分类汇总 |
| 🔢 少广章 | 5 | 数学计算、反向推导、解方程 |
| ⚙️ 商功章 | 25 | 批量操作、图表管理、格式化、筛选 |
| 📤 均输章 | 7 | 导入导出、迁移、合并 |
| ✅ 盈不足章 | 8 | 数据验证、文件验证、质量检查 |
| 📐 方程章 | 5 | 公式创建、诊断、重算、生成 |
| 📈 勾股章 | 12 | 统计分析、趋势分析、数据可视化 |

---

## 📁 项目结构

```
abacus/
├── src/abacus/           # 核心代码
│   ├── core/             # 九章能力
│   ├── cli.py            # CLI 入口
│   └── mcp_server.py     # MCP Server
├── tests/                # 测试（1136 个）
├── skills/               # SKILL.md 文档
├── docs/                 # API 文档
└── examples/             # 使用示例
```

---

## 📚 文档

- [🚀 快速开始](docs/guides/getting-started.md)
- [🏛️ 九章使用指南](docs/guides/nine-chapters-guide.md)
- [📖 API 参考](docs/api/reference.md)
- [📋 项目规约](AGENTS.md)
- [📝 变更日志](CHANGELOG.md)

---

## 📄 许可证

[MIT License](LICENSE)
