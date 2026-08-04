# Abacus 项目规约

> 本文件定义项目规约，AI Agent 和开发者应优先参考此文件。

---

## 一、项目定位

**Abacus（算盘）** 是本地 Excel 自动化框架，以中国算学文化为命名体系。

- **技术栈**：Python 3.10+ / openpyxl / click / fastmcp
- **入口**：CLI + MCP Server + Python 库
- **命名体系**：九章分类（方田、粟米、衰分、少广、商功、均输、盈不足、方程、勾股）
- **版本**：1.0.0（从 1.0.0 开始版本控制）

---

## 二、九章分类体系

| 章节 | 英文 | 能力数 | 核心功能 |
|------|------|--------|----------|
| **方田章** | Field | 10 | 数据读取、结构查看、单元格操作、工作表管理 |
| **粟米章** | Grain | 10 | 格式转换、类型转换、数据清洗、转置、分列、模糊匹配、管道 |
| **衰分章** | Share | 4 | 分组、分配、汇总、分类汇总 |
| **少广章** | Dimension | 5 | 数学计算、反向推导、解方程、自动求和 |
| **商功章** | Work | 25 | 批量操作、图表管理、格式化、筛选、保护、报告、模板 |
| **均输章** | Transport | 7 | 导入导出、迁移、合并、Markdown转换、工作表拆分 |
| **盈不足章** | Balance | 8 | 数据验证、文件验证、质量检查、代码审计、文件审计 |
| **方程章** | Equation | 5 | 公式创建、公式诊断、公式重算、公式生成、数组公式 |
| **勾股章** | Triangle | 12 | 统计分析、趋势分析、相关性分析、数据可视化、高级分析、差异分析、金融分析 |

**总计：103 个能力**

---

## 三、版本控制

### 3.1 版本号规范

本项目从 **1.0.0** 开始版本控制，遵循 [Semantic Versioning](https://semver.org/) 规范：

- **主版本号 (Major)**：不兼容的 API 变更
- **次版本号 (Minor)**：向后兼容的功能新增
- **修订号 (Patch)**：向后兼容的问题修复

### 3.2 版本记录

| 版本 | 日期 | 说明 |
|------|------|------|
| 1.0.0 | 2026-08-09 | 首次发布，103 个能力，1136 个测试 |

### 3.3 发布流程

1. 更新版本号（pyproject.toml, __init__.py, cli.py）
2. 更新 CHANGELOG.md
3. 运行完整测试
4. 提交代码
5. 打 tag：`git tag v1.0.0`
6. 推送：`git push origin main --tags`

---

## 四、AI 开发加速工具

### 4.1 SkillIndexer（知识库索引器）

```bash
# CLI 命令
abacus skill-search "公式验证"
abacus skill-graph abacus-field
abacus skill-stats
abacus skill-index-build

# MCP 工具
abacus_skill_search(query="formula")
abacus_skill_graph(skill_name="abacus-field")
abacus_skill_stats()
abacus_skill_index_build()
```

- **位置**：`src/abacus/skill/indexer.py`
- **用途**：全文搜索 SKILL.md 和知识文件、查看关联图谱

---

## 五、核心设计原则

### 5.1 能力设计原则

1. **单一职责**：每个能力只做一件事
2. **九章分类**：所有能力必须归属到九章之一
3. **实际实现**：不允许空壳实现，必须有实际功能
4. **测试覆盖**：每个能力必须有测试用例
5. **文档完整**：每个能力必须有 SKILL.md 文档

### 5.2 代码风格

- 遵循 PEP 8 规范
- Python 3.10+ 新语法（`list[Type]` 而非 `List[Type]`）
- 使用类型提示
- 编写清晰的函数文档
- 使用结构化异常（DataError, FileNotFoundError 等）

### 5.3 测试策略

- **TDD**：测试驱动开发
- **覆盖率**：目标 90%
- **测试类型**：单元测试、集成测试、MCP 工具测试、CLI 命令测试

---

## 六、目录结构

```
abacus/
├── src/abacus/
│   ├── __init__.py         # 包入口，导出所有能力
│   ├── cli.py              # CLI 入口
│   ├── mcp_server.py       # MCP Server
│   ├── i18n.py             # 国际化支持
│   ├── logging.py          # 日志配置
│   ├── core/               # 九章能力
│   │   ├── field/          # 方田章（10 能力）
│   │   ├── grain/          # 粟米章（10 能力）
│   │   ├── share/          # 衰分章（4 能力）
│   │   ├── dimension/      # 少广章（5 能力）
│   │   ├── work/           # 商功章（25 能力）
│   │   ├── transport/      # 均输章（7 能力）
│   │   ├── balance/        # 盈不足章（8 能力）
│   │   ├── equation/       # 方程章（5 能力）
│   │   └── triangle/       # 勾股章（12 能力）
│   ├── adapters/           # 文件适配器
│   ├── skill/              # 知识库索引器
│   └── lock/               # 文件锁
├── tests/                  # 测试（1136 个）
├── skills/                 # SKILL.md 文档
├── docs/                   # API 文档
└── examples/               # 使用示例
```

---

## 七、关键约束

### 7.1 发布约束

- **建仓和发布仅在用户主动提起时执行**
- 不作为主动事项提出

### 7.2 代码约束

- 所有能力必须有实际实现（非空壳）
- 所有能力必须有测试用例
- 所有 MCP 工具必须有详细文档
- 所有 CLI 命令必须有帮助信息

### 7.3 文档约束

- 知识库文件放入 skills/shared/ 目录
- SKILL.md 文件放入 skills/<chapter>/ 目录
- 设计文档放入 docs/ 目录
- 使用示例放入 examples/ 目录

---

**最后更新**：2026-08-09
