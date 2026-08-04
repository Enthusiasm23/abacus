# Abacus API 参考文档

> 更新时间：2026-08-09  
> 版本：0.3.0

---

## 概述

Abacus（算盘）是一个本地 Excel 自动化框架，以中国算学文化为命名体系。通过 MCP Server 暴露 **86+ 个工具**，覆盖 Excel 数据读取、格式转换、分组汇总、数学计算、批量操作、导入导出、数据验证、公式管理和数据分析等场景。

## 九章分类体系

| 章节 | 英文 | 工具数 | 核心功能 | 详见 |
|------|------|--------|----------|------|
| **方田章** | Field | 12 | 数据读取、单元格操作、工作表管理 | [field.md](field.md) |
| **粟米章** | Grain | 10 | 格式转换、类型转换、数据清洗、转置 | [grain.md](grain.md) |
| **衰分章** | Share | 5 | 分组、分配、汇总、分类汇总 | [share.md](share.md) |
| **少广章** | Dimension | 5 | 数学计算、反向推导、解方程、自动求和 | [dimension.md](dimension.md) |
| **商功章** | Work | 33 | 批量操作、图表管理、格式化、筛选、保护、报告 | [work.md](work.md) |
| **均输章** | Transport | 7 | 导入导出、迁移、合并、Markdown转换 | [transport.md](transport.md) |
| **盈不足章** | Balance | 7 | 数据验证、文件验证、质量检查、代码审计 | [balance.md](balance.md) |
| **方程章** | Equation | 5 | 公式创建、诊断、重算、生成、数组公式 | [equation.md](equation.md) |
| **勾股章** | Triangle | 6 | 统计分析、趋势分析、相关性分析、可视化 | [triangle.md](triangle.md) |

**总计：90 个工具**

---

## 快速开始

### 1. 启动 MCP Server

```bash
python -m abacus.mcp_server
```

### 2. 配置 OpenCode

在 `opencode.json` 中添加 MCP Server：

```json
{
  "mcpServers": {
    "abacus": {
      "command": "python",
      "args": ["-m", "abacus.mcp_server"]
    }
  }
}
```

### 3. 调用示例

```python
# 读取数据
measure_range(file="data.xlsx", sheet="Sheet1", range="A1:D10")

# 转换格式
convert_format(file="data.xlsx", sheet="Sheet1", range="B2:B100", format_type="number")

# 创建图表
create_chart(file="data.xlsx", sheet="Sales", range="A1:C10", chart_type="bar", title="销售趋势")
```

---

## 通用参数说明

大多数工具共享以下参数：

| 参数 | 类型 | 必填 | 说明 |
|------|------|------|------|
| `file` | string | 是 | Excel 文件路径 |
| `sheet` | string | 是 | 工作表名称 |
| `range` | string | 是 | 数据范围，使用 A1 表示法（如 `A1:D10`） |

### 返回值格式

所有工具返回 `dict` 类型，通常包含：

| 字段 | 说明 |
|------|------|
| `success` | 操作是否成功 |
| `error` | 错误信息（失败时） |
| 其他字段 | 工具特定的返回数据 |

---

## 工具总览

### 方田章 (Field) - 12 个工具

| 工具 | 描述 |
|------|------|
| `measure_range` | 读取指定范围数据 |
| `measure_cells` | 读取单元格详细信息 |
| `measure_structure` | 读取工作表结构 |
| `list_sheets` | 返回工作表名称列表 |
| `peek_preview` | 快速预览前几行数据 |
| `detect_columns` | 检测列名和数据类型 |
| `search_content` | 搜索关键词 |
| `get_summary` | 获取文件摘要信息 |
| `get_sample_data` | 获取样本数据 |
| `manage_named_range` | 管理命名范围 |
| `manage_sheet_style` | 管理工作表样式 |
| `manage_sheet_visibility` | 管理工作表可见性 |

### 粟米章 (Grain) - 10 个工具

| 工具 | 描述 |
|------|------|
| `convert_format` | 转换数据格式 |
| `convert_type` | 转换数据类型 |
| `convert_unit` | 转换单位 |
| `transpose` | 转置数据 |
| `text_to_columns` | 文本分列 |
| `clean_data` | 数据清洗 |
| `transform_data` | 高级数据转换 |
| `fuzzy_match_columns` | 模糊匹配列名 |
| `auto_type_infer` | 自动类型推断 |
| `standardize_data` | 数据标准化 |
| `transform_pipeline` | 数据转换管道 |

### 衰分章 (Share) - 5 个工具

| 工具 | 描述 |
|------|------|
| `group_by` | 按字段分组 |
| `distribute` | 按比例分配 |
| `summarize` | 分组汇总 |
| `pivot_analysis` | 数据透视分析 |
| `subtotal` | 分类汇总 |

### 少广章 (Dimension) - 5 个工具

| 工具 | 描述 |
|------|------|
| `find_dimension` | 已知面积求边长 |
| `derive` | 反向推导 |
| `calculate` | 执行计算 |
| `solve_equation` | 解方程 |
| `auto_sum` | 自动求和 |

### 商功章 (Work) - 33 个工具

| 工具 | 描述 |
|------|------|
| `batch_execute` | 批量执行多个操作 |
| `batch_transform` | 批量转换 |
| `batch_validate` | 批量验证 |
| `create_pivot` | 创建数据透视表 |
| `format_range` | 格式化单元格 |
| `create_chart` | 创建图表 |
| `update_chart` | 更新图表 |
| `list_charts` | 列出所有图表 |
| `delete_chart` | 删除图表 |
| `create_advanced_chart` | 创建高级图表 |
| `manage_table` | 管理 Excel 表格 |
| `manage_comment` | 批注管理 |
| `freeze_panes` | 冻结窗格 |
| `set_auto_filter` | 设置自动筛选 |
| `advanced_filter` | 高级筛选 |
| `manage_row_column_visibility` | 管理行列可见性 |
| `create_mapping_template` | 创建数据映射模板 |
| `create_basic_report` | 生成基础报表 |
| `create_advanced_report` | 生成高级报表 |
| `fill_template` | 基于模板填充数据 |
| `split_sheet` | 拆分工作表 |
| `protect_workbook` | 保护工作簿 |
| `protect_sheet` | 保护工作表 |
| `unprotect_sheet` | 解除工作表保护 |
| `insert_excel_image` | 插入图片 |
| `group_rows` | 分组行 |
| `export_chart_as_image` | 导出图表为图片 |
| `pack_file` | 打包为 ZIP |
| `unpack_file` | 解包 ZIP |
| `set_print_area` | 设置打印区域 |
| `set_zoom` | 控制缩放 |
| `generate_summary_report` | 数据摘要报告 |
| `generate_diff_report` | 变化检测报告 |
| `manage_data_view` | 数据视图管理 |

### 均输章 (Transport) - 7 个工具

| 工具 | 描述 |
|------|------|
| `import_data` | 导入数据 |
| `export_data` | 导出数据 |
| `migrate` | 数据迁移 |
| `merge_files` | 合并文件 |
| `join_tables` | SQL 风格关联 |
| `batch_merge` | 多表批量合并 |
| `excel_to_markdown` | Excel 转 Markdown |

### 盈不足章 (Balance) - 7 个工具

| 工具 | 描述 |
|------|------|
| `validate_range` | 验证数据范围 |
| `validate_type` | 验证数据类型 |
| `validate_formula` | 验证公式正确性 |
| `validate_file` | 验证文件结构 |
| `data_quality_check` | 数据质量检测 |
| `excel_lint` | 代码审计 |
| `file_analyze` | 文件分析 |

### 方程章 (Equation) - 5 个工具

| 工具 | 描述 |
|------|------|
| `create_formula` | 创建公式 |
| `diagnose_formula` | 诊断公式错误 |
| `recalc_formulas` | 公式重算 |
| `generate_formula` | 生成常用公式 |
| `set_array_formula` | 设置数组公式 |

### 勾股章 (Triangle) - 6 个工具

| 工具 | 描述 |
|------|------|
| `analyze_stats` | 统计分析 |
| `analyze_trend` | 趋势分析 |
| `analyze_correlation` | 相关性分析 |
| `analyze_data` | 智能数据分析 |
| `visualize` | 数据可视化 |
| `visualize_data` | CSV 数据可视化 |
| `advanced_analysis` | 高级数据分析 |
| `variance_analysis` | 预算差异分析 |

---

## 相关文档

- [方田章工具](field.md) - 数据读取与单元格操作
- [粟米章工具](grain.md) - 格式转换与数据处理
- [衰分章工具](share.md) - 分组汇总与数据分析
- [少广章工具](dimension.md) - 反向计算与推导
- [商功章工具](work.md) - 批量操作与图表管理
- [均输章工具](transport.md) - 导入导出与数据迁移
- [盈不足章工具](balance.md) - 数据验证与审计
- [方程章工具](equation.md) - 公式计算与诊断
- [勾股章工具](triangle.md) - 数据分析与可视化
