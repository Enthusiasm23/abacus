# Changelog

## [1.0.0] - 2026-08-09

### 🎯 Abacus - 本地 Excel 自动化框架

基于《九章算术》的章节体系设计，提供完整的 Excel 数据处理能力。

---

#### 📊 核心能力（103 个）

**方田章 (Field)** - 数据读取与单元格操作
- measure_range, measure_cells, measure_structure
- list_sheets, peek_preview, detect_columns
- search_content, get_summary, get_sample_data, manage_named_range

**粟米章 (Grain)** - 格式转换与数据处理
- convert_format, convert_unit, convert_type, data_transform
- transpose, text_to_columns, fuzzy_match
- auto_type_infer, standardize, transform_pipeline

**衰分章 (Share)** - 分组汇总与数据分析
- group_by, distribute, summarize, subtotal

**少广章 (Dimension)** - 数学计算与反向推导
- find_dimension, derive, calculate
- solve_equation, auto_sum

**商功章 (Work)** - 批量操作与图表管理
- batch_execute, batch_transform, batch_validate
- create_chart, update_chart, list_charts, delete_chart, create_advanced_chart
- format_range, create_pivot, manage_comment, freeze_panes
- set_auto_filter, manage_visibility, protect_workbook, protect_sheet
- unprotect_sheet, set_array_formula, insert_image, group_rows
- export_chart_as_image, pack_file, unpack_file, set_print_area, set_zoom
- summary_report, diff_report, data_view, advanced_filter, create_mapping_template

**均输章 (Transport)** - 导入导出与数据迁移
- import_data, export_data, migrate, join_tables, merge_files
- excel_to_markdown, split_sheet

**盈不足章 (Balance)** - 数据验证与审计
- validate_range, validate_type, validate_formula
- data_validation, validate_file, quality_check
- excel_lint, file_analyze, validation_engine

**方程章 (Equation)** - 公式计算与诊断
- create_formula, diagnose_formula, recalc_formulas
- generate_formula, set_array_formula

**勾股章 (Triangle)** - 数据分析与可视化
- analyze_stats, analyze_trend, analyze_correlation, visualize
- analyze_data, visualize_data, variance_analysis, advanced_analysis

---

#### 🛠️ 技术特性

| 特性 | 说明 |
|------|------|
| 🐍 Python 3.10+ | 使用最新语法特性 |
| 📦 三入口 | CLI + MCP Server + Python 库 |
| 🧪 1136 个测试 | 测试覆盖率 88% |
| 🌐 中英文支持 | i18n 国际化 |
| 📝 结构化日志 | 可配置日志级别 |

---

#### 🔒 安全特性

- ✅ 路径遍历检查
- ✅ ReDoS 防护
- ✅ 公式注入防护
- ✅ 批量操作限制（MAX_OPERATIONS=1000）

---

#### 📚 文档

- 9 个章节 API 文档
- 103 个能力 SKILL.md
- 完整的使用示例
