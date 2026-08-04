# 触发条件设计指南

## 设计原则

### 1. 正面触发（Trigger When）
- 明确列出支持的文件格式
- 列出具体操作场景
- 包含隐式触发条件

### 2. 负面排除（Do NOT Trigger When）
- 明确列出不触发的场景
- 基于"最终产出"判断
- 避免边界模糊

### 3. 场景化描述
- 使用用户语言
- 包含具体示例
- 耆虑边缘情况

## 示例模板

```yaml
description: |
  Use this skill when:
  - The primary input or output is a [文件类型] file
  - The user wants to [操作1], [操作2], or [操作3]
  - The user references a [文件类型] file by name or path (even casually)
  
  Do NOT trigger when:
  - The primary deliverable is a [其他类型] document
  - The task is about [无关场景]
  - The user wants standalone [其他工具] scripts
```

## 常见触发词

### 文件格式
- .xlsx, .xlsm, .xls, .csv, .tsv
- spreadsheet, workbook, worksheet

### 操作动词
- create, make, generate
- open, read, load
- edit, modify, update
- fix, repair, debug
- convert, transform, export

### 场景描述
- "帮我做一个报表"
- "把这个数据放到 Excel 里"
- "检查一下这个表格有没有问题"
