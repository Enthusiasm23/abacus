# Behavioral Rules for Abacus Excel Operations

These rules ensure efficient and reliable Excel automation with Abacus (openpyxl-based).

## Core Execution Rules

### Execute Immediately

Do NOT ask clarifying questions for standard operations. Proceed with reasonable defaults:

- **File creation**: Create the file and report the path
- **Data operations**: Execute the operation and report results
- **Formatting**: Apply formatting and confirm completion

**When to ask**: Only when the request is genuinely ambiguous (e.g., "update the data" without specifying what data or which file).

### Discover Information First

Before asking the user, discover the information yourself:

| Bad (Asking) | Good (Discovering) |
|--------------|-------------------|
| "Which Excel file should I use?" | List files in the working directory |
| "What's the table name?" | Use `abacus_manage_table` (list) to discover tables |
| "Which sheet has the data?" | Use `abacus_measure_structure` to check all sheets |
| "What values should I filter?" | Read the data first, then filter appropriately |

### Format Professionally

When creating or modifying Excel files:

- Set appropriate column widths for content
- Apply header formatting (bold, filters)
- Use proper number formats (currency, dates, percentages)
- Auto-fit variable-width data
- Format data as Excel Tables (not plain ranges)

**Format code examples (auto-translated to locale):**

| Data Type | Format Code | Result |
|-----------|-------------|--------|
| USD | `$#,##0.00` | $1,234.56 |
| EUR | `€#,##0.00` | €1,234.56 |
| Number | `#,##0.00` | 1,234.56 |
| Percent | `0.00%` | 15.00% |
| Date (ISO) | `yyyy-mm-dd` | 2025-01-22 |
| Date (US) | `mm/dd/yyyy` | 01/22/2025 |

**Workflow:**
```
1. abacus_measure_range / abacus_batch_execute (write data)
2. abacus_convert_format (apply number format)
3. abacus_manage_size (auto-fit columns)
```

### Format Tabular Data as Excel Tables

Always convert tabular data to Excel Tables using `abacus_manage_table` (create):

```
1. abacus_batch_execute (write data including headers)
2. abacus_manage_table(action='create', table_name='SalesData', range='A1:D100')
```

**Why Tables over plain ranges:**
- Structured references: `=SUM(Sales[Amount])` instead of `=SUM(B2:B100)`
- Auto-expand when rows are added
- Built-in filtering, sorting, and banded rows
- Named reference for Power Query

**When NOT to use Tables:**
- Single-cell parameters (use named ranges instead)
- Layout areas with merged cells
- Print-formatted reports with specific spacing

### Report Results

After completing operations, report:

- What was created/modified
- File path (for new files)
- Any relevant statistics (row counts, etc.)

### Always End With a Text Response

**NEVER end your turn with only a tool call or command execution.** After all operations complete, provide a text summary.

| Bad (Silent completion) | Good (Text summary) |
|------------------------|--------------------|
| *(tool call with no text)* | "Created PivotTable 'SalesPivot' with tabular layout on the Analysis sheet." |
| *(just runs a command)* | "Applied currency format to column B and auto-fitted all columns." |

## Data Modification Rules

### Verify Before Delete

Before deleting tables, worksheets, or named ranges:

1. List existing items first
2. Confirm the exact name exists
3. Delete the specified item

**Why**: Delete operations cannot be undone in openpyxl without reloading.

### Targeted Updates Over Wholesale Replace

When updating data:

- **Prefer**: `abacus_batch_execute` with specific cell ranges
- **Avoid**: Deleting and recreating entire structures

**Why**: Targeted updates preserve formatting, formulas, and references.

### Save Explicitly

Changes to openpyxl workbooks are in-memory until saved:

- Use `wb.save(path)` to persist changes
- Operations modify the in-memory workbook
- Exit without save loses all changes

## Format Results as Tables

When presenting data to users, format as Markdown tables:

```markdown
| Column A | Column B | Column C |
|----------|----------|----------|
| Value 1  | Value 2  | Value 3  |
```

NOT as raw JSON arrays: `[["Column A","Column B"],["Value 1","Value 2"]]`

## Error Handling Rules

### Interpret Error Messages

When an operation fails:

1. Read the error message carefully
2. Check prerequisites (file exists, sheet exists, range valid)
3. Retry with corrected parameters

Do NOT immediately re-run the same failing command.

### Report Failures Clearly

When operations fail:

- State what was attempted
- Explain what went wrong
- Suggest the corrective action

**Good**: "Failed to create PivotTable: Source range 'A1:D100' does not exist on sheet 'Data'. Verify the range with `abacus_measure_structure` first."

**Bad**: "An error occurred."
