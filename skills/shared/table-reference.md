# Table Reference for Abacus

## Table Operations

| Action | Description | Tool |
|--------|-------------|------|
| create | Create new table from range | `abacus_manage_table` |
| list | List all tables | `abacus_manage_table` |
| delete | Remove table (keeps data) | `abacus_manage_table` |
| append | Add rows to table | `abacus_manage_table` |

## Table Creation

```
abacus_manage_table(
    file='data.xlsx',
    sheet='Sheet1',
    action='create',
    table_name='SalesData',
    range='A1:D100',
    style='TableStyleMedium2'
)
```

## Table Styles

| Style | Description |
|-------|-------------|
| TableStyleLight1 | Minimal borders, no header fill |
| TableStyleLight2-21 | Various light themes |
| TableStyleMedium2 | Standard blue, most widely used |
| TableStyleMedium9 | Orange accent |
| TableStyleDark1-11 | Dark header with white text |

## Data Model Workflow

Excel Tables on worksheets are NOT automatically in the Data Model. To use with DAX:

1. Ensure data is formatted as an Excel Table
2. Add table to Data Model for DAX analysis
3. Then create DAX measures on it

**Why Tables over plain ranges:**
- Structured references: `=SUM(Sales[Amount])` instead of `=SUM(B2:B100)`
- Auto-expand when rows are added
- Built-in filtering, sorting, and banded rows
- Named reference for Power Query: `Excel.CurrentWorkbook(){[Name="SalesData"]}`

**When NOT to use Tables:**
- Single-cell parameters (use named ranges instead)
- Layout areas with merged cells
- Print-formatted reports with specific spacing

## Table Append

Add new rows to an existing table:

```
abacus_manage_table(
    file='data.xlsx',
    sheet='Sheet1',
    action='append',
    table_name='SalesData',
    data=[{'Name': 'Product A', 'Sales': 100}, {'Name': 'Product B', 'Sales': 200}]
)
```

## Common Mistakes

- Using `abacus_manage_table` on plain ranges (must be a ListObject first)
- Forgetting to specify `table_name` for append/delete operations
- Table names must be unique within workbook
- Style parameter applies table style, not cell formatting

## Tool Selection

| Goal | Tool |
|------|------|
| Create/manage worksheet tables | `abacus_manage_table` |
| Format table appearance | `abacus_manage_style` |
| Analyze table data | `abacus_create_pivot` or `abacus_pivot_analysis` |
| Convert range to table | `abacus_manage_table` (create) |
| Read table data | `abacus_measure_range` |
