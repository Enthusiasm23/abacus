# Worksheet Reference for Abacus

## Same-File Operations

Use these tools for worksheet lifecycle within the same workbook:

| Operation | Tool | Key Parameters |
|-----------|------|----------------|
| Create sheet | `abacus_batch_execute` | Write to new sheet name |
| Read structure | `abacus_measure_structure` | file, sheet |
| Rename sheet | `abacus_manage_sheet_visibility` | N/A - use openpyxl directly |
| Copy range | `abacus_copy_range` | source, target |
| Delete content | `abacus_clear_range` | range, clear_type |

## Worksheet Visibility

| Action | Tool | Description |
|--------|------|-------------|
| Show sheet | `abacus_manage_sheet_visibility` | action='show' |
| Hide sheet | `abacus_manage_sheet_visibility` | action='hide' |
| Very hidden | `abacus_manage_sheet_visibility` | action='very-hide' |
| Get status | `abacus_manage_sheet_visibility` | action='get' |

## Sheet Style

| Action | Tool | Description |
|--------|------|-------------|
| Set tab color | `abacus_manage_sheet_style` | action='set', color='FF0000' |
| Get tab color | `abacus_manage_sheet_style` | action='get' |
| Clear tab color | `abacus_manage_sheet_style` | action='clear' |

## Freeze Panes

```
abacus_freeze_panes(
    file='data.xlsx',
    sheet='Sheet1',
    rows=1,           # Freeze first row
    columns=1         # Freeze first column
)
```

## Positioning

Use `before_sheet` or `after_sheet` (not both) to control sheet order:

- `before_sheet: "Sheet1"` - Insert before Sheet1
- `after_sheet: "Sheet1"` - Insert after Sheet1
- Neither specified - Append to end

## Common Operations

### Read Sheet Structure
```
abacus_measure_structure(file='data.xlsx', sheet='Sheet1')
```
Returns: row count, column count, merged cells, dimensions

### Copy Range Between Sheets
```
abacus_copy_range(
    file='data.xlsx',
    sheet='Source',
    range='A1:D10',
    target='Destination!A1'
)
```

### Clear Sheet Content
```
abacus_clear_range(
    file='data.xlsx',
    sheet='Sheet1',
    range='A1:Z100',
    clear_type='contents'  # or 'formats' or 'all'
)
```

## Common Errors

| Error | Cause | Solution |
|-------|-------|----------|
| "Sheet not found" | Typo in sheet name | Use `abacus_measure_structure` to see available sheets |
| "Range invalid" | Wrong A1 notation | Check with `abacus_measure_range` |
| "File not found" | Wrong path | Verify file path with `Test-Path` |
| "Merged cell conflict" | Range overlaps merged cells | Use `abacus_measure_structure` to check merged cells |

## Best Practices

1. **Verify sheet exists** before operations
2. **Use A1 notation** for ranges (e.g., `A1:D10`, not `A:D`)
3. **Check merged cells** before writing to ranges
4. **Save after changes** - openpyxl changes are in-memory only
5. **Use freeze panes** for large datasets to keep headers visible
