# Conditional Format Reference for Abacus

## Rule Types

| Type | Description | Parameters |
|------|-------------|------------|
| `cell-value` | Format based on cell value comparison | operator + formula1 (+ formula2 for between) |
| `expression` | Format based on formula result | formula only |

## Operators (for cell-value type)

| Operator | Description | Formulas Required |
|----------|-------------|-------------------|
| `equal` | Cell equals value | formula1 |
| `not-equal` | Cell doesn't equal value | formula1 |
| `greaterThan` | Cell greater than value | formula1 |
| `lessThan` | Cell less than value | formula1 |
| `greaterThanOrEqual` | Cell greater or equal | formula1 |
| `lessThanOrEqual` | Cell less or equal | formula1 |
| `between` | Cell between two values | formula1 AND formula2 |
| `notBetween` | Cell not between two values | formula1 AND formula2 |

## Format Options

| Option | Type | Example |
|--------|------|---------|
| `interiorColor` | hex color | `FFFF00` (yellow fill) |
| `fontColor` | hex color | `FF0000` (red text) |
| `fontBold` | bool | `True` / `False` |
| `fontItalic` | bool | `True` / `False` |
| `borderStyle` | string | `thin`, `medium`, `thick` |
| `borderColor` | hex color | `000000` |

## Actions

### Add Rule (cell-value)

```
abacus_format_range(
    file='data.xlsx',
    sheet='Sheet1',
    range='A1:A10',
    conditional={'type': 'cell', 'operator': 'greaterThan', 'value': 100}
)
```

### Add Rule (expression)

```
abacus_format_range(
    file='data.xlsx',
    sheet='Sheet1',
    range='A1:D10',
    conditional={'type': 'expression', 'formula': '=$A1="Active"'}
)
```

### Clear Rules

```
abacus_clear_range(
    file='data.xlsx',
    sheet='Sheet1',
    range='A1:E100',
    clear_type='formats'
)
```

## Formula Notes

- For `cell-value` type: formula1/formula2 can be numbers, strings, or cell references
- For `expression` type: formula must return TRUE/FALSE
- Formulas use the top-left cell perspective (e.g., `=$A1>100` for relative rows)
- Use absolute references (`$A$1`) when comparing to a fixed cell

## Examples

### Highlight cells greater than 100
```
abacus_format_range(
    file='data.xlsx',
    sheet='Sheet1',
    range='A1:A10',
    conditional={'type': 'cell', 'operator': 'greaterThan', 'value': 100}
)
```

### Highlight cells between 50 and 100
```
abacus_format_range(
    file='data.xlsx',
    sheet='Sheet1',
    range='A1:A10',
    conditional={'type': 'cell', 'operator': 'between', 'value': 50, 'value2': 100}
)
```

### Highlight row if column A is "Active"
```
abacus_format_range(
    file='data.xlsx',
    sheet='Sheet1',
    range='A1:D10',
    conditional={'type': 'expression', 'formula': '=$A1="Active"'}
)
```

## Common Mistakes

| Mistake | Cause | Solution |
|---------|-------|----------|
| Rule not applying | Missing operator for cell-value | Always specify operator |
| Wrong highlighting | Missing `$` in formula | Use `$A1` for column, `$A$1` for fixed cell |
| Colors not working | Missing hex prefix | Use `FFFF00` not `#FFFF00` (openpyxl uses no `#`) |
| Rule applies to wrong cells | Range too broad | Use precise A1 notation |

## Best Practices

1. Test expression formulas in Excel first to verify logic
2. Clear existing rules before applying new ones if replacing
3. For row-based highlighting, apply rule to full range (not just one column)
4. Use relative row references (`$A1`) and absolute column references for row highlighting
5. Keep conditional formatting simple - complex rules slow down Excel
