# Range Reference for Abacus

## Formatting Split Across Tools

| Use | Tool | Action | When |
|-----|------|--------|------|
| Number display format | `abacus_convert_format` | format_type | Dates, currency, percentages |
| Type conversion | `abacus_convert_type` | target_type | Text to number, etc. |
| Cell formatting | `abacus_format_range` | font/fill/border | Bold, colors, borders |
| Auto-fit columns | `abacus_manage_size` | auto | Fit to content |
| Named styles | `abacus_manage_style` | apply_header/apply_kpi | Semantic styling |
| Find and replace | `abacus_find_replace` | find/replace | Text substitution |

## Quick Pattern: Write, Format, Auto-Fit

```
1. abacus_batch_execute (write data to cells)
2. abacus_convert_format (apply number format)
3. abacus_manage_size (auto-fit columns)
```

## Quick Pattern: Header Row With Fill Colour

```
abacus_format_range(
    file='data.xlsx',
    sheet='Sheet1',
    range='A1:D1',
    font={'bold': True, 'color': 'FFFFFF'},
    fill={'color': '4472C4', 'pattern_type': 'solid'},
    alignment={'horizontal': 'center'}
)
```

## Quick Pattern: Semantic Status Cells

```
abacus_format_range(
    file='data.xlsx',
    sheet='Sheet1',
    range='B2:B10',
    fill={'color': 'C6EFCE', 'pattern_type': 'solid'}  # Green for good
)
abacus_format_range(
    file='data.xlsx',
    sheet='Sheet1',
    range='C2:C10',
    fill={'color': 'FFC7CE', 'pattern_type': 'solid'}  # Red for bad
)
```

## Format Codes

| Type | Code | Example |
|------|------|---------|
| Number | `#,##0.00` | 1,234.56 |
| Dollar | `$#,##0.00` | $1,234.56 |
| Euro | `€#,##0.00` | €1,234.56 |
| Yen | `¥#,##0` | ¥1,235 |
| Percent | `0.00%` | 12.34% |
| Date (ISO) | `yyyy-mm-dd` | 2023-03-15 |
| Date (US) | `mm/dd/yyyy` | 03/15/2023 |
| Date (EU) | `dd/mm/yyyy` | 15/03/2023 |
| Time | `h:mm AM/PM` | 2:30 PM |
| Time (24h) | `hh:mm:ss` | 14:30:00 |
| Text | `@` | (as-is) |

## format-range Properties

| Property | Type | Example |
|----------|------|---------|
| `bold` | bool | `True` |
| `italic` | bool | `True` |
| `underline` | bool | `True` |
| `size` | number | `14` |
| `name` | string | `"Calibri"` |
| `color` (font) | hex | `"FFFFFF"` |
| `color` (fill) | hex | `"4472C4"` |
| `horizontal` | string | `"center"`, `"left"`, `"right"` |
| `vertical` | string | `"middle"`, `"top"`, `"bottom"` |
| `wrap_text` | bool | `True` |
| `style` (border) | string | `"thin"`, `"medium"`, `"thick"` |

## Related Tools

| Tool | Purpose |
|------|---------|
| `abacus_format_range` | Apply fills, fonts, borders, alignment |
| `abacus_convert_format` | Number display formats |
| `abacus_manage_size` | Auto-fit columns/rows |
| `abacus_find_replace` | Find and replace text |
| `abacus_manage_style` | Apply semantic styles |
| `abacus_clear_range` | Clear contents/formats |
| `abacus_copy_range` | Copy range data |
