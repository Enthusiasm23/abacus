# Chart Reference for Abacus

## Chart Creation

### From Range
```
abacus_create_chart(file, sheet, range, chart_type)
```
Best for: Simple data in worksheet ranges

### From PivotTable
```
abacus_create_chart(file, sheet, range, chart_type)  # with pivot table data range
```
Best for: PivotTable summary data

### From Table
```
abacus_create_chart(file, sheet, range, chart_type)
```
Best for: Excel Tables with structured references

## Chart Types

| Type | Value | Use Case |
|------|-------|----------|
| Column Clustered | `bar` | Comparing categories |
| Line | `line` | Trends over time |
| Pie | `pie` | Part-to-whole composition |
| Area | `area` | Cumulative trends |
| Scatter | `scatter` | Correlation between variables |

## Configuration

### Titles and Labels
- `title`: Chart title
- `x_axis`: X-axis label
- `y_axis`: Y-axis label

### Positioning
- `position`: Cell anchor (default "A1")
- `width`: Chart width in inches (default 15)
- `height`: Chart height in inches (default 10)

## Chart Management

### List Charts
```
abacus_list_charts(file, sheet)
```

### Update Chart
```
abacus_update_chart(file, sheet, chart_index, title="New Title")
```

### Delete Chart
```
abacus_delete_chart(file, sheet, chart_index)
```

## Common Workflows

### Create Chart with Formatting
```
1. abacus_create_chart(file, sheet, range, chart_type='bar', title='Monthly Sales', x_axis='Month', y_axis='Revenue')
2. abacus_update_chart(file, sheet, chart_index=0, title='Updated Title')
```

### Multi-Chart Layout
When creating dashboards with multiple charts, use explicit positioning:

```
Data at A1:D10. Place 4 charts in a 2×2 grid below data:

abacus_create_chart(..., position='A12')   # Top-left
abacus_create_chart(..., position='G12')   # Top-right
abacus_create_chart(..., position='A27')   # Bottom-left
abacus_create_chart(..., position='G27')   # Bottom-right
```

### Rules
- Leave at least 1-2 columns gap between charts
- Use `abacus_list_charts` to verify layout
- Verify no overlaps before creating additional charts

## Best Practices

1. **Use appropriate chart type**: Bar for comparison, line for trends, pie for composition
2. **Format numbers**: Ensure data is properly formatted before charting
3. **Keep it simple**: Avoid excessive gridlines or 3D effects
4. **Title everything**: Charts should be self-explanatory
5. **Consider data labels**: Add for key data points when precision matters
