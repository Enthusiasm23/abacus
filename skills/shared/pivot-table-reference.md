# Pivot Table Reference for Abacus

## Calculated Fields

PivotTable calculated fields work well for simple single-table formulas.

| Feature | Calculated Field | DAX Measure |
|---------|------------------|-------------|
| Single-table formulas | ✅ Works (`=Qty*Price`) | ✅ Works |
| Cross-table | NOT SUPPORTED | Full support |
| Complex logic | Limited | Full DAX |
| Reusable | Per PivotTable only | Across all PivotTables |

### Calculated Field Workflow

```
abacus_create_pivot(file, sheet, range, row_fields, value_field, agg_function='sum')
```

## PivotTable Source Types

| Source | Create Action | Supports Calculated Fields? |
|--------|---------------|----------------------------|
| Worksheet Range | `abacus_create_pivot` | YES - simple formulas |
| Excel Table | `abacus_create_pivot` | YES - structured references |

## Refresh Behavior

PivotTables do NOT auto-refresh when source data changes!

**After adding rows to source data:**
```
1. Add rows to source data
2. abacus_create_pivot (recreate) or abacus_pivot_analysis (new analysis)
```

## Field Configuration

### Row/Column/Value Fields

When creating PivotTables, configure fields:

1. **Row fields**: Categories to group by
2. **Value fields**: Numeric data to aggregate
3. **Aggregation function**: sum, avg, count, min, max

### Aggregation Functions

| Function | Use Case |
|----------|----------|
| sum | Totals (revenue, quantity) |
| count | Record counts |
| avg | Mean values |
| min/max | Extremes |

## Common Patterns

### Revenue Analysis

```
abacus_create_pivot(
    file='data.xlsx',
    sheet='Sales',
    range='A1:D100',
    row_fields=['Region', 'Product'],
    value_field='Revenue',
    agg_function='sum'
)
```

### Cross-Tabulation

```
abacus_create_pivot(
    file='data.xlsx',
    sheet='Sales',
    range='A1:D100',
    row_fields=['Region'],
    value_field='Amount',
    agg_function='sum'
)
```

## Pivot Analysis

For quick analysis without creating a new sheet:

```
abacus_pivot_analysis(
    file='data.xlsx',
    group_by='Category',
    value_field='Sales',
    agg_function='sum',
    output='result.xlsx'
)
```

## Common Errors and Solutions

| Error | Cause | Solution |
|-------|-------|----------|
| "Field not found" | Typo in field name | Check column headers with `abacus_measure_range` |
| Data doesn't update | Source changed without refresh | Recreate PivotTable |
| Wrong aggregation | Wrong agg_function | Specify correct function (sum/avg/count) |
| Empty results | Range includes empty rows | Use precise range in A1 notation |

## Best Practices

1. **Use Excel Tables as source**: Auto-expanding ranges simplify refresh
2. **Name fields clearly**: Use descriptive column headers in source data
3. **Choose right aggregation**: sum for amounts, count for records, avg for rates
4. **Verify data types**: Ensure numeric columns are numbers, not text
5. **Test with small data first**: Verify layout before large datasets
