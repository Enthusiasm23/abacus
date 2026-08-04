# DMV 查询参考

## TMSCHEMA 系列

| 查询 | 说明 |
|------|------|
| TMSCHEMA_MEASURES | 列出所有度量 |
| TMSCHEMA_TABLES | 列出所有表 |
| TMSCHEMA_COLUMNS | 列出所有列 |
| TMSCHEMA_RELATIONSHIPS | 列出所有关系 |

## DISCOVER 系列

| 查询 | 说明 |
|------|------|
| DISCOVER_CALC_DEPENDENCY | 计算依赖关系 |
| DISCOVER_METADATA | 元数据信息 |

## 使用示例

```sql
SELECT * FROM $SYSTEM.TMSCHEMA_MEASURES
SELECT * FROM $SYSTEM.DISCOVER_CALC_DEPENDENCY
```