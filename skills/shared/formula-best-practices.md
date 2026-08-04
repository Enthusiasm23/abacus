# 公式最佳实践

## 核心原则
**始终使用 Excel 公式，不要在 Python 中计算后硬编码值**

## 常用公式模板

### 求和
```
=SUM(A1:A100)
=SUMIF(B1:B100,"条件",A1:A100)
=SUMIFS(A1:A100,B1:B100,"条件1",C1:C100,"条件2")
```

### 查找
```
=VLOOKUP(查找值,范围,列号,FALSE)
=INDEX(范围,MATCH(查找值,查找范围,0))
=XLOOKUP(查找值,查找范围,返回范围)
```

### 条件判断
```
=IF(条件,真值,假值)
=IFS(条件1,值1,条件2,值2,TRUE,默认值)
=SWITCH(表达式,值1,结果1,值2,结果2,默认值)
```

### 日期
```
=TODAY()
=NOW()
=DATEDIF(开始日期,结束日期,"Y"/"M"/"D")
=EOMONTH(日期,月数)
```

### 文本
```
=LEFT(文本,字符数)
=RIGHT(文本,字符数)
=MID(文本,起始位置,字符数)
=SUBSTITUTE(文本,旧文本,新文本)
=CONCATENATE(文本1,文本2,...)
```

## AI 常见错误
1. QUARTER 函数不存在 → 用 ROUNDUP(MONTH()/3,0)
2. COUNTUNIQUE 不存在 → 用 SUMPRODUCT(1/COUNTIF(...))
3. AVERAGEIF 参数顺序错误
4. SUMIFS 条件语法错误