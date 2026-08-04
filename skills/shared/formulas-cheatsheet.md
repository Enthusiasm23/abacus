# 公式速查表

## 查找函数

### XLOOKUP（推荐）
```
=XLOOKUP(查找值, 查找范围, 返回范围, [未找到时的值])
=XLOOKUP(D2, A:A, B:B, "未找到")
```

### INDEX + MATCH（兼容性好）
```
=INDEX(返回范围, MATCH(查找值, 查找范围, 0))
=INDEX(B:B, MATCH(D2, A:A, 0))
```

### VLOOKUP（传统）
```
=VLOOKUP(查找值, 表范围, 列号, FALSE)
=VLOOKUP(D2, A:B, 2, FALSE)
```

## 条件求和

### SUMIF / SUMIFS
```
=SUMIF(条件范围, 条件, 求和范围)
=SUMIFS(求和范围, 条件范围1, 条件1, 条件范围2, 条件2)
=SUMIFS(C:C, A:A, "北京", B:B, ">1000")
```

### COUNTIF / COUNTIFS
```
=COUNTIF(范围, 条件)
=COUNTIFS(A:A, "北京", B:B, ">1000")
```

### AVERAGEIF / AVERAGEIFS
```
=AVERAGEIF(条件范围, 条件, 平均范围)
=AVERAGEIFS(C:C, A:A, "北京", B:B, ">1000")
```

## 日期函数

### 当前日期时间
```
=TODAY()        # 当前日期
=NOW()          # 当前日期时间
```

### 日期差
```
=DATEDIF(开始日期, 结束日期, "Y")  # 年数
=DATEDIF(开始日期, 结束日期, "M")  # 月数
=DATEDIF(开始日期, 结束日期, "D")  # 天数
```

### 月末日期
```
=EOMONTH(日期, 0)    # 本月月末
=EOMONTH(日期, 1)    # 下月月末
```

### 工作日
```
=NETWORKDAYS(开始日期, 结束日期)  # 工作日天数
=WORKDAY(开始日期, 天数)          # N个工作日后的日期
```

## 文本函数

### 提取
```
=LEFT(文本, 字符数)      # 左侧提取
=RIGHT(文本, 字符数)     # 右侧提取
=MID(文本, 起始位置, 字符数)  # 中间提取
```

### 查找替换
```
=FIND(查找文本, 文本)    # 查找位置（区分大小写）
=SEARCH(查找文本, 文本)  # 查找位置（不区分大小写）
=SUBSTITUTE(文本, 旧文本, 新文本)  # 替换
```

### 连接
```
=CONCATENATE(文本1, 文本2, ...)
=TEXTJOIN(",", TRUE, 范围)  # 用逗号连接
```

## 逻辑函数

### 条件判断
```
=IF(条件, 真值, 假值)
=IFS(条件1, 值1, 条件2, 值2, TRUE, 默认值)
=SWITCH(表达式, 值1, 结果1, 值2, 结果2, 默认值)
```

### 错误处理
```
=IFERROR(公式, 错误时的值)
=IFNA(公式, #N/A时的值)
```

## 数组函数（Excel 365）

### 去重排序筛选
```
=UNIQUE(范围)           # 去重
=SORT(范围, 列号, 1)    # 排序
=FILTER(范围, 条件)     # 筛选
```

### 序列
```
=SEQUENCE(行数, 列数, 起始值, 步长)
```

## AI 常见错误

1. **QUARTER 不存在**
   - 错误：`=QUARTER(A1)`
   - 正确：`=ROUNDUP(MONTH(A1)/3, 0)`

2. **COUNTUNIQUE 不存在**
   - 错误：`=COUNTUNIQUE(A:A)`
   - 正确：`=COUNTA(UNIQUE(A:A))`

3. **SUMIFS 条件语法**
   - 错误：`=SUMIFS(C:C, A:A = "北京", B:B > 1000)`
   - 正确：`=SUMIFS(C:C, A:A, "北京", B:B, ">1000")`

4. **AVERAGEIF 参数顺序**
   - 错误：`=AVERAGEIF(C:C, A:A, "北京")`
   - 正确：`=AVERAGEIF(A:A, "北京", C:C)`
