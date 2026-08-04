# CRITICAL RULES

## 必须遵守的规则

1. **NEVER** 硬编码计算值 - 始终使用 Excel 公式
2. **ALWAYS** 验证文件路径和工作表名称
3. **NEVER** 修改源文件而不创建备份
4. **ALWAYS** 关闭工作簿以释放资源
5. **NEVER** 忽略异常 - 捕获并转换为结构化错误
6. **ALWAYS** 使用 A1 表示法（如 A1:D10）
7. **NEVER** 假设工作表存在 - 先检查
8. **ALWAYS** 返回结构化的 JSON 结果
9. **NEVER** 在日志中泄露敏感数据
10. **ALWAYS** 提供有意义的错误消息

## 新增规则（来自 sbroenne-excel）

11. **ALWAYS** 使用 try/finally 确保工作簿关闭
12. **NEVER** 在循环中创建 COM 对象
13. **ALWAYS** 批量写入时使用手动计算模式
14. **NEVER** 跳过 Power Query 测试步骤
15. **ALWAYS** 修改数据后刷新数据模型
16. **NEVER** 假设格式会自动继承
17. **ALWAYS** 多图表布局时明确指定位置
18. **NEVER** 忽略重叠警告
19. **ALWAYS** 删除前验证对象存在
20. **NEVER** 盲目重试失败操作

## 财务模型规则（来自 anthropic-skills）

11. **ALWAYS** 使用 Excel 公式，不要硬编码计算值
12. **NEVER** 在 Python 中计算后写入静态值（除非必要）
13. **ALWAYS** 将假设放在独立单元格，使用引用
14. **NEVER** 在公式中硬编码数字（使用单元格引用）
15. **ALWAYS** 遵循颜色编码标准（蓝=输入，黑=公式，绿=跨表，红=外部）
16. **NEVER** 忽略公式验证（使用 recalc.py 或类似工具）
17. **ALWAYS** 测试 2-3 个样本引用后再构建完整模型
18. **NEVER** 假设列映射正确（验证列号）
19. **ALWAYS** 检查 NaN 和除零错误
20. **NEVER** 跳过重算步骤（公式需要计算才能得到正确值）