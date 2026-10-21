---
name: nesma-count
description: 从已确认证据生成四表 NESMA 计数包和供成稿使用计数模型。
disable-model-invocation: true
---

# NESMA 功能点计数

一次调用只处理一个计数对象。最终只交付：

- `功能点计数证据.xlsx`
- `计数模型.json`

不交付 Markdown 审阅报告、Word、成本、工期、SNAP、调整因子或客户终稿。最终包内不得存在“假设”；所有判断必须由证据或用户确认解决。

## 工作流

1. 读取并完成 [启动访谈.md](references/启动访谈.md)。先做浅层只读侦察，再把全部启动参数一次问清；用户确认完整摘要后才进入正式计数。
2. 读取 [证据审查.md](references/证据审查.md)、[计数对象与方法选择.md](references/计数对象与方法选择.md) 和适用的计数规则：
   - 全量：[通用计数规则.md](references/通用计数规则.md)
   - 概要：[概要FPA.md](references/概要FPA.md)
   - 详细：[详细FPA-数据功能.md](references/详细FPA-数据功能.md)、[详细FPA-事务功能.md](references/详细FPA-事务功能.md)
3. 在用户授权范围内只读检查代码、文档、接口、数据结构和运行证据。仓库较大时使用子代理并行盘点，但由主 agent 汇总、判断和写产物。
4. 建立候选数据功能、事务功能、证据索引和未决队列。凡证据不能唯一决定的拆合、类型、FTR、DET 或基本过程边界，按 [计数判别访谈.md](references/计数判别访谈.md) 批量询问用户；不得静默选边或写入假设。
5. 用户确认后的判断写成证据清单中的“用户确认”记录。用户无法确认的项目不进入正式行；写入 `excluded` 并说明排除理由。
6. 按 [复核与输出.md](references/复核与输出.md) 组装 `计数模型.json`，运行：

   ```bash
   python <skill-directory>/scripts/generate_count_package.py <model.json> <new-output-directory>
   ```

7. 验证输出目录只有 `计数模型.json` 和 `功能点计数证据.xlsx`，工作簿恰好包含 `汇总`、`数据功能`、`事务功能`、`证据清单` 四张表；所有编号、RET/DET/FTR、FP、证据引用和汇总数可复算。完成后报告绝对路径和 UFP，不设置额外复核关卡。

## 契约边界

当前四表契约支持 `新项目` 和 `存量应用`。增强项目需要扩展 `counting`、`change` 和变更前规模字段，暂不在本契约内静默兼容。

输出目录必须不存在或为空。重跑使用新版本目录，永不覆盖旧包。
