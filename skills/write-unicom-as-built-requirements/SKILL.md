---
name: write-unicom-as-built-requirements
description: 从阶段一 NESMA 计数模型直接生成联通需求说明书 Word、功能清单 XLSX 和接口表图片。
disable-model-invocation: true
---

# 联通竣工态需求说明书

阶段二只接受阶段一产出的 `计数模型.json`。`功能点计数证据.xlsx` 用于人读核对，不能成为第二套计数事实。旧交付件和旧 `api-tables` 不是输入。

一次调用只处理一个计数对象，在新目录生成：

- `映射表.xlsx`
- `需求说明书.docx`
- `功能清单.xlsx`
- `接口表/`：YAML、TeX 卡片源和图片
- `delivery-model.json`：阶段二内部模型，可选保留

不重新计数，不改数据功能/事务功能类型，不改 FP，不增加假设。

## 工作流

1. 读取 [启动访谈.md](references/启动访谈.md)，确认计数模型路径、代码基线、业务材料、输出目录和模板。用户确认后固定在证据基线上。
2. 校验计数模型：JSON 合法、四表一致、无 `assumptions`/`unresolved`、每个正式行有证据和用户确认。发现计数问题返回阶段一，不在本阶段修正。
3. 按 [交付模型契约.md](references/交付模型契约.md) 在 `delivery-model.json` 中建立映射：每条 NESMA 记录映射到功能清单行、正文章节/页面/数据模型和接口 YAML。具体条目数量取自 JSON，不写死。
4. 在用户授权的只读范围内检查后端控制器/服务/DTO/DDL和前端路由/页面/API，补足业务描述、数据模型、页面行为和接口契约。仓库较大时使用子代理并行盘点，主 agent 负责合并和最终判断。
5. 按 [接口表YAML规范.md](references/接口表YAML规范.md) 生成接口 YAML。YAML 是本 Skill 的编辑面；TeX 和图片是投影。接口表只写代码基线支持的契约，不根据旧交付件补字段或错误码。
6. 按 [成稿与校验.md](references/成稿与校验.md) 建立 `delivery-model.json`，再用 `python <skill-directory>/scripts/build_delivery.py <delivery-model.json> <输出目录>` 生成 DOCX、功能清单、映射表、YAML、TeX 卡片和图片。接口表以图片插入 Word，保留图题、编号和正文引用。
7. 运行 `python <skill-directory>/scripts/validate_delivery.py <计数模型.json> <输出目录>`，并核对：计数与映射全覆盖、功能清单与正文一致、YAML 与代码基线一致、每张卡片有图片、Word 可打开且无旧 UFP/旧编号、数据割接等非计数工作未进入 FP。通过后报告绝对路径。

## 边界

- 只写新输出目录，目录必须不存在或为空；不覆盖旧交付件。
- 不把交付件、接口 YAML 或生成脚本写回被估后端仓库。
- Word 是主交付格式，XLSX 直接生成；TeX 只用于接口卡片渲染，不作为整本文档的中间格式。
- 数据割接、迁移、环境切换、线下操作属于非计数支撑工作，在文档中描述，不进入 UFP。
- 独立重渲染已有接口 YAML 使用 `python <skill-directory>/scripts/render_cards.py <输出目录>/接口表`。
