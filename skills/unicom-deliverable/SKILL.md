---
name: unicom-deliverable
description: 按联通客户范例生成功能清单 xlsx 和需求说明书 docx，并配系统总体功能结构图。
disable-model-invocation: true
---

# 联通交付稿

一次调用处理一个系统，产出两份客户交付件和一张图：

- `功能清单.xlsx`：六列，一行一个功能或数据实体。
- `需求说明书.docx`：固定 12 章，第 8 章与功能清单逐行同名同序。
- `系统总体功能结构.drawio` 及导出的 png：放在 1.2 节。

这两份文件的读者是按 NESMA 评估功能规模的客户评估人。他们逐条对照清单和说明书，核对每个功能的输入项、输出项和外部接口。所以写法只有一个标准：**与客户范例的章节、句式和粒度一致**。交付件里不出现任何计数结果（功能点数、UFP、人月、TF/DF 编号）；阶段一的 `计数模型.json` 只用于内部核对覆盖。

## 工作流

先定清单，再写说明书。清单未经用户确认前不写说明书：说明书第 8 章是清单的展开，清单一改就得重生成。

1. 按 [启动访谈.md](references/启动访谈.md) 确认输入、输出位置和版式范本。
2. 读代码和业务材料，按 [功能清单规则.md](references/功能清单规则.md) 写 `rows.json`，运行 `build_feature_list.py` 生成清单。把清单交给用户确认；用户改过的清单是唯一事实，之后一律从 xlsx 读回，不再用自己的 `rows.json`。
3. 按 [需求说明书规则.md](references/需求说明书规则.md) 写 `spec.json`。
4. 按 [结构图与版式.md](references/结构图与版式.md) 画系统总体功能结构图，需要物理表附录时渲染卡片。
5. 运行 `build_requirement_doc.py` 生成 Word，再用 `refresh_toc.py` 刷新目录页码并导出 PDF。
6. 运行 `validate_delivery.py`，然后把 PDF 的封面、目录、1.2 节、第 8 章某页、第 10 章某页和附录某页渲染成图片逐张看：图片不得溢出页面，表格不得错位，不得有空白页。
7. 报告产物绝对路径、与代码不一致而改正的条目、没有核实的字段。

```bash
python <skill>/scripts/build_feature_list.py rows.json <系统名称> 功能清单.xlsx
python <skill>/scripts/build_structure_diagram.py diagram.json 系统总体功能结构.drawio
drawio -x -f png -s 2 -o 系统总体功能结构.png 系统总体功能结构.drawio
python <skill>/scripts/build_requirement_doc.py spec.json 功能清单.xlsx raw.docx
timeout 120 /usr/bin/python3 <skill>/scripts/refresh_toc.py raw.docx 需求说明书.docx 需求说明书.pdf
python <skill>/scripts/validate_delivery.py 功能清单.xlsx 需求说明书.docx [--count-model 计数模型.json --mapping 映射.json]
```

`refresh_toc.py` 需要带 UNO 的解释器（通常是系统的 `/usr/bin/python3`），其余脚本需要 `python-docx`、`openpyxl`、`Pillow`、`PyYAML`。

## 事实来源

- 功能是否存在、输入输出字段、校验规则：以被估系统的代码为准，前端没有调用的后端接口不列。
- 项目背景、往期建设、痛点、数字：只用用户给的材料里有依据的内容。材料里没有就不写，或向用户要；不得为了像范例而编造。
- 清单和说明书与代码冲突时改交付件，并在报告里列出改了什么。
- 核实不了的字段不写进输入项、输出项。

## 边界

- 默认写新的输出目录。只有用户明确要求时才覆盖已有交付件；覆盖前先把原文件复制到临时目录，并确认它在你读取之后没有被改动。
- 不把交付件或生成脚本写回被估系统的仓库。
- 派子代理写内容时，把范本、已定稿清单、规则文档的路径和只许写的目录交给它；它交回后抽查与代码的一致性再落盘。
- 交付件定稿前清除全部占位标记；脚本会拒绝含 `【`、`待确认`、`TODO` 的内容。
