"""Structural checks for the Unicom delivery document."""

import re


TOP_LEVEL = (
    "1 项目概述",
    "2 文档目的",
    "3 文档范围",
    "4 术语和缩写解释",
    "5 参考文档",
    "6 范围定义",
    "7 功能需求",
    "8 界面需求",
    "9 非功能性需求",
    "10 其它需求（可选）",
)
OVERVIEW_SECTIONS = (
    "1.1 项目背景",
    "1.2 系统总体功能结构",
    "1.3 往期建设情况",
    "1.4 本项目预算工作量",
    "1.5 本项目投资计划预算",
)
REQUIRED_CONTENT = (
    "1.1 项目背景",
    "1.2 系统总体功能结构",
    "1.3 往期建设情况",
    "1.4 本项目预算工作量",
    "1.5 本项目投资计划预算",
    "2 文档目的",
    "3 文档范围",
    "4 术语和缩写解释",
    "5 参考文档",
    "6 范围定义",
    "7.1 数据模型",
    "7.2 界面功能模块",
    "8 界面需求",
    "9 非功能性需求",
)
DATA_MODEL_LABELS = ("需求描述", "数据来源及流向", "加工逻辑")


def require(condition, message):
    if not condition:
        raise SystemExit(f"FAIL: {message}")


def compact(value):
    return re.sub(r"\s+", " ", str(value or "")).strip()


def section_level(section):
    value = section.get("level")
    require(isinstance(value, int), f"章节 {section.get('heading', '')} 缺少整数 level")
    return value


def section_text(section):
    parts = [section.get("heading", "")]
    parts.extend(section.get("paras", []))
    for table in section.get("tables", []):
        parts.extend(table.get("headers", []))
        for row in table.get("rows", []):
            parts.extend(row)
    for image in section.get("images", []):
        parts.append(image.get("caption", ""))
    return compact(" ".join(str(part) for part in parts if part))


def has_content(section):
    return bool(section_text(section)) or bool(section.get("tables")) or bool(section.get("images"))


def descendants(sections, index):
    base = section_level(sections[index])
    for later in sections[index + 1:]:
        if section_level(later) <= base:
            break
        yield later


def section_with_prefix(sections, prefix):
    normalized = compact(prefix)
    for section in sections:
        if compact(section.get("heading")).startswith(normalized):
            return section
    return None


def leading_ref(value):
    match = re.match(r"\s*(\d+(?:\.\d+)+)\b", str(value or ""))
    return match.group(1) if match else None


def mapped_section(sections, mapping, ident, minimum_level):
    ref = leading_ref(mapping.get("section_page_model"))
    require(ref, f"{ident} 的 section_page_model 未指向具体章节")
    section = section_with_prefix(sections, ref)
    require(section, f"{ident} 映射的章节不存在：{ref}")
    require(section_level(section) >= minimum_level, f"{ident} 映射到过高层章节：{section.get('heading')}")
    return section


def nonempty_paragraphs(sections):
    return sum(1 for section in sections for paragraph in section.get("paras", []) if compact(paragraph))


def check_document_model(model, document, mappings):
    require(isinstance(document, dict), "交付模型缺少 document")
    require(isinstance(mappings, list), "交付模型缺少 mappings")
    sections = document.get("sections")
    require(isinstance(sections, list) and sections, "delivery document 缺少 sections")
    headings = [compact(section.get("heading")) for section in sections]

    for heading in TOP_LEVEL:
        require(any(value.startswith(heading) for value in headings), f"缺少联通交付骨架章节：{heading}")
    for heading in OVERVIEW_SECTIONS:
        require(any(value.startswith(heading) for value in headings), f"缺少项目概述子章节：{heading}")

    for heading in REQUIRED_CONTENT:
        section = section_with_prefix(sections, heading)
        require(section, f"缺少必需章节：{heading}")
        nested = [section, *descendants(sections, sections.index(section))]
        require(any(has_content(item) for item in nested), f"必需章节没有内容：{heading}")

    section_7 = section_with_prefix(sections, "7 功能需求")
    require(section_7, "缺少章节：7 功能需求")
    for heading in ("7.1 数据模型", "7.2 界面功能模块"):
        require(section_with_prefix(sections, heading), f"缺少功能需求子章节：{heading}")

    mapping_by_id = {mapping.get("id"): mapping for mapping in mappings}
    for row in model["data_functions"]:
        ident = row["id"]
        mapping = mapping_by_id.get(ident)
        require(mapping, f"{ident} 缺少映射")
        section = mapped_section(sections, mapping, ident, 3)
        nested = [section, *descendants(sections, sections.index(section))]
        text = " ".join(section_text(item) for item in nested)
        for label in DATA_MODEL_LABELS:
            require(label in text, f"{ident} 数据模型缺少「{label}」")
        require(nonempty_paragraphs(nested) >= 3, f"{ident} 数据模型正文不足三段")

    for row in model["transactions"]:
        ident = row["id"]
        mapping = mapping_by_id.get(ident)
        require(mapping, f"{ident} 缺少映射")
        section = mapped_section(sections, mapping, ident, 3)
        require(nonempty_paragraphs([section]) >= 1, f"{ident} 缺少页面/后台功能正文")
        require(section.get("tables") or section.get("images"), f"{ident} 缺少接口或规则载体")

    return {"sections": len(sections), "headings": len(headings)}


def check_feature_granularity(features, transactions):
    require(isinstance(features, list) and features, "交付模型缺少 features")
    require(len(features) == len(transactions), "功能清单必须一条事务功能一行")
    transaction_ids = {row["id"] for row in transactions}
    seen = []
    for index, feature in enumerate(features, 1):
        remark = str(feature.get("remark", ""))
        found = [ident for ident in transaction_ids if ident in remark]
        require(len(found) == 1, f"功能清单第 {index} 行必须且只能对应一条事务功能")
        seen.append(found[0])
    require(sorted(seen) == sorted(transaction_ids), "功能清单未逐条覆盖全部事务功能")
    return {"rows": len(features)}
