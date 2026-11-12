#!/usr/bin/env python3
"""校验一套联通交付件：功能清单、需求说明书，以及可选的计数模型覆盖。"""

from __future__ import annotations

import argparse
import json
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn

from delivery_rules import CHAPTERS, check_rows, forbidden_in, is_entity, read_list


def read_doc(path: Path) -> list[tuple[str, str]]:
    """返回正文段落的（样式, 文字），不含目录。"""
    document = Document(path)
    items = []
    for element in document.element.body.iterchildren():
        if element.tag == qn("w:p"):
            style = element.find(".//" + qn("w:pStyle"))
            text = "".join(node.text or "" for node in element.iter(qn("w:t"))).strip()
            if text:
                items.append((style.get(qn("w:val")) if style is not None else "", text))
        elif element.tag == qn("w:tbl"):
            for node in element.iter(qn("w:t")):
                if node.text and node.text.strip():
                    items.append(("Table", node.text.strip()))
    return items


def check_doc(items: list[tuple[str, str]], system: str, rows: list[dict[str, str]]) -> list[str]:
    errors = []
    chapters = [text for style, text in items if style == "Heading1" and not text.startswith("附录")]
    expected = [f"{number} {title}" for number, title in enumerate(CHAPTERS, 1)]
    if chapters != expected:
        errors.append(f"章节骨架应为 {expected}，实际为 {chapters}")
    if f"8.1 {system}" not in [text for style, text in items if style == "Heading2"]:
        errors.append(f"8.1 的标题应与功能清单系统名称一致：{system!r}")
    modules = [text.split(" ", 1)[1] for style, text in items if style == "Heading3"]
    want_modules = list(dict.fromkeys(row["module"] for row in rows))
    if modules != want_modules:
        errors.append(f"第 8 章一级模块 {modules} 与功能清单 {want_modules} 不一致")
    functions, current = [], None
    for style, text in items:
        if style == "Heading4":
            current = {"name": text.split(" ", 1)[1], "lines": []}
            functions.append(current)
        elif style.startswith("Heading"):
            current = None
        elif current is not None:
            current["lines"].append(text)
    names = [function["name"] for function in functions]
    want_names = [row["name"] for row in rows]
    if names != want_names:
        for index in range(max(len(names), len(want_names))):
            a = names[index] if index < len(names) else None
            b = want_names[index] if index < len(want_names) else None
            if a != b:
                errors.append(f"第 8 章第 {index + 1} 条功能 {a!r} 与功能清单 {b!r} 不一致")
                break
    for function, row in zip(functions, rows):
        lines = function["lines"]
        marks = {"新增": "☑新增", "修改": "☑修改"}
        if not any(line.startswith("修改类型：") and marks[row["type"]] in line for line in lines):
            errors.append(f"{function['name']}: 修改类型应勾选 {row['type']}")
        if not any(line.startswith("功能描述：") for line in lines):
            errors.append(f"{function['name']}: 缺少功能描述")
        has_io = [any(line.startswith(prefix) for line in lines) for prefix in ("输入项：", "输出项：")]
        if is_entity(function["name"]) and any(has_io):
            errors.append(f"{function['name']}: 数据实体不写输入项、输出项")
        if not is_entity(function["name"]) and not all(has_io):
            errors.append(f"{function['name']}: 缺少输入项或输出项")
    for _, text in items:
        bad = forbidden_in(text)
        if bad:
            errors.append(f"需求说明书含禁用内容 {bad!r}：{text[:40]}")
    return errors


def check_coverage(count_model: dict, mapping: dict, rows: list[dict[str, str]]) -> list[str]:
    """计数模型里每个数据功能和事务功能都要映射到至少一条清单行。mapping: {计数编号: ["一级模块/二级模块", ...]}"""
    errors = []
    known = {f"{row['module']}/{row['name']}" for row in rows}
    for item in count_model.get("data_functions", []) + count_model.get("transactions", []):
        targets = mapping.get(item["id"]) or []
        if not targets:
            errors.append(f"{item['id']} {item['name']}: 没有映射到任何功能清单行")
        for target in targets:
            if target not in known:
                errors.append(f"{item['id']}: 映射目标不在功能清单中：{target}")
    for key in set(mapping) - {item["id"] for item in count_model.get("data_functions", []) + count_model.get("transactions", [])}:
        errors.append(f"映射里的 {key} 不在计数模型中")
    return errors


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("feature_list", type=Path)
    parser.add_argument("requirement_doc", type=Path)
    parser.add_argument("--count-model", type=Path, help="阶段一 计数模型.json；给出时必须同时给 --mapping")
    parser.add_argument("--mapping", type=Path, help="内部映射 JSON：{计数编号: [\"一级模块/二级模块\"]}")
    args = parser.parse_args()
    if bool(args.count_model) != bool(args.mapping):
        parser.error("--count-model 与 --mapping 必须同时给出")
    system, rows = read_list(args.feature_list)
    errors = check_rows(rows) + check_doc(read_doc(args.requirement_doc), system, rows)
    if args.count_model:
        errors += check_coverage(json.loads(args.count_model.read_text(encoding="utf-8")),
                                 json.loads(args.mapping.read_text(encoding="utf-8")), rows)
    if errors:
        parser.exit(1, "FAIL:\n" + "\n".join(errors) + "\n")
    print(f"OK: {len(rows)} 行功能清单与需求说明书第 8 章一致" + ("，计数模型全覆盖" if args.count_model else ""))


if __name__ == "__main__":
    main()
