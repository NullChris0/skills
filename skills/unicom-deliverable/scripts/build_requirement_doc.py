#!/usr/bin/env python3
"""按验收通过的版式生成联通需求说明书。

版式全部取自 assets/需求说明书版式模板.docx 里带 {{名称}} 标记的原型段落；
本脚本只复制原型并替换文字，不自行定义字体、间距或编号。
"""

from __future__ import annotations

import argparse
import copy
import json
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn

from delivery_rules import CHAPTERS, forbidden_in, is_entity, read_list

TEMPLATE = Path(__file__).resolve().parent.parent / "assets" / "需求说明书版式模板.docx"
PAGE_WIDTH_EMU = 5274310      # 版心宽度
MAX_IMAGE_HEIGHT_EMU = 7200000
TEXT_LISTS = ("background", "history", "purpose", "scope", "references", "boundary", "interactions", "interfaces", "nonfunctional")
TEXT_FIELDS = ("title", "date", "structure_intro", "system_name", "ui_intro")
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"


def check_spec(spec: dict, system: str, rows: list[dict[str, str]]) -> list[str]:
    """返回 spec 的全部错误；空列表表示通过。"""
    errors = []
    for key in TEXT_FIELDS:
        if not (isinstance(spec.get(key), str) and spec[key].strip()):
            errors.append(f"{key} 必须是非空字符串")
    for key in TEXT_LISTS:
        value = spec.get(key)
        if not (isinstance(value, list) and value and all(isinstance(item, str) and item.strip() for item in value)):
            errors.append(f"{key} 必须是非空字符串数组")
    if not (spec.get("terms") and all(len(row) == 2 for row in spec["terms"])):
        errors.append("terms 必须是非空数组，每行 2 列")
    if not all(len(row) == 3 for row in spec.get("structure_table") or [[]]):
        errors.append("structure_table 每行 3 列")
    if errors:
        return errors
    if spec["system_name"] != system:
        errors.append(f"system_name 应与功能清单 A 列一致：{system!r}")
    modules = spec.get("modules") or []
    got = [(m["name"], f["name"], f.get("type", "新增")) for m in modules for f in m["functions"]]
    want = [(row["module"], row["name"], row["type"]) for row in rows]
    for index in range(max(len(got), len(want))):
        a = got[index] if index < len(got) else None
        b = want[index] if index < len(want) else None
        if a != b:
            errors.append(f"第 8 章与功能清单第 {index + 1} 条不一致：spec={a} 清单={b}")
            break
    if [m["name"] for m in modules] != [row[0] for row in spec["structure_table"]]:
        errors.append("structure_table 第一列必须与 modules 名称同序一致")
    for module in modules:
        for function in module["functions"]:
            where = f"{module['name']}/{function['name']}"
            if len(function.get("desc") or "") < 12:
                errors.append(f"{where}: desc 缺失或过短")
            entity = is_entity(function["name"])
            if entity and (function.get("input") or function.get("output")):
                errors.append(f"{where}: 数据实体不写 input/output")
            if not entity and not (function.get("input") and function.get("output")):
                errors.append(f"{where}: 功能必须同时有 input 和 output")
    for path, text in iter_text(spec):
        bad = forbidden_in(text)
        if bad:
            errors.append(f"{path}: 含禁用内容 {bad!r}")
    for key in ("structure_figure",):
        if spec.get(key) and not Path(spec[key]["image"]).is_file():
            errors.append(f"{key}: 图片不存在 {spec[key]['image']}")
    for key in ("figures", "appendix"):
        for item in spec.get(key) or []:
            if not Path(item["image"]).is_file():
                errors.append(f"{key}: 图片不存在 {item['image']}")
    return errors


def iter_text(value, path="spec"):
    """遍历 spec 里会写进文档的文字（图片路径除外）。"""
    if isinstance(value, str):
        yield path, value
    elif isinstance(value, list):
        for index, item in enumerate(value):
            yield from iter_text(item, f"{path}[{index}]")
    elif isinstance(value, dict):
        for key, item in value.items():
            if key != "image":
                yield from iter_text(item, f"{path}.{key}")


class Builder:
    def __init__(self, template: Path):
        self.doc = Document(template)
        self.body = self.doc.element.body
        children = list(self.body.iterchildren())
        self.sect = children[-1]
        if self.sect.tag != qn("w:sectPr"):
            raise ValueError("版式模板末尾缺少节属性")
        self.proto, pending = {}, None
        for element in children[:-1]:
            marker = self._marker(element)
            if marker:
                self.proto[marker] = element
            elif element.tag == qn("w:sdt"):
                self.proto["toc"] = element
            elif not self._text(element) and element.find(".//" + qn("a:blip")) is None:
                self.proto["empty"] = element
            self.body.remove(element)
        missing = {"cover1", "cover2", "cover3", "toc_title", "toc", "empty", "h1", "h2", "h3", "h4", "p1", "body", "fn",
                   "nf", "tbl3", "tbl2", "fig", "img", "cap", "h1_plain"} - set(self.proto)
        if missing:
            raise ValueError(f"版式模板缺少原型：{sorted(missing)}")
        self.image_id = 1000
        self.chapter = 0

    @staticmethod
    def _text(element) -> str:
        return "".join(node.text or "" for node in element.iter(qn("w:t")))

    def _marker(self, element) -> str | None:
        names = [e.get("name", "") for e in element.iter(qn("wp:docPr"))]
        for text in [self._text(element)] + names:
            if "{{" in text:
                return text.split("{{", 1)[1].split("}}", 1)[0]
        return None

    @staticmethod
    def _set_text(paragraph, *parts: str) -> None:
        runs = [run for run in paragraph.iter(qn("w:r")) if run.find(qn("w:t")) is not None]
        if len(parts) > len(runs):
            parts = ("".join(parts),)
        for run, part in zip(runs, parts):
            texts = run.findall(qn("w:t"))
            for extra in texts[1:]:
                run.remove(extra)
            texts[0].text = part
            texts[0].set(XML_SPACE, "preserve")
        for run in runs[len(parts):]:
            run.getparent().remove(run)

    def add(self, kind: str, *parts: str):
        element = copy.deepcopy(self.proto[kind])
        if parts:
            self._set_text(element, *parts)
        self.sect.addprevious(element)
        return element

    def heading1(self, title: str) -> None:
        self.chapter += 1
        self.add("h1", f"{self.chapter} ", title)

    def table(self, kind: str, header: list[str], rows: list[list[str]]) -> None:
        table = copy.deepcopy(self.proto[kind])
        head, sample = table.findall(qn("w:tr"))[:2]
        table.remove(sample)

        def fill(row, values):
            cells = row.findall(qn("w:tc"))
            if len(cells) != len(values):
                raise ValueError(f"表格列数不符：{values}")
            for cell, value in zip(cells, values):
                self._set_text(cell.find(qn("w:p")), value)

        fill(head, header)
        for values in rows:
            row = copy.deepcopy(sample)
            fill(row, values)
            table.append(row)
        self.sect.addprevious(table)

    def image(self, kind: str, path: str, caption: str) -> None:
        """kind=fig：图题与图片同段（界面截图）；kind=img：图片段后跟图题段。"""
        rid, image = self.doc.part.get_or_add_image(path)
        element = copy.deepcopy(self.proto[kind])
        cx = PAGE_WIDTH_EMU
        cy = int(cx * image.px_height / image.px_width)
        if cy > MAX_IMAGE_HEIGHT_EMU:
            cx, cy = int(cx * MAX_IMAGE_HEIGHT_EMU / cy), MAX_IMAGE_HEIGHT_EMU
        for extent in [e for e in element.iter(qn("wp:extent"))] + [e for e in element.iter(qn("a:ext")) if e.get("cx")]:
            extent.set("cx", str(cx))
            extent.set("cy", str(cy))
        next(element.iter(qn("a:blip"))).set(qn("r:embed"), rid)
        self.image_id += 1
        for node in list(element.iter(qn("wp:docPr"))) + list(element.iter(qn("pic:cNvPr"))):
            node.set("id", str(self.image_id))
            node.set("name", caption)
        if kind == "fig":
            self._set_text(element, caption)
        self.sect.addprevious(element)
        if kind == "img":
            self.add("cap", caption)


def build(spec: dict, feature_list: Path, output: Path, template: Path = TEMPLATE) -> dict[str, int]:
    system, rows = read_list(feature_list)
    errors = check_spec(spec, system, rows)
    if errors:
        raise ValueError("\n".join(errors))
    b = Builder(template)
    b.add("cover1", spec["title"])
    b.add("cover2", "需求说明书")
    b.add("cover3", spec["date"])
    b.add("toc_title", "目录")
    b.add("toc")
    b.add("empty")

    b.heading1(CHAPTERS[0])
    b.add("h2", "1.1 ", "项目背景")
    for text in spec["background"]:
        b.add("p1", text)
    b.add("h2", "1.2 ", "系统总体功能结构")
    b.add("p1", spec["structure_intro"])
    if spec.get("structure_figure"):
        b.image("img", spec["structure_figure"]["image"], spec["structure_figure"].get("caption", "系统总体功能结构图"))
    b.table("tbl3", ["一级模块", "主要功能", "使用角色"], spec["structure_table"])
    b.add("h2", "1.3 ", "往期建设情况")
    for text in spec["history"]:
        b.add("p1", text)
    b.add("h2", "1.4 ", "本项目预算工作量")
    b.add("p1", spec.get("budget_workload", "无。"))
    b.add("h2", "1.5 ", "本项目投资计划预算")
    b.add("p1", spec.get("budget_invest", "无。"))

    for title, key in zip(CHAPTERS[1:3], ("purpose", "scope")):
        b.heading1(title)
        for text in spec[key]:
            b.add("body", text)
    b.heading1(CHAPTERS[3])
    b.table("tbl2", ["术语或缩写", "定义"], spec["terms"])
    b.add("empty")
    for title, key in zip(CHAPTERS[4:7], ("references", "boundary", "interactions")):
        b.heading1(title)
        for text in spec[key]:
            b.add("body", text)

    b.heading1(CHAPTERS[7])
    b.add("h2", "8.1 ", spec["system_name"])
    for i, module in enumerate(spec["modules"], 1):
        b.add("h3", f"8.1.{i} ", module["name"])
        for j, function in enumerate(module["functions"], 1):
            b.add("h4", f"8.1.{i}.{j} ", function["name"])
            kind = function.get("type", "新增")
            b.add("fn", "修改类型：" + " ".join(("☑" if kind == mark else "□") + mark for mark in ("新增", "修改", "删除")))
            b.add("fn", "功能描述：" + function["desc"])
            if function.get("input"):
                b.add("fn", "输入项：" + function["input"])
            if function.get("output"):
                b.add("fn", "输出项：" + function["output"])

    b.heading1(CHAPTERS[8])
    for text in spec["interfaces"]:
        b.add("fn", text)
    b.heading1(CHAPTERS[9])
    b.add("fn", spec["ui_intro"])
    for figure in spec.get("figures") or []:
        b.image("fig", figure["image"], figure["caption"])
    b.heading1(CHAPTERS[10])
    for text in spec["nonfunctional"]:
        b.add("nf", text)
    b.add("empty")
    b.heading1(CHAPTERS[11])
    for text in spec.get("other") or ["无。"]:
        b.add("nf", text)
    if spec.get("appendix"):
        b.add("h1_plain", "附录 A 物理表清单")
        for item in spec["appendix"]:
            b.image("img", item["image"], item["caption"])
    properties = b.doc.core_properties
    properties.title, properties.subject = spec["title"], "需求说明书"
    b.doc.save(output)
    return {"modules": len(spec["modules"]), "functions": len(rows),
            "figures": len(spec.get("figures") or []), "appendix": len(spec.get("appendix") or [])}


def main() -> None:
    parser = argparse.ArgumentParser(description="按验收版式生成需求说明书（目录页码需再用 refresh_toc.py 刷新）。")
    parser.add_argument("spec", type=Path)
    parser.add_argument("feature_list", type=Path, help="已定稿的功能清单 xlsx；第 8 章与它逐行比对")
    parser.add_argument("output", type=Path)
    parser.add_argument("--template", type=Path, default=TEMPLATE)
    args = parser.parse_args()
    try:
        stats = build(json.loads(args.spec.read_text(encoding="utf-8")), args.feature_list, args.output, args.template)
    except ValueError as error:
        parser.exit(1, f"FAIL:\n{error}\n")
    print(f"OK {args.output}: {stats['modules']} 个模块, {stats['functions']} 条功能, "
          f"界面图 {stats['figures']}, 附录表 {stats['appendix']}")


if __name__ == "__main__":
    main()
