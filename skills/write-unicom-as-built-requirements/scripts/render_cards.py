#!/usr/bin/env python3
"""Render interface YAML files as TeX cards and PNG images."""

from __future__ import annotations

import argparse
import json
import os
import re
import shutil
import subprocess
import tempfile
from pathlib import Path
from typing import Any

import yaml
from PIL import Image, ImageChops


COMMON_FIELDS = ("id", "kind", "title")
ENVELOPES = {"ajax-result", "table-data-info", "file-stream"}
DEFAULT_STYLE = Path(__file__).resolve().parent.parent / "assets" / "api-card.sty"
TEX_ESCAPES = {
    "\\": r"\textbackslash{}",
    "&": r"\&",
    "%": r"\%",
    "$": r"\$",
    "#": r"\#",
    "{": r"\{",
    "}": r"\}",
    "~": r"\textasciitilde{}",
    "^": r"\textasciicircum{}",
    "<": r"\textless{}",
    ">": r"\textgreater{}",
    "_": r"\_",
}


def fail(message: str) -> None:
    raise ValueError(message)


def required(record: dict[str, Any], key: str, source: Path) -> Any:
    value = record.get(key)
    if value is None or value == "" or value == []:
        fail(f"{source}: 缺少必需字段 {key}")
    return value


def as_text(value: Any) -> str:
    return re.sub(r"\s+", " ", str(value)).strip()


def tex_escape(value: Any) -> str:
    return "".join(TEX_ESCAPES.get(char, char) for char in as_text(value))


def tex_name(value: Any) -> str:
    return tex_escape(value).replace(r"\_", r"\_\allowbreak{}")


def tex_id(value: Any, source: Path) -> str:
    ident = as_text(value)
    if not re.fullmatch(r"[A-Za-z0-9_.:-]+", ident):
        fail(f"{source}: id 只能包含字母、数字、下划线、点、冒号或连字符")
    return ident


def field_rows(fields: list[Any], source: Path) -> str:
    lines = []
    for index, field in enumerate(fields):
        if isinstance(field, str):
            field = {"name": field}
        elif isinstance(field, (list, tuple)):
            field = {
                "name": field[0] if field else None,
                "type": field[1] if len(field) > 1 else "",
                "label": field[2] if len(field) > 2 else "",
            }
        if not isinstance(field, dict) or not field.get("name"):
            fail(f"{source}: fields/request/response/columns 项必须是名称字符串或含 name 的对象")
        name = tex_name(field["name"])
        field_type = field.get("type") or field.get("java_type") or ""
        label = tex_escape(field.get("label") or field_type or field["name"])
        description = []
        if field.get("required"):
            description.append("必填")
        elif field.get("required") is False:
            description.append("可选")
        if field.get("nullable") is False:
            description.append("非空")
        elif field.get("nullable") is True:
            description.append("可空")
        detail = field.get("desc") or field.get("description") or field.get("purpose")
        if detail:
            description.append(as_text(detail))
        macro = r"\EaFieldFirst" if index == 0 else r"\EaField"
        lines.append(
            f"{macro}{{{name}}}{{{label}}}{{{tex_escape('；'.join(description))}}}"
        )
        nested = field.get("fields")
        if nested:
            lines.append(r"\EaKids{%")
            lines.extend(field_rows(nested, source).splitlines())
            lines.append("}")
    return "\n".join(lines)


def field_list(record: dict[str, Any], key: str, source: Path) -> list[Any]:
    value = required(record, key, source)
    if not isinstance(value, list):
        fail(f"{source}: {key} 必须是列表")
    return value


def optional_fields(record: dict[str, Any], key: str, source: Path) -> list[Any]:
    value = record.get(key)
    if value in (None, ""):
        return []
    if not isinstance(value, list):
        fail(f"{source}: {key} 必须是列表")
    return value


def rules_text(record: dict[str, Any], source: Path) -> str:
    value = record.get("rules")
    if value is None:
        return ""
    if isinstance(value, str):
        return as_text(value)
    if not isinstance(value, list) or not all(isinstance(item, str) for item in value):
        fail(f"{source}: rules 必须是字符串或字符串列表")
    return "；".join(as_text(item) for item in value if as_text(item))


def key_block(key: str, rows: str) -> list[str]:
    if not rows:
        return []
    return [f"  {key}={{", rows, "  },", f"  {key}rows={{1}},"]


def http_card(record: dict[str, Any], source: Path) -> str:
    for key in ("nesma", "section", "table", "caption"):
        required(record, key, source)
    method = as_text(required(record, "method", source)).upper()
    path = required(record, "path", source)
    controller = as_text(required(record, "controller", source))
    auth = required(record, "auth", source)
    request = optional_fields(record, "request", source)
    response_value = record.get("response")
    envelope = (
        record.get("envelope")
        or (
            "file-stream"
            if record.get("file")
            or record.get("response_kind") == "file"
            or (
                isinstance(response_value, dict)
                and {"kind", "name_pattern"} & set(response_value)
            )
            else None
        )
        or "ajax-result"
    )
    if envelope not in ENVELOPES:
        fail(f"{source}: envelope 必须是 {', '.join(sorted(ENVELOPES))}")
    if isinstance(response_value, list):
        response = response_value
    elif envelope == "file-stream" and response_value is not None:
        response = []
    else:
        response = optional_fields(record, "response", source)
    if not response and envelope != "file-stream" and record.get("data"):
        response = optional_fields(record, "data", source)

    rules = rules_text(record, source)
    rule = "；".join(part for part in (f"控制器：{controller}", rules) if part)
    request_key = "query" if method in {"GET", "HEAD"} else "body"
    response_rows = field_rows(response, source)
    parts = [
        rf"\EaHttpCard{{{tex_id(record['id'], source)}}}{{%",
        f"  title={{{tex_escape(record['title'])}}},",
        f"  method={{{tex_escape(method)}}},",
        f"  path={{{tex_escape(path)}}},",
        f"  nesma={{{tex_escape(record['nesma'])}}},",
        f"  table={{{tex_escape(record['table'])}}},",
        f"  auth={{{tex_escape(auth)}}},",
        f"  envelope={{{envelope}}},",
        *(key_block(request_key, field_rows(request, source)) if request else []),
        *(
            key_block("query", field_rows(optional_fields(record, "query", source), source))
            if not request and "query" in record
            else []
        ),
        *(
            key_block("body", field_rows(optional_fields(record, "body", source), source))
            if not request and "body" in record
            else []
        ),
        *(
            key_block("form", field_rows(optional_fields(record, "form", source), source))
            if not request and "form" in record
            else []
        ),
        *(
            key_block("pathparams", field_rows(optional_fields(record, "path_params", source), source))
            if not request and "path_params" in record
            else []
        ),
    ]
    if envelope == "file-stream":
        file_info = record.get("file") if isinstance(record.get("file"), dict) else {}
        if not file_info and isinstance(response_value, dict):
            file_info = response_value
        file_kind = (
            file_info.get("kind")
            or record.get("file_kind")
            or (record.get("file") if isinstance(record.get("file"), str) else None)
            or (response_value if isinstance(response_value, str) else None)
            or "file"
        )
        parts.append(f"  filekind={{{tex_escape(file_kind)}}},")
        if file_info.get("name_pattern"):
            parts.append(f"  filename={{{tex_escape(file_info['name_pattern'])}}},")
    elif record.get("response_java"):
        parts.append(f"  responsejava={{{tex_escape(record['response_java'])}}},")
    else:
        parts.append("  responsejava={成功响应业务字段},")
    parts.extend(
        [
            "  data={",
            response_rows,
            "  },",
            "  responserows={1},",
            f"  rule={{{tex_escape(rule)}}},",
            "}",
        ]
    )
    return "\n".join(parts)


def elf_card(record: dict[str, Any], source: Path) -> str:
    for key in ("nesma", "section", "table", "caption"):
        required(record, key, source)
    fields = field_list(record, "fields", source)
    if not fields:
        fail(f"{source}: fields 不得为空")
    source_name = as_text(required(record, "source", source))
    direction = as_text(required(record, "direction", source))
    rules = rules_text(record, source)
    rule = "；".join(
        part
        for part in (f"来源：{source_name}", f"方向：{direction}", rules)
        if part
    )
    return "\n".join(
        [
            rf"\EaElfCard{{{tex_id(record['id'], source)}}}{{%",
            f"  title={{{tex_escape(record['title'])}}},",
            f"  nesma={{{tex_escape(record['nesma'])}}},",
            f"  table={{{tex_escape(record['table'])}}},",
            "  data={",
            field_rows(fields, source),
            "  },",
            "  datarows={1},",
            f"  rule={{{tex_escape(rule)}}},",
            "}",
        ]
    )


def db_card(record: dict[str, Any], source: Path) -> str:
    columns = field_list(record, "columns", source)
    if not columns:
        fail(f"{source}: columns 不得为空")
    for column in columns:
        if not isinstance(column, dict):
            fail(f"{source}: columns 项必须是对象")
        for key in ("name", "type"):
            if key not in column:
                fail(f"{source}: columns 项缺少必需字段 {key}")
    parts = [rf"\EaDbCard{{{tex_id(record['id'], source)}}}{{%"]
    parts.append(f"  title={{{tex_escape(record['title'])}}},")
    if record.get("nesma"):
        parts.append(f"  nesma={{{tex_escape(record['nesma'])}}},")
    if record.get("table"):
        parts.append(f"  table={{{tex_escape(record['table'])}}},")
    parts.extend(
        [
            f"  tablename={{{tex_name(required(record, 'tableName', source))}}},",
            "  data={",
            field_rows(columns, source),
            "  },",
            "  datarows={1},",
        ]
    )
    description = record.get("purpose") or record.get("rule")
    if description:
        parts.append(f"  rule={{{tex_escape(description)}}},")
    parts.append("}")
    return "\n".join(parts)


def load_cards(root: Path) -> list[tuple[Path, dict[str, Any], str]]:
    paths = sorted(root.rglob("*.yaml"))
    if not paths:
        fail(f"{root}: 未找到 YAML")
    cards = []
    seen_ids = set()
    seen_names = set()
    for path in paths:
        try:
            record = yaml.safe_load(path.read_text(encoding="utf-8"))
        except yaml.YAMLError as error:
            fail(f"{path}: YAML 解析失败：{error}")
        if not isinstance(record, dict):
            fail(f"{path}: YAML 根节点必须是对象")
        for key in COMMON_FIELDS:
            required(record, key, path)
        if not isinstance(record["id"], str):
            fail(f"{path}: id 必须是字符串")
        if record["id"] in seen_ids:
            fail(f"{path}: id 重复：{record['id']}")
        seen_ids.add(record["id"])
        safe_name = re.sub(r"[^A-Za-z0-9_.-]+", "-", record["id"]).strip(".-")
        if not safe_name or safe_name in seen_names:
            fail(f"{path}: id 无法生成唯一的安全文件名：{record['id']}")
        seen_names.add(safe_name)
        kind = record["kind"]
        if kind == "http":
            card = http_card(record, path)
        elif kind == "elf":
            card = elf_card(record, path)
        elif kind == "db":
            card = db_card(record, path)
        else:
            fail(f"{path}: kind 必须是 http、elf 或 db")
        cards.append((path, record, card))
    return cards


def trim_white(image: Image.Image) -> Image.Image:
    rgb = image.convert("RGB")
    bbox = ImageChops.difference(rgb, Image.new("RGB", rgb.size, "white")).getbbox()
    if not bbox:
        return rgb
    left, top, right, bottom = bbox
    return rgb.crop(
        (
            max(0, left - 8),
            max(0, top - 8),
            min(rgb.width, right + 8),
            min(rgb.height, bottom + 8),
        )
    )


def render_png(
    card_name: str,
    card_tex: str,
    output: Path,
    style: Path,
    xelatex: str,
    pdftoppm: str,
) -> None:
    with tempfile.TemporaryDirectory(prefix="render-cards-") as temp:
        work = Path(temp)
        (work / "card.tex").write_text(
            "\n".join(
                [
                    r"\documentclass[UTF8,10pt,fontset=fandol]{ctexart}",
                    r"\usepackage[paperwidth=210mm,paperheight=1000mm,left=10mm,right=10mm,top=10mm,bottom=10mm]{geometry}",
                    r"\usepackage{api-card}",
                    r"\pagestyle{empty}",
                    r"\setlength{\parindent}{0pt}",
                    r"\begin{document}",
                    card_tex,
                    r"\end{document}",
                ]
            ),
            encoding="utf-8",
        )
        env = os.environ.copy()
        env["TEXINPUTS"] = f"{style.parent}:{env.get('TEXINPUTS', '')}"
        compiled = subprocess.run(
            [xelatex, "-interaction=nonstopmode", "-halt-on-error", "card.tex"],
            cwd=work,
            env=env,
            capture_output=True,
            text=True,
        )
        pdf = work / "card.pdf"
        if compiled.returncode != 0 or not pdf.exists():
            log = work / "card.log"
            detail = log.read_text(encoding="utf-8", errors="replace") if log.exists() else compiled.stdout
            fail(f"{card_name}: xelatex 失败：\n{detail[-3000:]}")
        converted = subprocess.run(
            [pdftoppm, "-r", "144", "-png", str(pdf), str(work / "page")],
            capture_output=True,
            text=True,
        )
        pages = sorted(work.glob("page-*.png"))
        if converted.returncode != 0 or not pages:
            fail(f"{card_name}: pdftoppm 失败：{converted.stderr[-1000:]}")
        images = [trim_white(Image.open(page)) for page in pages]
        stitched = Image.new(
            "RGB",
            (max(image.width for image in images), sum(image.height for image in images)),
            "white",
        )
        y = 0
        for image in images:
            stitched.paste(image, (0, y))
            y += image.height
        stitched.save(output, optimize=True)


def render(root: Path, style: Path) -> dict[str, Any]:
    root = root.resolve()
    style = style.resolve()
    if not root.is_dir():
        fail(f"接口表目录不存在：{root}")
    if not style.is_file():
        fail(f"样式文件不存在：{style}")
    tools = {name: shutil.which(name) for name in ("xelatex", "pdftoppm")}
    missing = [name for name, path in tools.items() if path is None]
    if missing:
        fail(f"缺少命令：{', '.join(missing)}")

    cards = load_cards(root)
    tex_dir = root / "tex" / "cards"
    image_dir = root / "images"
    tex_dir.mkdir(parents=True, exist_ok=True)
    image_dir.mkdir(parents=True, exist_ok=True)
    for stale in tex_dir.rglob("*.tex"):
        stale.unlink()
    for stale in image_dir.rglob("*.png"):
        stale.unlink()

    rendered = []
    for _, record, card_tex in cards:
        safe_name = re.sub(r"[^A-Za-z0-9_.-]+", "-", record["id"]).strip(".-")
        tex_path = tex_dir / f"{safe_name}.tex"
        image_path = image_dir / f"{safe_name}.png"
        tex_path.write_text(card_tex, encoding="utf-8")
        render_png(
            record["id"],
            card_tex,
            image_path,
            style,
            tools["xelatex"],
            tools["pdftoppm"],
        )
        rendered.append(
            {
                "id": record["id"],
                "kind": record["kind"],
                "tex": str(tex_path),
                "image": str(image_path),
            }
        )
    return {"root": str(root), "style": str(style), "cards": rendered}


def main() -> None:
    parser = argparse.ArgumentParser(
        description="把接口表 YAML 渲染为 TeX 卡片和 PNG 图片。"
    )
    parser.add_argument("root", type=Path, help="包含 YAML 的接口表目录")
    parser.add_argument(
        "style",
        nargs="?",
        type=Path,
        default=DEFAULT_STYLE,
        help="api-card.sty 路径；默认使用 Skill 自带样式",
    )
    args = parser.parse_args()
    try:
        result = render(args.root, args.style)
    except (OSError, ValueError) as error:
        parser.exit(1, f"FAIL: {error}\n")
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
