#!/usr/bin/env python3
"""Fill TOC page hints from a rendered delivery PDF."""

import argparse
import json
import re
import shutil
import subprocess
from pathlib import Path


def require_tool(name):
    path = shutil.which(name)
    if not path:
        raise SystemExit(f"FAIL: 未找到 {name}")
    return path


def compact(value):
    return re.sub(r"\s+", "", str(value or ""))


def page_text(pdf, page):
    result = subprocess.run(
        [
            require_tool("pdftotext"),
            "-f",
            str(page),
            "-l",
            str(page),
            "-layout",
            str(pdf),
            "-",
        ],
        check=True,
        stdout=subprocess.PIPE,
        stderr=subprocess.DEVNULL,
        text=True,
    )
    return compact(result.stdout)


def page_count(pdf):
    result = subprocess.run(
        [require_tool("pdfinfo"), str(pdf)],
        check=True,
        stdout=subprocess.PIPE,
        stderr=subprocess.DEVNULL,
        text=True,
    )
    match = re.search(r"^Pages:\s+(\d+)", result.stdout, re.MULTILINE)
    if not match:
        raise SystemExit("FAIL: 无法读取 PDF 页数")
    return int(match.group(1))


def body_start(pdf, sections):
    for section in sections:
        for paragraph in section.get("paras", []):
            needle = compact(paragraph)
            if len(needle) >= 12:
                for page in range(1, min(page_count(pdf), 20) + 1):
                    if needle[:24] in page_text(pdf, page):
                        return page
    for section in sections:
        if int(section.get("level", 0)) == 1:
            needle = compact(section.get("heading"))
            seen = 0
            for page in range(1, page_count(pdf) + 1):
                if needle in page_text(pdf, page):
                    seen += 1
                    if seen == 2:
                        return page
    raise SystemExit("FAIL: 无法定位正文起始页")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("delivery_model", type=Path)
    parser.add_argument("pdf", type=Path)
    args = parser.parse_args()
    payload = json.loads(args.delivery_model.read_text(encoding="utf-8"))
    sections = payload["document"]["sections"]
    start = body_start(args.pdf, sections)
    total = page_count(args.pdf)
    current = start
    for section in sections:
        level = int(section.get("level", 0))
        if level < 1 or level > 4:
            continue
        heading = compact(section.get("heading"))
        found = None
        for page in range(current, total + 1):
            if heading in page_text(args.pdf, page):
                found = page
                break
        if found is None:
            raise SystemExit(f"FAIL: 无法定位目录标题：{section.get('heading')}")
        section["page_hint"] = str(found)
        current = found
    args.delivery_model.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    print(f"Updated {args.delivery_model}")


if __name__ == "__main__":
    main()
