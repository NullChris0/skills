#!/usr/bin/env python3
from __future__ import annotations

import json
import re
import subprocess
import sys
from pathlib import Path


MODEL_PATH = Path("/tmp/hx-delivery-work/delivery-model.json")
PDF_PATH = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(
    "/tmp/hx-old-final/需求说明书.pdf"
)


def page_text(page: int) -> str:
    return subprocess.run(
        [
            "pdftotext",
            "-f",
            str(page),
            "-l",
            str(page),
            "-layout",
            str(PDF_PATH),
            "-",
        ],
        check=True,
        stdout=subprocess.PIPE,
        stderr=subprocess.DEVNULL,
        text=True,
    ).stdout


def compact(text: str) -> str:
    return re.sub(r"\s+", "", text)


def main() -> None:
    payload = json.loads(MODEL_PATH.read_text(encoding="utf-8"))
    sections = payload["document"]["sections"]
    page_cache: dict[int, str] = {}

    body_start = None
    for page in range(2, 20):
        text = page_text(page)
        page_cache[page] = compact(text)
        if "M域信创改造" in compact(text):
            body_start = page
            break
    if body_start is None:
        raise RuntimeError("Cannot find body start")

    current_page = body_start
    max_page = 120
    for section in sections:
        level = int(section["level"])
        if level > 4:
            continue
        needle = compact(section["heading"])
        found = None
        for page in range(current_page, max_page + 1):
            if page not in page_cache:
                try:
                    page_cache[page] = compact(page_text(page))
                except subprocess.CalledProcessError:
                    break
            if needle in page_cache[page]:
                found = page
                break
        if found is None:
            raise RuntimeError(f"Cannot locate heading: {section['heading']}")
        section["page_hint"] = str(found)
        current_page = found

    MODEL_PATH.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    print(f"Updated TOC pages from PDF; body starts at {body_start}")


if __name__ == "__main__":
    main()
