#!/usr/bin/env python3
"""生成“系统总体功能结构图”的 draw.io 源文件（分块式，与客户范例同款）。

导出图片：drawio -x -f png -s 2 -o 结构图.png 结构图.drawio
"""

from __future__ import annotations

import argparse
import json
import math
from pathlib import Path
from xml.sax.saxutils import quoteattr

NEW, EXISTING = "#D9D9D9", "#FFD6D6"   # 客户范例图例：灰=本次新增，粉=已有能力
WIDTH, GAP, PAD, BLOCK_H, TITLE_H, LEFT_W, MARGIN = 1080, 14, 8, 32, 30, 150, 20
BLOCK = "rounded=1;whiteSpace=wrap;html=1;arcSize=12;strokeColor=none;fontSize=12;fillColor={};"
GROUP = ("rounded=1;whiteSpace=wrap;html=1;arcSize=5;fillColor=#FFFFFF;strokeColor=#BDBDBD;"
         "verticalAlign=top;fontStyle=1;fontSize=14;spacingTop=2;")


def build(spec: dict) -> tuple[str, int, int]:
    cells, counter = [], [1]

    def cell(value, style, x, y, w, h):
        counter[0] += 1
        cells.append(f'<mxCell id="{counter[0]}" value={quoteattr(value)} style="{style}" vertex="1" parent="1">'
                     f'<mxGeometry x="{x:.0f}" y="{y:.0f}" width="{w:.0f}" height="{h:.0f}" as="geometry"/></mxCell>')

    def block(label, existing, x, y, w):
        if len(label) > int(w // 13) * 2:      # 12 号字两行放不下
            raise ValueError(f"方块文字过长：{label}（宽 {w:.0f}）")
        cell(label, BLOCK.format(EXISTING if existing else NEW), x, y, w, BLOCK_H)

    x0, top = MARGIN + LEFT_W + 30, 64
    y = top
    for row in spec["rows"]:
        columns = sum(group["cols"] for group in row)
        block_w = (WIDTH - (len(row) - 1) * GAP - sum((group["cols"] + 1) * PAD for group in row)) / columns
        if block_w < 96:
            raise ValueError(f"一行里列数太多：{[group['title'] for group in row]}")
        height = TITLE_H + max(math.ceil(len(group["items"]) / group["cols"]) for group in row) * (BLOCK_H + PAD) + PAD
        x = x0
        for group in row:
            group_w = group["cols"] * block_w + (group["cols"] + 1) * PAD
            cell(group["title"], GROUP, x, y, group_w, height)
            for index, item in enumerate(group["items"]):
                label, existing = (item, False) if isinstance(item, str) else (item[0], item[1] == "exist")
                block(label, existing, x + PAD + (index % group["cols"]) * (block_w + PAD),
                      y + TITLE_H + (index // group["cols"]) * (BLOCK_H + PAD), block_w)
            x += group_w + GAP
        y += height + GAP
    left_height = 10 + sum(46 + len(group["items"]) * (BLOCK_H + 10) + 14 for group in spec["left"])
    bottom = max(y - GAP, top + left_height)   # 右侧模块少时，画框跟着左侧外部系统栏撑高

    cell("", "rounded=1;whiteSpace=wrap;html=1;arcSize=8;fillColor=#F2F2F2;strokeColor=#BDBDBD;", MARGIN, top, LEFT_W, bottom - top)
    left_y = top + 10
    for group in spec["left"]:
        cell(group["title"], "text;html=1;align=center;verticalAlign=middle;whiteSpace=wrap;fontStyle=1;fontSize=13;",
             MARGIN + 4, left_y, LEFT_W - 8, 40)
        left_y += 46
        for item in group["items"]:
            block(item, True, MARGIN + 12, left_y, LEFT_W - 24)
            left_y += BLOCK_H + 10
        left_y += 14
    counter[0] += 1
    cells.append(f'<mxCell id="{counter[0]}" style="endArrow=classic;startArrow=classic;html=1;strokeColor=#7F7F7F;" edge="1" parent="1">'
                 f'<mxGeometry relative="1" as="geometry"><mxPoint x="{MARGIN + LEFT_W + 2}" y="{top + 24}" as="sourcePoint"/>'
                 f'<mxPoint x="{x0 - 2}" y="{top + 24}" as="targetPoint"/></mxGeometry></mxCell>')
    legend_x = x0 + WIDTH - 2 * 120 - 12
    for index, (label, color) in enumerate((("本次新增", NEW), ("已有能力", EXISTING))):
        cell(label, BLOCK.format(color), legend_x + index * 132, bottom + 16, 120, 28)
    total_w, total_h = x0 + WIDTH + MARGIN, bottom + 16 + 28 + MARGIN
    frame = (f'<mxCell id="frame" value="" style="rounded=0;whiteSpace=wrap;html=1;fillColor=#FAFAFA;strokeColor=#9E9E9E;" vertex="1" parent="1">'
             f'<mxGeometry x="0" y="0" width="{total_w}" height="{total_h}" as="geometry"/></mxCell>'
             f'<mxCell id="title" value={quoteattr(spec["title"])} style="text;html=1;align=center;verticalAlign=middle;fontStyle=1;fontSize=24;" vertex="1" parent="1">'
             f'<mxGeometry x="0" y="10" width="{total_w}" height="44" as="geometry"/></mxCell>')
    xml = (f'<mxfile host="app.diagrams.net"><diagram name="系统总体功能结构"><mxGraphModel dx="{total_w}" dy="{total_h}" grid="0" page="0" math="0" shadow="0">'
           f'<root><mxCell id="0"/><mxCell id="1" parent="0"/>{frame}{"".join(cells)}</root></mxGraphModel></diagram></mxfile>')
    return xml, total_w, total_h


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("spec", type=Path)
    parser.add_argument("output", type=Path)
    args = parser.parse_args()
    spec = json.loads(args.spec.read_text(encoding="utf-8"))
    try:
        xml, width, height = build(spec)
    except ValueError as error:
        parser.exit(1, f"FAIL: {error}\n")
    args.output.write_text(xml, encoding="utf-8")
    blocks = sum(len(group["items"]) for row in spec["rows"] for group in row)
    print(f"OK {args.output}: {width}x{height}, 模块 {sum(len(row) for row in spec['rows'])}, 方块 {blocks}")


if __name__ == "__main__":
    main()
