#!/usr/bin/env python3
"""把功能清单行数据写成联通功能清单 xlsx。"""

from __future__ import annotations

import argparse
import json
from copy import copy
from pathlib import Path

from openpyxl import load_workbook

from delivery_rules import ROW_KEYS, check_rows

TEMPLATE = Path(__file__).resolve().parent.parent / "assets" / "联通功能清单模板.xlsx"


def build(rows: list[dict[str, str]], system: str, output: Path, template: Path = TEMPLATE) -> None:
    errors = check_rows(rows)
    if errors:
        raise ValueError("\n".join(errors))
    book = load_workbook(template)
    sheet = book.active
    styles = [copy(cell._style) for cell in sheet[2][:6]]
    for merged in list(sheet.merged_cells.ranges):
        sheet.unmerge_cells(str(merged))
    if sheet.max_row > 2:
        sheet.delete_rows(3, sheet.max_row - 2)
    start, last = 2, len(rows) + 1
    for index, row in enumerate(rows, 2):
        for column in range(1, 7):
            cell = sheet.cell(index, column)
            cell.value, cell._style = None, copy(styles[column - 1])
        for column, key in zip(range(3, 7), ROW_KEYS[1:]):
            sheet.cell(index, column).value = row[key]
        # 行高按字数估算：功能描述列一行约 60 个汉字。
        lines = max(len(row["desc"]), len(row["external"]) * 2) // 60 + 1
        sheet.row_dimensions[index].height = max(24, 13 * lines + 8)
        if index == last or rows[index - 1]["module"] != row["module"]:
            sheet.cell(start, 2).value = row["module"]
            if index > start:
                sheet.merge_cells(start_row=start, start_column=2, end_row=index, end_column=2)
            start = index + 1
    sheet.cell(2, 1).value = system
    if last > 2:
        sheet.merge_cells(start_row=2, start_column=1, end_row=last, end_column=1)
    book.save(output)


def main() -> None:
    parser = argparse.ArgumentParser(description="把 rows.json 写成联通功能清单。")
    parser.add_argument("rows", type=Path, help="行数据 JSON：对象数组，键为 module/name/desc/external/type")
    parser.add_argument("system", help="系统名称，写入 A 列")
    parser.add_argument("output", type=Path)
    parser.add_argument("--template", type=Path, default=TEMPLATE)
    args = parser.parse_args()
    rows = json.loads(args.rows.read_text(encoding="utf-8"))
    try:
        build(rows, args.system, args.output, args.template)
    except ValueError as error:
        parser.exit(1, f"FAIL:\n{error}\n")
    modules: dict[str, int] = {}
    for row in rows:
        modules[row["module"]] = modules.get(row["module"], 0) + 1
    print(f"OK {args.output}: {len(rows)} 行")
    for name, count in modules.items():
        print(f"  {name}: {count}")


if __name__ == "__main__":
    main()
