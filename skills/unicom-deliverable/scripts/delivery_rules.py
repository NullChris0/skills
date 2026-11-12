"""功能清单与需求说明书共用的读取和规则检查。"""

from __future__ import annotations

import re
from pathlib import Path

from openpyxl import load_workbook

HEADER = ["系统名称", "一级模块", "二级模块", "功能需求描述", "外部接口需求描述", "功能（新增/修改）"]
ROW_KEYS = ("module", "name", "desc", "external", "type")
CHAPTERS = [
    "项目概述", "文档目的", "文档范围", "术语和缩写解释", "参考文档", "范围定义",
    "系统交互关系", "功能需求", "接口需求", "界面需求", "非功能性需求", "其它需求",
]
# 交付件里不得出现的内容：计数编号与数字、工作量、接口路径、未处理的占位。
FORBIDDEN = re.compile(r"(?<![A-Za-z])[TD]F-?\d+|(?<![A-Za-z])\d+\s*FP\b|UFP|人月|/api/|【|TODO|待确认|待补")


def is_entity(name: str) -> bool:
    """名称以“信息”结尾的行是数据实体，其余是功能。"""
    return name.endswith("信息")


def forbidden_in(text: str) -> str | None:
    found = FORBIDDEN.search(text or "")
    return found.group(0) if found else None


def read_list(path: Path) -> tuple[str, list[dict[str, str]]]:
    """读功能清单，返回（系统名称, 行）。合并单元格按上一行的值向下填充。"""
    sheet = load_workbook(path).active
    header = [sheet.cell(1, column).value for column in range(1, 7)]
    if header != HEADER:
        raise ValueError(f"{path}: 表头应为 {HEADER}，实际为 {header}")
    system = str(sheet.cell(2, 1).value or "").strip()
    rows, module = [], None
    for index in range(2, sheet.max_row + 1):
        module = sheet.cell(index, 2).value or module
        values = [sheet.cell(index, column).value for column in (3, 4, 5, 6)]
        if values[0]:
            rows.append(dict(zip(ROW_KEYS, [str(module or "").strip()] + [str(v or "").strip() for v in values])))
    return system, rows


def check_rows(rows: list[dict[str, str]]) -> list[str]:
    """返回功能清单行的全部规则错误；空列表表示通过。"""
    errors, seen, closed, previous = [], set(), set(), None
    functions_started: set[str] = set()
    if not rows:
        return ["功能清单没有任何行"]
    for number, row in enumerate(rows, 1):
        where = f"第 {number} 行 {row.get('module')}/{row.get('name')}"
        missing = [key for key in ROW_KEYS if not str(row.get(key) or "").strip()]
        if missing:
            errors.append(f"{where}: 缺少 {missing}")
            continue
        if row["type"] not in ("新增", "修改"):
            errors.append(f"{where}: 功能列只能是 新增 或 修改")
        if row["module"] != previous:
            if row["module"] in closed:
                errors.append(f"{where}: 同一一级模块的行必须相邻")
            if previous is not None:
                closed.add(previous)
            previous = row["module"]
        if (row["module"], row["name"]) in seen:
            errors.append(f"{where}: 同一一级模块内二级模块重名")
        seen.add((row["module"], row["name"]))
        if is_entity(row["name"]):
            if row["module"] in functions_started:
                errors.append(f"{where}: 数据实体行必须排在本模块的功能行之前")
        else:
            functions_started.add(row["module"])
        for key in ("name", "desc", "external"):
            bad = forbidden_in(row[key])
            if bad:
                errors.append(f"{where}: {key} 含禁用内容 {bad!r}")
    return errors
