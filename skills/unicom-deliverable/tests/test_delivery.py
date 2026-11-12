import copy
import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

from docx import Document
from openpyxl import load_workbook

SCRIPTS = Path(__file__).parents[1] / "scripts"
sys.path.insert(0, str(SCRIPTS))

import build_feature_list  # noqa: E402
import build_requirement_doc  # noqa: E402
import build_structure_diagram  # noqa: E402
import delivery_rules  # noqa: E402
import validate_delivery  # noqa: E402


def row(module, name, kind="新增", desc="输入名称，显示匹配的记录列表。", external="不涉及"):
    return {"module": module, "name": name, "desc": desc, "external": external, "type": kind}


ROWS = [
    row("栏目管理", "栏目信息", desc="栏目信息包括栏目名称、启用状态；其中栏目名称为输入项。"),
    row("栏目管理", "栏目新增"),
    row("栏目管理", "栏目修改", kind="修改"),
    row("文章管理", "文章列表查询", external="读取主系统人员信息。"),
]
SPEC = {
    "title": "示例系统", "date": "2026年1月", "system_name": "示例系统",
    "background": ["示例系统用于内容发布。", "具体需求如下：", "1、维护栏目。"],
    "structure_intro": "示例系统包括栏目管理和文章管理两个模块。",
    "structure_table": [["栏目管理", "栏目维护", "管理员"], ["文章管理", "文章查询", "管理员"]],
    "history": ["无。"], "purpose": ["明确业务需求。"], "scope": ["本文覆盖两个模块。"],
    "terms": [["栏目", "文章的分类"]], "references": ["《示例规范》"], "boundary": ["登录由主系统提供。"],
    "interactions": ["涉及系统间交互：", "1、对接主系统，读取人员信息。"],
    "interfaces": ["主系统人员信息查询 - 复用"], "ui_intro": "管理端包括栏目和文章两个页面。",
    "nonfunctional": ["性能：页面加载 ≤2s。"],
    "modules": [
        {"name": "栏目管理", "functions": [
            {"name": "栏目信息", "desc": "新增栏目信息，涉及属性包含栏目名称、启用状态等。"},
            {"name": "栏目新增", "desc": "新增栏目新增服务，管理员输入栏目名称后保存。", "input": "栏目名称", "output": "保存结果"},
            {"name": "栏目修改", "type": "修改", "desc": "新增栏目修改服务，管理员修改栏目名称后保存。", "input": "栏目标识、栏目名称", "output": "保存结果"}]},
        {"name": "文章管理", "functions": [
            {"name": "文章列表查询", "desc": "新增文章列表查询服务，管理员按标题查询文章。", "input": "标题", "output": "文章标题、作者"}]},
    ],
}


class RowRulesTest(unittest.TestCase):
    def test_合规的清单行没有错误(self):
        self.assertEqual(delivery_rules.check_rows(ROWS), [])

    def test_同一一级模块的行不相邻时报错(self):
        rows = [ROWS[1], ROWS[3], row("栏目管理", "栏目删除")]
        self.assertTrue(any("必须相邻" in e for e in delivery_rules.check_rows(rows)))

    def test_数据实体排在功能之后时报错(self):
        self.assertTrue(any("数据实体行必须排在" in e for e in delivery_rules.check_rows([ROWS[1], ROWS[0]])))

    def test_同模块二级模块重名时报错(self):
        self.assertTrue(any("重名" in e for e in delivery_rules.check_rows([ROWS[1], ROWS[1]])))

    def test_描述里出现计数编号或接口路径或占位时报错(self):
        for text in ("对应 TF-01 的功能。", "合计 4 FP。", "调用 /api/news/list 查询。", "字段【待确认】。", "约 3 人月。"):
            with self.subTest(text=text):
                self.assertTrue(any("禁用内容" in e for e in delivery_rules.check_rows([row("甲", "乙查询", desc=text)])))

    def test_功能列只能是新增或修改(self):
        self.assertTrue(any("只能是" in e for e in delivery_rules.check_rows([row("甲", "乙查询", kind="删除")])))


class DeliveryRoundTripTest(unittest.TestCase):
    def setUp(self):
        temp = tempfile.TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        self.dir = Path(temp.name)
        self.xlsx = self.dir / "功能清单.xlsx"
        self.docx = self.dir / "需求说明书.docx"
        build_feature_list.build(ROWS, "示例系统", self.xlsx)

    def test_清单写出六列且合并单元格读回后与行数据一致(self):
        sheet = load_workbook(self.xlsx).active
        self.assertEqual([c.value for c in sheet[1]], delivery_rules.HEADER)
        self.assertEqual(sheet.max_column, 6)
        self.assertIn("B2:B4", {str(r) for r in sheet.merged_cells.ranges})
        self.assertEqual(delivery_rules.read_list(self.xlsx), ("示例系统", ROWS))

    def test_说明书是十二章且第八章与清单逐条一致(self):
        stats = build_requirement_doc.build(SPEC, self.xlsx, self.docx)
        self.assertEqual(stats["functions"], 4)
        items = validate_delivery.read_doc(self.docx)
        self.assertEqual(validate_delivery.check_doc(items, "示例系统", ROWS), [])
        texts = [text for _, text in items]
        self.assertIn("8.1.1.3 栏目修改", texts)
        self.assertIn("修改类型：□新增 ☑修改 □删除", texts)
        self.assertIn("1.4 本项目预算工作量", texts)
        self.assertNotIn("{{", "".join(texts))
        self.assertEqual(len(Document(self.docx).inline_shapes), 0)

    def test_数据实体只写功能描述而功能带输入输出(self):
        build_requirement_doc.build(SPEC, self.xlsx, self.docx)
        texts = [text for _, text in validate_delivery.read_doc(self.docx)]
        entity = texts.index("8.1.1.1 栏目信息")
        self.assertEqual([t[:4] for t in texts[entity + 1:entity + 3]], ["修改类型", "功能描述"])
        self.assertEqual(texts[entity + 3], "8.1.1.2 栏目新增")
        self.assertEqual(texts[entity + 6:entity + 8], ["输入项：栏目名称", "输出项：保存结果"])

    def test_第八章与清单不一致时拒绝生成(self):
        spec = copy.deepcopy(SPEC)
        spec["modules"][0]["functions"].reverse()
        with self.assertRaisesRegex(ValueError, "与功能清单第 1 条不一致"):
            build_requirement_doc.build(spec, self.xlsx, self.docx)

    def test_功能缺输出项或实体带输入项时拒绝生成(self):
        spec = copy.deepcopy(SPEC)
        del spec["modules"][0]["functions"][1]["output"]
        spec["modules"][0]["functions"][0]["input"] = "栏目名称"
        with self.assertRaises(ValueError) as caught:
            build_requirement_doc.build(spec, self.xlsx, self.docx)
        self.assertIn("功能必须同时有 input 和 output", str(caught.exception))
        self.assertIn("数据实体不写 input/output", str(caught.exception))

    def test_正文出现人月或功能点数时拒绝生成(self):
        spec = copy.deepcopy(SPEC)
        spec["history"] = ["本项目分配 3.3 人月。"]
        with self.assertRaisesRegex(ValueError, "禁用内容"):
            build_requirement_doc.build(spec, self.xlsx, self.docx)

    def test_清单改动后旧说明书校验失败(self):
        build_requirement_doc.build(SPEC, self.xlsx, self.docx)
        build_feature_list.build(ROWS + [row("文章管理", "文章导出")], "示例系统", self.xlsx)
        result = subprocess.run([sys.executable, str(SCRIPTS / "validate_delivery.py"), str(self.xlsx), str(self.docx)],
                                capture_output=True, text=True)
        self.assertEqual(result.returncode, 1)
        self.assertIn("文章导出", result.stderr)

    def test_带结构图和附录时图片进入文档(self):
        from PIL import Image
        image = self.dir / "图.png"
        Image.new("RGB", (400, 200), "white").save(image)
        spec = {**SPEC, "structure_figure": {"image": str(image)}, "figures": [{"image": str(image), "caption": "栏目管理"}],
                "appendix": [{"image": str(image), "caption": "表1 栏目物理表"}]}
        build_requirement_doc.build(spec, self.xlsx, self.docx)
        texts = [text for _, text in validate_delivery.read_doc(self.docx)]
        self.assertIn("系统总体功能结构图", texts)
        self.assertIn("附录 A 物理表清单", texts)
        self.assertEqual(texts[-1], "表1 栏目物理表")
        self.assertEqual(Document(self.docx).element.body.xml.count("<a:blip "), 3)


class CoverageTest(unittest.TestCase):
    MODEL = {"data_functions": [{"id": "DF-01", "name": "栏目"}], "transactions": [{"id": "TF-01", "name": "新增栏目"}]}

    def test_计数项全部映射到清单行时通过(self):
        mapping = {"DF-01": ["栏目管理/栏目信息"], "TF-01": ["栏目管理/栏目新增"]}
        self.assertEqual(validate_delivery.check_coverage(self.MODEL, mapping, ROWS), [])

    def test_计数项未映射或映射到不存在的行时报错(self):
        errors = validate_delivery.check_coverage(self.MODEL, {"DF-01": ["栏目管理/不存在"]}, ROWS)
        self.assertTrue(any("TF-01" in e and "没有映射" in e for e in errors))
        self.assertTrue(any("不在功能清单中" in e for e in errors))


class StructureDiagramTest(unittest.TestCase):
    SPEC = {"title": "示例系统", "left": [{"title": "主系统", "items": ["统一登录"]}],
            "rows": [[{"title": "栏目管理", "cols": 2, "items": ["栏目增删改", ["对象存储", "exist"]]}]]}

    def test_生成的源文件含标题模块方块和图例(self):
        xml, _, _ = build_structure_diagram.build(self.SPEC)
        for text in ("示例系统", "栏目管理", "栏目增删改", "本次新增", "已有能力"):
            self.assertIn(text, xml)
        self.assertEqual(xml.count(build_structure_diagram.EXISTING), 3)

    def test_方块文字过长时报错(self):
        spec = json.loads(json.dumps(self.SPEC))
        spec["rows"][0][0]["cols"] = 8
        spec["rows"][0][0]["items"] = ["字" * 30]
        with self.assertRaises(ValueError):
            build_structure_diagram.build(spec)


if __name__ == "__main__":
    unittest.main()
