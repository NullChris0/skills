import json
import subprocess
import tempfile
import unittest
from pathlib import Path

from openpyxl import load_workbook


SCRIPT = Path(__file__).parents[1] / "scripts" / "generate_review_package.py"


def sample_model(method="overview"):
    return {
        "subject": "名片管理系统",
        "method": method,
        "count_type": "application",
        "user": "名片管理操作员",
        "scope": "当前明确需求",
        "boundary": "单机名片管理应用",
        "evidence": [{"id": "E1", "role": "规格", "source": "需求.md", "version": "v1", "location": "原始需求", "fact": "操作员维护并查询联系人", "claims": "D1,T1", "traceable": True}],
        "data_functions": [{"id": "D1", "name": "联系人", "type": "ILF", "change": "UNCHANGED", "ret": 1, "det": 9, "evidence_ids": ["E1"]}],
        "transactions": [{"id": "T1", "name": "录入联系人", "type": "EI", "change": "UNCHANGED", "det": 9, "ftr": 1, "evidence_ids": ["E1"]}],
        "non_counted": [{"requirement": "硬盘尽可能小", "reason": "非功能约束且无阈值"}],
        "assumptions": [],
    }


class GeneratorTest(unittest.TestCase):
    def generate(self, value):
        temp = tempfile.TemporaryDirectory()
        root = Path(temp.name)
        input_path = root / "input.json"
        input_path.write_text(json.dumps(value, ensure_ascii=False), encoding="utf-8")
        result = subprocess.run(["python", str(SCRIPT), str(input_path), str(root / "out")], text=True, capture_output=True)
        self.addCleanup(temp.cleanup)
        self.assertEqual(result.returncode, 0, result.stderr)
        files = sorted(p.name for p in (root / "out").iterdir())
        self.assertEqual(files, ["估算审阅报告.md", "功能点计数证据.xlsx", "规范化功能清单.md"])
        return root / "out"

    def test_ready_overview_uses_fixed_weights(self):
        out = self.generate(sample_model())
        wb = load_workbook(out / "功能点计数证据.xlsx", data_only=False)
        self.assertEqual(wb["数据功能"]["K2"].value, "=IF(J2,7,0)")
        self.assertEqual(wb["事务功能"]["L2"].value, "=IF(K2,4,0)")
        self.assertIn("待复核", (out / "估算审阅报告.md").read_text())

    def test_ready_detailed_applies_complexity_matrices(self):
        value = sample_model("detailed")
        value["data_functions"][0]["detail_evidence_ids"] = ["E1"]
        value["transactions"][0]["detail_evidence_ids"] = ["E1"]
        value["data_functions"][0].update(ret=2, det=30)
        value["transactions"][0].update(det=20, ftr=3)
        out = self.generate(value)
        wb = load_workbook(out / "功能点计数证据.xlsx", data_only=False)
        self.assertEqual(wb["数据功能"]["H2"].value, "中")
        self.assertEqual(wb["数据功能"]["I2"].value, 10)
        self.assertEqual(wb["事务功能"]["I2"].value, "高")
        self.assertEqual(wb["事务功能"]["J2"].value, 6)

    def test_missing_boundary_blocks_and_excludes_rows(self):
        value = sample_model()
        value["boundary"] = ""
        out = self.generate(value)
        wb = load_workbook(out / "功能点计数证据.xlsx", data_only=False)
        self.assertFalse(wb["数据功能"]["J2"].value)
        self.assertNotIn("正式功能点总数", (out / "估算审阅报告.md").read_text())

    def test_untraceable_derived_prd_blocks(self):
        value = sample_model()
        value["evidence"][0].update(role="意图", derived=True, traceable=False)
        out = self.generate(value)
        self.assertIn("派生材料不可追溯", (out / "估算审阅报告.md").read_text())

    def test_evidence_conflict_blocks_affected_rows(self):
        value = sample_model()
        value["evidence"][0]["conflict"] = "联系人由外部系统维护"
        out = self.generate(value)
        wb = load_workbook(out / "功能点计数证据.xlsx", data_only=False)
        self.assertFalse(wb["数据功能"]["J2"].value)
        self.assertIn("证据冲突", (out / "估算审阅报告.md").read_text())

    def test_enhancement_formulas_and_protection(self):
        value = sample_model("detailed")
        value.update(count_type="enhancement", baseline_ufp=100, baseline_evidence_ids=["E1"])
        value["transactions"][0]["detail_evidence_ids"] = ["E1"]
        value["data_functions"] = [
            {"id": "D1", "name": "新增", "type": "ILF", "change": "ADD", "ret": 1, "det": 9, "evidence_ids": ["E1"], "detail_evidence_ids": ["E1"]},
            {"id": "D2", "name": "删除", "type": "ILF", "change": "DEL", "ret": 1, "det": 9, "before_fp": 7, "evidence_ids": ["E1"], "detail_evidence_ids": ["E1"], "before_evidence_ids": ["E1"]},
            {"id": "D3", "name": "修改", "type": "ILF", "change": "CHG", "ret": 2, "det": 30, "before_fp": 7, "evidence_ids": ["E1"], "detail_evidence_ids": ["E1"], "before_evidence_ids": ["E1"]},
        ]
        out = self.generate(value)
        wb = load_workbook(out / "功能点计数证据.xlsx", data_only=False)
        summary = wb["汇总"]
        self.assertEqual(summary["B7"].value, "=B3+B4+B5")
        self.assertEqual(summary["B8"].value, "=B2+B3-B4+(B5-B6)")
        self.assertTrue(wb["数据功能"].protection.sheet)
        self.assertTrue(wb["数据功能"]["K2"].protection.locked)
        self.assertFalse(wb["数据功能"]["L2"].protection.locked)
        self.assertEqual(wb["数据功能"]["P2"].value, "E1")
        self.assertEqual(wb["数据功能"]["Q3"].value, "E1")
        self.assertEqual(summary["C2"].value, "E1")

    def test_workbook_persists_markdown_projection_state(self):
        out = self.generate(sample_model())
        wb = load_workbook(out / "功能点计数证据.xlsx", data_only=False)
        issue_rows = list(wb["假设与冲突"].values)
        self.assertIn(("非计数需求", "硬盘尽可能小：非功能约束且无阈值", "已排除", None), issue_rows)
        inventory = (out / "规范化功能清单.md").read_text()
        data_row = wb["数据功能"][2]
        self.assertIn(f"| {data_row[0].value} | {data_row[1].value} | {data_row[2].value} | {data_row[6].value} |", inventory)
        self.assertIn("硬盘尽可能小", inventory)
        report = (out / "估算审阅报告.md").read_text()
        self.assertIn(wb["计数说明"]["B7"].value, report)

    def test_rejects_unsupported_method(self):
        value = sample_model("estimated")
        temp = tempfile.TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        root = Path(temp.name)
        input_path = root / "input.json"
        input_path.write_text(json.dumps(value, ensure_ascii=False), encoding="utf-8")
        result = subprocess.run(["python", str(SCRIPT), str(input_path), str(root / "out")], text=True, capture_output=True)
        self.assertNotEqual(result.returncode, 0)
        self.assertIn("method", result.stderr)

    def test_blocked_package_retains_resolved_row_formulas(self):
        value = sample_model()
        value["evidence"].append({"id": "E2", "role": "规格", "source": "冲突.md", "version": "v1", "location": "1", "fact": "候选功能", "claims": "D2", "traceable": True, "conflict": "维护方不一致"})
        value["data_functions"].append({"id": "D2", "name": "候选", "type": "ILF", "change": "UNCHANGED", "ret": 1, "det": 2, "evidence_ids": ["E2"]})
        out = self.generate(value)
        wb = load_workbook(out / "功能点计数证据.xlsx", data_only=False)
        self.assertTrue(wb["数据功能"]["J2"].value)
        self.assertFalse(wb["数据功能"]["J3"].value)
        self.assertEqual(wb["汇总"]["B9"].value, "阻断")

    def test_rejects_invalid_function_type_and_counts(self):
        value = sample_model("detailed")
        value["data_functions"][0].update(type="TABLE", ret=0, detail_evidence_ids=["E1"])
        temp = tempfile.TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        root = Path(temp.name)
        input_path = root / "input.json"
        input_path.write_text(json.dumps(value, ensure_ascii=False), encoding="utf-8")
        result = subprocess.run(["python", str(SCRIPT), str(input_path), str(root / "out")], text=True, capture_output=True)
        self.assertNotEqual(result.returncode, 0)


if __name__ == "__main__":
    unittest.main()
