import shutil
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

import yaml


REPO = Path(__file__).parents[3]
RENDER = Path(__file__).parents[1] / "scripts/render_cards.py"


@unittest.skipUnless(shutil.which("xelatex") and shutil.which("pdftoppm"), "card renderer unavailable")
class RenderCardsTest(unittest.TestCase):
    def setUp(self):
        temp = tempfile.TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        self.root = Path(temp.name) / "接口表"
        self.root.mkdir()

    def write_card(self, record):
        (self.root / f"{record['id']}.yaml").write_text(
            yaml.safe_dump(record, allow_unicode=True, sort_keys=False),
            encoding="utf-8",
        )

    def test_uncounted_http_card_renders(self):
        self.write_card({
            "id": "status.read",
            "kind": "http",
            "title": "查询状态",
            "nesma": "uncounted",
            "section": "8.1",
            "table": "8-1",
            "caption": "查询状态接口",
            "method": "GET",
            "path": "/status",
            "controller": "StatusController#read",
            "auth": "admin",
            "response": [{"name": "status", "type": "string"}],
        })
        result = subprocess.run(
            [sys.executable, str(RENDER), str(self.root)],
            text=True,
            capture_output=True,
        )
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertIn("nesma={uncounted}", (self.root / "tex/cards/status.read.tex").read_text(encoding="utf-8"))
        self.assertTrue((self.root / "images/status.read.png").is_file())

    def test_db_without_nesma_renders(self):
        self.write_card({
            "id": "db.example",
            "kind": "db",
            "title": "示例表",
            "tableName": "example",
            "columns": [{"name": "id", "type": "bigint", "nullable": False}],
        })
        result = subprocess.run(
            [sys.executable, str(RENDER), str(self.root)],
            text=True,
            capture_output=True,
        )
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertTrue((self.root / "images/db.example.png").is_file())


class CallRuleTextTest(unittest.TestCase):
    def test_call_rule_prints_business_text_only(self):
        import importlib.util

        spec = importlib.util.spec_from_file_location("render_cards", RENDER)
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        tex = module.http_card(
            {
                "id": "item.list",
                "kind": "http",
                "title": "查询记录",
                "nesma": "TF-01",
                "section": "7.1",
                "table": "7-1",
                "caption": "查询记录",
                "method": "GET",
                "path": "/items",
                "controller": "ItemController#list",
                "auth": "管理端",
                "rules": ["按名称筛选。"],
                "response": [{"name": "id", "type": "string"}],
            },
            Path("item.list.yaml"),
        )
        self.assertIn("rule={按名称筛选。}", tex)
        self.assertIn("auth={管理端}", tex)
        self.assertNotIn("ItemController", tex)
        self.assertNotIn("控制器", tex)

    def test_data_model_cards_do_not_print_nesma(self):
        import importlib.util

        spec = importlib.util.spec_from_file_location("render_cards", RENDER)
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        db = module.db_card(
            {
                "id": "db.item",
                "kind": "db",
                "title": "记录物理表",
                "nesma": "DF-01",
                "section": "7.1.1",
                "table": "7.1-1",
                "caption": "表7.1-1 记录物理表",
                "tableName": "item",
                "columns": [{"name": "id", "type": "bigint"}],
            },
            Path("db.item.yaml"),
        )
        self.assertNotIn("DF-01", db)
        elf = module.elf_card(
            {
                "id": "master.item",
                "kind": "elf",
                "title": "外部记录",
                "nesma": "DF-02",
                "section": "7.1.2",
                "table": "7.1-2",
                "caption": "表7.1-2 外部记录",
                "source": "外部应用",
                "direction": "外部应用 -> 本应用",
                "fields": ["id"],
            },
            Path("master.item.yaml"),
        )
        self.assertNotIn("DF-02", elf)


if __name__ == "__main__":
    unittest.main()
