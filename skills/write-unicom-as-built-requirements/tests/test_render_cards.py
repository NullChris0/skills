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


if __name__ == "__main__":
    unittest.main()
