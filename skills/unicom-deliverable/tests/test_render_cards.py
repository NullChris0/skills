import importlib.util
import shutil
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

import yaml

RENDER = Path(__file__).parents[1] / "scripts/render_cards.py"
CARD = {
    "id": "db.example",
    "title": "示例物理表",
    "tableName": "td_example",
    "purpose": "保存示例记录。",
    "columns": [{"name": "id", "type": "bigint", "label": "主键"}, {"name": "name", "type": "varchar(64)", "label": "名称"}],
}


def load_module():
    spec = importlib.util.spec_from_file_location("render_cards", RENDER)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


class DbCardTest(unittest.TestCase):
    def test_物理表卡片输出表名字段且不带计数编号(self):
        tex = load_module().db_card({**CARD, "nesma": "DF-01"}, Path("x.yaml"))
        self.assertIn("td\\_", tex)
        self.assertIn("主键", tex)
        self.assertNotIn("DF-01", tex)

    def test_缺少列定义时报错(self):
        with self.assertRaises(ValueError):
            load_module().db_card({**CARD, "columns": []}, Path("x.yaml"))

    def test_非物理表卡片被拒绝(self):
        module = load_module()
        with tempfile.TemporaryDirectory() as temp:
            (Path(temp) / "a.yaml").write_text(yaml.safe_dump({**CARD, "kind": "http"}, allow_unicode=True), encoding="utf-8")
            with self.assertRaisesRegex(ValueError, "只支持物理表卡片"):
                module.load_cards(Path(temp))


@unittest.skipUnless(shutil.which("xelatex") and shutil.which("pdftoppm"), "card renderer unavailable")
class RenderCardsTest(unittest.TestCase):
    def test_物理表卡片渲染出图片(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp) / "物理表"
            root.mkdir()
            (root / "db.example.yaml").write_text(yaml.safe_dump(CARD, allow_unicode=True), encoding="utf-8")
            result = subprocess.run([sys.executable, str(RENDER), str(root)], capture_output=True, text=True)
            self.assertEqual(result.returncode, 0, result.stderr)
            self.assertTrue((root / "images/db.example.png").is_file())


if __name__ == "__main__":
    unittest.main()
