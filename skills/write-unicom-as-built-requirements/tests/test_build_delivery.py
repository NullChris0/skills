import json
import shutil
import subprocess
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path


REPO = Path(__file__).parents[3]
SCRIPTS = Path(__file__).parents[1] / "scripts"
BUILD = SCRIPTS / "build_delivery.py"
VALIDATE = SCRIPTS / "validate_delivery.py"
COUNT = REPO / "tests/fixtures/nesma-employee-assessment-v3/count/计数模型.json"
MODEL = REPO / "tests/fixtures/nesma-employee-assessment-v3/delivery/delivery-model.json"


@unittest.skipUnless(shutil.which("xelatex") and Path.home().joinpath(".dotnet/dotnet").exists(), "render toolchain unavailable")
class DeliveryBuildTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.temp = tempfile.TemporaryDirectory()
        cls.output = Path(cls.temp.name) / "delivery"
        result = subprocess.run(
            [sys.executable, str(BUILD), str(MODEL), str(cls.output)],
            text=True,
            capture_output=True,
        )
        if result.returncode:
            raise AssertionError(result.stderr or result.stdout)
        cls.validation = subprocess.run(
            [sys.executable, str(VALIDATE), str(COUNT), str(cls.output)],
            text=True,
            capture_output=True,
        )

    @classmethod
    def tearDownClass(cls):
        cls.temp.cleanup()

    def test_golden_delivery_builds_and_passes(self):
        self.assertEqual(self.validation.returncode, 0, self.validation.stderr)
        payload = json.loads(self.validation.stdout)
        self.assertEqual(payload["status"], "PASS")
        self.assertEqual(payload["counts"]["ufp"], 267)
        self.assertEqual(payload["interfaces"]["yaml"], 73)
        self.assertEqual(payload["docx"]["media"], 74)

    def copy_output(self):
        temp = tempfile.TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        output = Path(temp.name) / "delivery"
        shutil.copytree(self.output, output)
        return output

    def replace_docx_text(self, output, old, new):
        path = output / "需求说明书.docx"
        rewritten = path.with_suffix(".new")
        with zipfile.ZipFile(path) as source, zipfile.ZipFile(rewritten, "w") as target:
            for item in source.infolist():
                data = source.read(item.filename)
                if item.filename == "word/document.xml":
                    data = data.decode("utf-8").replace(old, new, 1).encode("utf-8")
                target.writestr(item, data)
        rewritten.replace(path)

    def test_missing_card_image_is_rejected(self):
        output = self.copy_output()
        next((output / "接口表/images").glob("*.png")).unlink()
        result = subprocess.run(
            [sys.executable, str(VALIDATE), str(COUNT), str(output)],
            text=True,
            capture_output=True,
        )
        self.assertNotEqual(result.returncode, 0)
        self.assertIn("YAML 与图片数量不一致", result.stderr)

    def test_old_ufp_is_rejected(self):
        output = self.copy_output()
        self.replace_docx_text(output, "267 UFP", "267 UFP 旧版 233 UFP")
        result = subprocess.run(
            [sys.executable, str(VALIDATE), str(COUNT), str(output)],
            text=True,
            capture_output=True,
        )
        self.assertNotEqual(result.returncode, 0)
        self.assertIn("UFP", result.stderr)

    def test_unknown_nesma_id_in_word_is_rejected(self):
        output = self.copy_output()
        self.replace_docx_text(output, "UFP", "UFP TF-99")
        result = subprocess.run(
            [sys.executable, str(VALIDATE), str(COUNT), str(output)],
            text=True,
            capture_output=True,
        )
        self.assertNotEqual(result.returncode, 0)
        self.assertIn("TF-99", result.stderr)


if __name__ == "__main__":
    unittest.main()
