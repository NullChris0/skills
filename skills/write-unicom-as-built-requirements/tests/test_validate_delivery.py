import json
import shutil
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

from openpyxl import load_workbook


REPO = Path(__file__).parents[3]
SCRIPT = Path(__file__).parents[1] / "scripts" / "validate_delivery.py"
COUNT = REPO / "tests/fixtures/nesma-employee-assessment-v3/count/计数模型.json"
DELIVERY = Path("/home/nc/projects/work/hx-employee-assessment-delivery-v3-new-v2")


class DeliveryValidatorTest(unittest.TestCase):
    def run_validator(self, output):
        return subprocess.run(
            [sys.executable, str(SCRIPT), str(COUNT), str(output)],
            text=True,
            capture_output=True,
        )

    @unittest.skipUnless(DELIVERY.exists(), "golden delivery is not available")
    def test_golden_delivery_passes(self):
        result = self.run_validator(DELIVERY)
        self.assertEqual(result.returncode, 0, result.stderr)
        payload = json.loads(result.stdout)
        self.assertEqual(payload["status"], "PASS")
        self.assertEqual(payload["counts"]["ufp"], 267)
        self.assertEqual(payload["interfaces"]["yaml"], 73)

    @unittest.skipUnless(DELIVERY.exists(), "golden delivery is not available")
    def test_missing_mapping_row_is_rejected(self):
        temp = tempfile.TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        copied = Path(temp.name) / "delivery"
        shutil.copytree(DELIVERY, copied)
        path = copied / "映射表.xlsx"
        workbook = load_workbook(path)
        sheet = workbook["NESMA映射"]
        sheet.delete_rows(2)
        workbook.save(path)
        result = self.run_validator(copied)
        self.assertNotEqual(result.returncode, 0)
        self.assertIn("映射表编号", result.stderr)


if __name__ == "__main__":
    unittest.main()
