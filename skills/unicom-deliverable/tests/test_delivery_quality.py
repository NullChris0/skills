import copy
import json
import sys
import unittest
from pathlib import Path


REPO = Path(__file__).parents[3]
SCRIPTS = Path(__file__).parents[1] / "scripts"
sys.path.insert(0, str(SCRIPTS))

from delivery_quality import check_document_model, check_feature_granularity


FIXTURE = REPO / "tests/fixtures/nesma-employee-assessment-v3/delivery/delivery-model.json"


class DeliveryQualityTest(unittest.TestCase):
    def load(self):
        return json.loads(FIXTURE.read_text(encoding="utf-8"))

    def test_employee_golden_has_required_document_structure(self):
        delivery = self.load()
        result = check_document_model(
            delivery["model"],
            delivery["document"],
            delivery["mappings"],
        )
        self.assertEqual(result["sections"], 151)
        granularity = check_feature_granularity(
            delivery["features"],
            delivery["model"]["transactions"],
        )
        self.assertEqual(granularity["rows"], 37)

    def test_sparse_document_is_rejected(self):
        delivery = self.load()
        sparse = {
            "sections": [
                {
                    "heading": "交付范围与计数摘要",
                    "level": 2,
                    "paras": ["这是一个紧凑说明。"],
                }
            ]
        }
        with self.assertRaisesRegex(SystemExit, "1 项目概述"):
            check_document_model(delivery["model"], sparse, delivery["mappings"])

    def test_missing_data_model_section_body_is_rejected(self):
        delivery = copy.deepcopy(self.load())
        target = next(
            section
            for section in delivery["document"]["sections"]
            if section["heading"].startswith("7.1.1.3.1")
        )
        target["paras"] = []
        with self.assertRaisesRegex(SystemExit, "正文不足三段"):
            check_document_model(
                delivery["model"],
                delivery["document"],
                delivery["mappings"],
            )

    def test_grouped_feature_rows_are_rejected(self):
        delivery = self.load()
        with self.assertRaisesRegex(SystemExit, "一条事务功能一行"):
            check_feature_granularity(
                delivery["features"][:15],
                delivery["model"]["transactions"],
            )

    def test_nonleaf_section_body_is_rejected(self):
        delivery = self.load()
        target = next(
            section
            for section in delivery["document"]["sections"]
            if section["heading"].startswith("7.2.1 ")
        )
        target["paras"] = ["非叶子章节不应有说明。"]
        with self.assertRaisesRegex(SystemExit, "非叶子章节"):
            check_document_model(
                delivery["model"],
                delivery["document"],
                delivery["mappings"],
            )

    def test_table_reference_is_required(self):
        delivery = self.load()
        target = next(
            section
            for section in delivery["document"]["sections"]
            if section["heading"].startswith("7.2.1.1 ")
        )
        target["images"][0]["caption"] = "表7-1 查询考核模板列表"
        target["paras"] = [
            paragraph.replace("接口契约见表7-1", "接口契约见附表")
            for paragraph in target["paras"]
        ]
        with self.assertRaisesRegex(SystemExit, "缺少正文引用"):
            check_document_model(
                delivery["model"],
                delivery["document"],
                delivery["mappings"],
            )

    def test_fixed_document_purpose_is_required(self):
        delivery = self.load()
        target = next(
            section
            for section in delivery["document"]["sections"]
            if section["heading"] == "2 文档目的"
        )
        target["paras"] = ["自定义目的。"]
        with self.assertRaisesRegex(SystemExit, "2 文档目的"):
            check_document_model(
                delivery["model"],
                delivery["document"],
                delivery["mappings"],
            )

    def test_data_model_heading_rejects_df_id(self):
        delivery = self.load()
        target = next(
            section
            for section in delivery["document"]["sections"]
            if section["heading"].startswith("7.1.1 ")
        )
        target["heading"] = "7.1.1 DF-01 新闻栏目"
        with self.assertRaisesRegex(SystemExit, "不应写 DF-xx"):
            check_document_model(
                delivery["model"],
                delivery["document"],
                delivery["mappings"],
            )


if __name__ == "__main__":
    unittest.main()
