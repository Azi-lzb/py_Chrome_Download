import unittest

import main


class BatchFillConfigTests(unittest.TestCase):
    def test_parse_batch_fill_config_groups_enabled_tables_and_rules(self):
        settings = {
            "task.max_items": "3",
            "task.auto_add_conditions": "yes",
            "task.sync_selected_tables": "yes",
            "task.on_extra_selected_tables": "uncheck",
            "wizard.next_clicks": "2",
        }
        tables = [
            {
                "enabled": "yes",
                "task_id": "T001",
                "report_group": "企业贷款",
                "report_name": "表1贷款利率时序表_省市用户",
                "multi_condition": "yes",
                "uniform_data_date": "2026-05-31",
            },
            {
                "enabled": "no",
                "task_id": "T002",
                "report_group": "跳过",
                "report_name": "跳过表",
            },
        ]
        rules = [
            {
                "enabled": "yes",
                "task_id": "T001",
                "condition_index": "1",
                "field_label": "查询方式",
                "field_type": "select",
                "field_value": "辖区金融机构",
                "sort_order": "10",
            },
            {
                "enabled": "yes",
                "task_id": "T001",
                "condition_index": "2",
                "field_label": "机构名称",
                "field_type": "text",
                "field_value": "中国工商银行惠州分行",
                "sort_order": "20",
            },
            {
                "enabled": "yes",
                "task_id": "UNKNOWN",
                "condition_index": "1",
                "field_label": "单位",
                "field_type": "select",
                "field_value": "万元",
            },
        ]

        config = main.parse_batch_fill_config(settings, tables, rules)

        self.assertEqual(config["max_items"], 3)
        self.assertTrue(config["auto_add_conditions"])
        self.assertTrue(config["sync_selected_tables"])
        self.assertEqual(config["on_extra_selected_tables"], "uncheck")
        self.assertEqual(len(config["tasks"]), 1)
        task = config["tasks"][0]
        self.assertEqual(task["task_id"], "T001")
        self.assertEqual(task["report_group"], "企业贷款")
        self.assertEqual(task["report_name"], "表1贷款利率时序表_省市用户")
        self.assertEqual(task["uniform_data_dates"], ["2026-05-31"])
        self.assertEqual([c["condition_index"] for c in task["conditions"]], [1, 2])
        self.assertEqual(task["conditions"][0]["field_label"], "查询方式")
        self.assertEqual(task["conditions"][0]["field_value"], "辖区金融机构")
        self.assertEqual(task["conditions"][1]["field_label"], "机构名称")
        self.assertEqual(task["conditions"][1]["field_value"], "中国工商银行惠州分行")

    def test_parse_batch_fill_config_rejects_missing_condition_index(self):
        settings = {}
        tables = [{"enabled": "yes", "task_id": "T001", "report_group": "企业贷款", "report_name": "表1"}]
        rules = [{"enabled": "yes", "task_id": "T001", "condition_index": "", "field_label": "查询方式", "field_value": "辖区金融机构"}]

        with self.assertRaises(ValueError):
            main.parse_batch_fill_config(settings, tables, rules)


if __name__ == "__main__":
    unittest.main()
