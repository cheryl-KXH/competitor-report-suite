from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from service import dingtalk_docs, dingtalk_table
from service.jobs import _task_docs_folder_id, run_generate_delivery_tables


class DeliveryJobTests(unittest.TestCase):
    def test_all_platform_results_use_attachment_cell_format(self) -> None:
        configs = {
            "dingtalk": {"baseId": "base", "tableId": "table"},
            "field_mapping": {
                "fields": {
                    "meituanData": {"fieldId": "mt"},
                    "elemeData": {"fieldId": "elm"},
                    "jdData": {"fieldId": "jd"},
                }
            },
        }
        with patch("service.dingtalk_table._update_cells") as update_cells:
            dingtalk_table.mark_links(
                configs,
                "record",
                {
                    "meituanData": ("美团.xlsx", "https://example/mt"),
                    "elemeData": ("饿了么.xlsx", "https://example/elm"),
                    "jdData": ("京东.xlsx", "https://example/jd"),
                },
            )

        cells = update_cells.call_args.args[2]
        self.assertEqual(cells["meituanData"], [{"url": "https://example/mt", "name": "美团.xlsx"}])
        self.assertEqual(cells["elemeData"], [{"url": "https://example/elm", "name": "饿了么.xlsx"}])
        self.assertEqual(cells["jdData"], [{"url": "https://example/jd", "name": "京东.xlsx"}])

    def test_attachment_name_is_supported_when_downloading_linked_xlsx(self) -> None:
        value = [{"name": "产品清单.xlsx", "url": "https://alidocs.dingtalk.com/i/nodes/menu-node"}]
        with tempfile.TemporaryDirectory() as tmp:
            expected = Path(tmp) / "产品清单.xlsx"
            with patch("service.dingtalk_docs.download_file", return_value=expected) as download_file:
                actual = dingtalk_docs.download_linked_file({}, value, Path(tmp))

        self.assertEqual(actual, expected)
        self.assertEqual(download_file.call_args.args[1]["nodeId"], "menu-node")
        self.assertEqual(download_file.call_args.args[1]["name"], "产品清单.xlsx")

    def test_attachment_name_is_supported_as_delivery_input(self) -> None:
        value = [{"name": "外卖数据.xlsx", "url": "https://alidocs.dingtalk.com/i/nodes/raw-node"}]
        with tempfile.TemporaryDirectory() as tmp:
            output_dir = Path(tmp)
            with patch("service.dingtalk_docs.download_file", return_value=output_dir / "外卖数据.xlsx") as download_file:
                actual = dingtalk_docs.download_linked_folder({}, value, output_dir)

        self.assertEqual(actual, output_dir)
        self.assertEqual(download_file.call_args.args[1]["nodeId"], "raw-node")

    def test_target_folder_falls_back_from_delivery_data_to_product_menu(self) -> None:
        record = dingtalk_table.TaskRecord(
            "record",
            {
                "deliveryData": [{"name": "外卖数据.xlsx", "url": "https://example/raw"}],
                "productMenu": [{"name": "产品清单.xlsx", "url": "https://example/menu"}],
            },
        )
        with patch("service.jobs.dingtalk_docs.folder_id_for_linked_node", side_effect=[None, "shared-folder"]) as finder:
            folder_id = _task_docs_folder_id({"dingtalk_docs": {}}, record)

        self.assertEqual(folder_id, "shared-folder")
        self.assertEqual(finder.call_count, 2)

    def test_delivery_job_uploads_all_results_to_one_folder_then_updates_row(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            annotation_path = root / "产品清单.xlsx"
            annotation_path.write_bytes(b"annotation")
            outputs = {
                "meituanData": root / "美团外卖数据.xlsx",
                "elemeData": root / "饿了么外卖数据.xlsx",
                "jdData": root / "京东外卖数据.xlsx",
            }
            for path in outputs.values():
                path.write_bytes(b"result")
            configs = {
                "dingtalk": {"baseId": "base", "tableId": "table"},
                "dingtalk_docs": {},
                "field_mapping": {"fields": {}},
                "report_rules": {"outputDirectory": str(root)},
            }
            record = dingtalk_table.TaskRecord(
                "record",
                {
                    "deliveryData": [{"name": "外卖数据.xlsx", "url": "https://example/raw"}],
                    "productMenu": [{"name": "产品清单.xlsx", "url": "https://example/menu"}],
                },
            )

            with (
                patch("service.jobs.load_configs", return_value=configs),
                patch("service.jobs.dingtalk_table.mark_status"),
                patch("service.jobs.dingtalk_table.fetch_record", return_value=record),
                patch("service.jobs._task_docs_folder_id", return_value="shared-folder"),
                patch("service.jobs._task_input_dir", return_value=root / "raw"),
                patch("service.jobs._task_annotation_path", return_value=annotation_path),
                patch("service.jobs.generate_delivery_tables", return_value=outputs),
                patch(
                    "service.jobs.dingtalk_docs.upload_file",
                    side_effect=lambda config, path, folder_id: f"https://example/{path.stem}",
                ) as upload_file,
                patch("service.jobs._filter_links_for_existing_fields", side_effect=lambda configs, links: links),
                patch("service.jobs.dingtalk_table.mark_links") as mark_links,
            ):
                result = run_generate_delivery_tables("record")

        self.assertTrue(result["ok"])
        self.assertEqual(upload_file.call_count, 3)
        self.assertEqual([item.args[2] for item in upload_file.call_args_list], ["shared-folder"] * 3)
        written_links = mark_links.call_args.args[2]
        self.assertEqual(set(written_links), {"meituanData", "elemeData", "jdData"})


if __name__ == "__main__":
    unittest.main()
