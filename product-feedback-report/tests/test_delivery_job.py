from __future__ import annotations

import io
import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import MagicMock, call, patch

from openpyxl import Workbook

from service import dingtalk_docs, dingtalk_table
from service.jobs import (
    _delivery_completion_message,
    _empty_delivery_platforms,
    _ensure_dingtalk_input_attachment,
    _ensure_local_upload_folder,
    _single_ai_table_attachment,
    _single_local_xlsx,
    _task_annotation_path,
    _task_delivery_input_dir,
    _task_docs_folder_id,
    _task_folder_name,
    run_generate_delivery_tables,
    run_generate_report,
    run_prepare_product_menu,
)
from scripts.reporting.html import ReportBuildResult


class DeliveryJobTests(unittest.TestCase):
    def _xlsx_bytes(self) -> bytes:
        buffer = io.BytesIO()
        workbook = Workbook()
        workbook.active.append(["测试"])
        workbook.save(buffer)
        return buffer.getvalue()

    def _write_delivery_output(self, path: Path, product: str | None = None, sales: float = 1) -> None:
        wb = Workbook()
        ws = wb.active
        ws.append(["排名", "商品名称", "销量"])
        if product is not None:
            ws.append([1, product, sales])
        wb.save(path)

    def test_report_job_uses_only_current_row_products_and_writes_existing_report_field(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            output = root / "霸王茶姬：糯青山柠檬奶、雾红尘柠檬奶 20260718.html"
            output.write_text("<html></html>", encoding="utf-8")
            mt = root / "美团.xlsx"
            elm = root / "饿了么.xlsx"
            social = root / "雾红尘.xlsx"
            unrelated_social = root / "不应进入.xlsx"
            for path in (mt, elm, social, unrelated_social):
                path.write_bytes(b"xlsx")
            configs = {
                "dingtalk": {},
                "dingtalk_docs": {},
                "field_mapping": {"fields": {}},
                "report_rules": {"outputDirectory": str(root / "outputs")},
            }
            main_record = dingtalk_table.TaskRecord(
                "record",
                {
                    "brand": {"name": "霸王茶姬"},
                    "productName": ["糯青山柠檬奶", "雾红尘柠檬奶"],
                    "launchDate": "2026-07-18",
                    "meituanData": {"name": "美团.xlsx", "link": "mt"},
                    "elemeData": {"name": "饿了么.xlsx", "link": "elm"},
                },
            )
            social_record = dingtalk_table.TaskRecord(
                "social",
                {
                    "brand": {"name": "霸王茶姬"},
                    "productName": ["雾红尘柠檬奶"],
                    "launchDate": "2026-06-18",
                    "day30": "2026-07-18",
                    "report": {"text": "社媒评论统计.xlsx", "link": "social"},
                },
            )
            unrelated_record = dingtalk_table.TaskRecord(
                "unrelated-social",
                {
                    "brand": {"name": "霸王茶姬"},
                    "productName": ["不应进入报告"],
                    "launchDate": "2026-06-18",
                    "day30": "2026-07-18",
                    "report": {"text": "社媒评论统计.xlsx", "link": "unrelated"},
                },
            )

            def download(_config, value, _output_dir):
                link = value.get("link") if isinstance(value, dict) else ""
                return {
                    "mt": mt,
                    "elm": elm,
                    "social": social,
                    "unrelated": unrelated_social,
                }.get(link)

            with (
                patch("service.jobs.load_configs", return_value=configs),
                patch("service.jobs.dingtalk_table.mark_status"),
                patch("service.jobs.dingtalk_table.fetch_record", return_value=main_record),
                patch(
                    "service.jobs.dingtalk_table.fetch_records",
                    return_value=[social_record, unrelated_record],
                ),
                patch("service.jobs._ensure_local_upload_folder", return_value="dated-folder") as ensure_folder,
                patch("service.jobs.dingtalk_docs.download_linked_file", side_effect=download),
                patch("service.jobs.generate_report", return_value=ReportBuildResult(output, ("图片缺失",))) as generate,
                patch("service.jobs.dingtalk_docs.upload_file", return_value="https://example/report") as upload,
                patch("service.jobs.dingtalk_table.mark_links") as mark_links,
                patch("service.jobs.dingtalk_table.update_feedback") as update_feedback,
            ):
                result = run_generate_report("record")

        grouped_record = ensure_folder.call_args.args[1]
        self.assertIs(grouped_record, main_record)
        self.assertEqual(
            grouped_record.cells["productName"],
            ["糯青山柠檬奶", "雾红尘柠檬奶"],
        )
        self.assertEqual(generate.call_args.args[2], ["糯青山柠檬奶", "雾红尘柠檬奶"])
        self.assertNotIn("不应进入报告", generate.call_args.args[2])
        upload.assert_called_once_with({}, output, "dated-folder")
        self.assertEqual(mark_links.call_args.args[2], {"report": (output.name, "https://example/report")})
        self.assertIn("图片缺失", update_feedback.call_args.args[2])
        self.assertEqual(
            [item.args[2] for item in update_feedback.call_args_list],
            [
                "1/4 正在读取当前行的品牌、报告日期和关注新品",
                "2/4 正在下载外卖数表和关注新品的社媒评论统计表",
                "3/4 正在查找关注新品的产品信息并生成跟踪报告",
                "4/4 正在上传跟踪报告到钉钉文档",
                "报告已生成；部分资料暂无法获取：图片缺失",
            ],
        )
        self.assertEqual(result["warnings"], ["图片缺失"])

    def test_report_job_failure_feedback_names_the_active_stage(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            mt = root / "美团.xlsx"
            elm = root / "饿了么.xlsx"
            output = root / "报告.html"
            for path in (mt, elm, output):
                path.write_bytes(b"data")
            configs = {
                "dingtalk": {},
                "dingtalk_docs": {},
                "field_mapping": {"fields": {}},
                "report_rules": {"outputDirectory": str(root / "outputs")},
            }
            record = dingtalk_table.TaskRecord(
                "record",
                {
                    "brand": {"name": "霸王茶姬"},
                    "productName": ["糯青山柠檬奶"],
                    "launchDate": "2026-07-18",
                    "meituanData": {"link": "mt"},
                    "elemeData": {"link": "elm"},
                },
            )

            for failed_stage, expected_stage in (
                ("read", "读取当前行信息"),
                ("download", "下载资料"),
                ("generate", "生成报告"),
                ("upload", "上传报告"),
            ):
                with self.subTest(stage=failed_stage):
                    def download(_config, value, _output_dir):
                        if failed_stage == "download":
                            raise RuntimeError("模拟下载失败")
                        return {"mt": mt, "elm": elm}.get(value.get("link"))

                    with (
                        patch("service.jobs.load_configs", return_value=configs),
                        patch("service.jobs.dingtalk_table.mark_status"),
                        patch(
                            "service.jobs.dingtalk_table.fetch_record",
                            side_effect=(
                                RuntimeError("模拟读取失败")
                                if failed_stage == "read"
                                else None
                            ),
                            return_value=record,
                        ),
                        patch("service.jobs.dingtalk_table.fetch_records", return_value=[]),
                        patch("service.jobs._ensure_local_upload_folder", return_value="folder"),
                        patch("service.jobs.dingtalk_docs.download_linked_file", side_effect=download),
                        patch(
                            "service.jobs.generate_report",
                            side_effect=(
                                RuntimeError("模拟生成失败")
                                if failed_stage == "generate"
                                else None
                            ),
                            return_value=ReportBuildResult(output, ()),
                        ),
                        patch(
                            "service.jobs.dingtalk_docs.upload_file",
                            side_effect=(
                                RuntimeError("模拟上传失败")
                                if failed_stage == "upload"
                                else None
                            ),
                            return_value="https://example/report",
                        ),
                        patch("service.jobs.dingtalk_table.mark_links"),
                        patch("service.jobs.dingtalk_table.update_feedback"),
                        patch("service.jobs.dingtalk_table.mark_failed") as mark_failed,
                    ):
                        with self.assertRaises(RuntimeError):
                            run_generate_report("record")

                    self.assertIn(
                        f"报告生成失败（{expected_stage}）",
                        mark_failed.call_args.args[2],
                    )

    def test_report_job_rejects_empty_current_row_products_in_first_stage(self) -> None:
        configs = {
            "dingtalk": {},
            "field_mapping": {"fields": {}},
            "report_rules": {},
        }
        record = dingtalk_table.TaskRecord(
            "record",
            {
                "brand": {"name": "霸王茶姬"},
                "productName": [],
                "launchDate": "2026-07-18",
            },
        )
        with (
            patch("service.jobs.load_configs", return_value=configs),
            patch("service.jobs.dingtalk_table.mark_status"),
            patch("service.jobs.dingtalk_table.fetch_record", return_value=record),
            patch("service.jobs.dingtalk_table.update_feedback"),
            patch("service.jobs.dingtalk_table.mark_failed") as mark_failed,
        ):
            with self.assertRaisesRegex(RuntimeError, "关注新品"):
                run_generate_report("record")

        self.assertIn(
            "报告生成失败（读取当前行信息）：“关注新品”为空",
            mark_failed.call_args.args[2],
        )

    def test_report_job_final_feedback_failure_does_not_change_success_status(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            mt = root / "美团.xlsx"
            elm = root / "饿了么.xlsx"
            output = root / "报告.html"
            for path in (mt, elm, output):
                path.write_bytes(b"data")
            configs = {
                "dingtalk": {},
                "dingtalk_docs": {},
                "field_mapping": {"fields": {}},
                "report_rules": {"outputDirectory": str(root / "outputs")},
            }
            record = dingtalk_table.TaskRecord(
                "record",
                {
                    "brand": {"name": "霸王茶姬"},
                    "productName": ["糯青山柠檬奶"],
                    "launchDate": "2026-07-18",
                    "meituanData": {"link": "mt"},
                    "elemeData": {"link": "elm"},
                },
            )

            def download(_config, value, _output_dir):
                return {"mt": mt, "elm": elm}.get(value.get("link"))

            with (
                patch("service.jobs.load_configs", return_value=configs),
                patch("service.jobs.dingtalk_table.mark_status"),
                patch("service.jobs.dingtalk_table.fetch_record", return_value=record),
                patch("service.jobs.dingtalk_table.fetch_records", return_value=[]),
                patch("service.jobs._ensure_local_upload_folder", return_value="folder"),
                patch("service.jobs.dingtalk_docs.download_linked_file", side_effect=download),
                patch(
                    "service.jobs.generate_report",
                    return_value=ReportBuildResult(output, ()),
                ),
                patch("service.jobs.dingtalk_docs.upload_file", return_value="https://example/report"),
                patch("service.jobs.dingtalk_table.mark_links") as mark_links,
                patch(
                    "service.jobs.dingtalk_table.update_feedback",
                    side_effect=[None, None, None, None, RuntimeError("反馈回写失败")],
                ),
                patch("service.jobs.dingtalk_table.mark_failed") as mark_failed,
            ):
                result = run_generate_report("record")

        self.assertTrue(result["ok"])
        mark_links.assert_called_once()
        mark_failed.assert_not_called()

    def test_all_platform_results_use_link_cell_format(self) -> None:
        configs = {
            "dingtalk": {"baseId": "base", "tableId": "table"},
            "field_mapping": {
                "fields": {
                    "productMenu": {"fieldId": "menu", "cellType": "link"},
                    "meituanData": {"fieldId": "mt", "cellType": "link"},
                    "elemeData": {"fieldId": "elm", "cellType": "link"},
                    "jdData": {"fieldId": "jd", "cellType": "link"},
                }
            },
        }
        with patch("service.dingtalk_table._update_cells") as update_cells:
            dingtalk_table.mark_links(
                configs,
                "record",
                {
                    "productMenu": ("产品清单.xlsx", "https://example/menu"),
                    "meituanData": ("美团.xlsx", "https://example/mt"),
                    "elemeData": ("饿了么.xlsx", "https://example/elm"),
                    "jdData": ("京东.xlsx", "https://example/jd"),
                },
            )

        cells = update_cells.call_args.args[2]
        self.assertEqual(cells["productMenu"], {"text": "产品清单.xlsx", "link": "https://example/menu"})
        self.assertEqual(cells["meituanData"], {"text": "美团.xlsx", "link": "https://example/mt"})
        self.assertEqual(cells["elemeData"], {"text": "饿了么.xlsx", "link": "https://example/elm"})
        self.assertEqual(cells["jdData"], {"text": "京东.xlsx", "link": "https://example/jd"})

    def test_report_cell_format_follows_table_specific_configuration(self) -> None:
        attachment_configs = {
            "field_mapping": {
                "fields": {"report": {"fieldId": "report-field", "cellType": "attachment"}}
            }
        }
        link_configs = {
            "field_mapping": {
                "fields": {"report": {"fieldId": "report-field", "cellType": "link"}}
            }
        }
        with patch("service.dingtalk_table._update_cells") as update_cells:
            dingtalk_table.mark_links(
                attachment_configs,
                "record",
                {"report": ("报告.html", "https://example/report")},
            )
            attachment_cells = update_cells.call_args.args[2]
            dingtalk_table.mark_links(
                link_configs,
                "record",
                {"report": ("社媒统计.xlsx", "https://example/social")},
            )
            link_cells = update_cells.call_args.args[2]

        self.assertEqual(
            attachment_cells["report"],
            [{"url": "https://example/report", "name": "报告.html"}],
        )
        self.assertEqual(
            link_cells["report"],
            {"text": "社媒统计.xlsx", "link": "https://example/social"},
        )

    def test_link_results_clear_without_attachment_polling(self) -> None:
        with patch("service.dingtalk_table._update_cells") as update_cells:
            dingtalk_table.clear_link_fields(
                {}, "record", ("meituanData", "elemeData", "jdData")
            )
            dingtalk_table.update_attachment_fields(
                {}, "record", {"deliveryData": ("外卖数据.xlsx", "https://example/input")}
            )

        self.assertEqual(
            update_cells.call_args_list[0].args[2],
            {"meituanData": "", "elemeData": "", "jdData": ""},
        )
        self.assertEqual(
            update_cells.call_args_list[1].args[2],
            {"deliveryData": [{"url": "https://example/input", "name": "外卖数据.xlsx"}]},
        )

    def test_attachment_clear_retries_until_the_fields_are_empty(self) -> None:
        stale = dingtalk_table.TaskRecord(
            "record", {"deliveryData": [{"name": "旧外卖数据.xlsx", "url": "https://example/old"}]}
        )
        empty = dingtalk_table.TaskRecord("record", {"deliveryData": None})
        with (
            patch("service.dingtalk_table._update_cells") as update_cells,
            patch("service.dingtalk_table.fetch_record", side_effect=[stale, empty]) as fetch_record,
            patch("service.dingtalk_table.time.sleep") as sleep,
        ):
            dingtalk_table.clear_attachment_fields({}, "record", ("deliveryData",))

        update_cells.assert_called_once_with({}, "record", {"deliveryData": ""})
        self.assertEqual(fetch_record.call_count, 2)
        sleep.assert_called_once_with(dingtalk_table.ATTACHMENT_CLEAR_DELAY_SECONDS)

    def test_attachment_clear_fails_when_the_old_value_never_disappears(self) -> None:
        stale = dingtalk_table.TaskRecord(
            "record", {"deliveryData": [{"name": "旧外卖数据.xlsx", "url": "https://example/old"}]}
        )
        with (
            patch("service.dingtalk_table._update_cells"),
            patch("service.dingtalk_table.fetch_record", return_value=stale) as fetch_record,
            patch("service.dingtalk_table.time.sleep"),
        ):
            with self.assertRaisesRegex(RuntimeError, "旧结果附件清空失败"):
                dingtalk_table.clear_attachment_fields(
                    {}, "record", ("deliveryData",), attempts=3, delay_seconds=0
                )

        self.assertEqual(fetch_record.call_count, 3)

    def test_product_menu_job_stops_before_generation_when_link_clear_fails(self) -> None:
        configs = {"dingtalk": {}, "field_mapping": {"fields": {}}, "report_rules": {}}
        record = dingtalk_table.TaskRecord(
            "record",
            {"productMenu": [{"name": "旧产品清单.xlsx", "url": "https://example/old"}]},
        )
        with (
            patch("service.jobs.load_configs", return_value=configs),
            patch("service.jobs.dingtalk_table.mark_status"),
            patch("service.jobs.dingtalk_table.fetch_record", return_value=record),
            patch("service.jobs._ensure_local_upload_folder", return_value="shared-folder"),
            patch(
                "service.jobs.dingtalk_table.clear_link_fields",
                side_effect=RuntimeError("旧结果链接清空失败，请刷新钉钉 AI 表后重试。"),
            ),
            patch("service.jobs.prepare_product_menu") as prepare,
            patch("service.jobs.dingtalk_docs.upload_file") as upload,
            patch("service.jobs.dingtalk_table.update_feedback") as update_feedback,
            patch("service.jobs.dingtalk_table.mark_failed") as mark_failed,
        ):
            with self.assertRaisesRegex(RuntimeError, "旧结果链接清空失败"):
                run_prepare_product_menu("record")

        prepare.assert_not_called()
        upload.assert_not_called()
        self.assertIn("旧结果链接清空失败", update_feedback.call_args.args[2])
        self.assertIn("旧结果链接清空失败", mark_failed.call_args.args[2])

    def test_delivery_job_stops_before_generation_when_link_clear_fails(self) -> None:
        configs = {"dingtalk": {}, "field_mapping": {"fields": {}}, "report_rules": {}}
        record = dingtalk_table.TaskRecord(
            "record",
            {
                "meituanData": [{"name": "旧美团.xlsx", "url": "https://example/old"}],
                "elemeData": [{"name": "旧饿了么.xlsx", "url": "https://example/old"}],
                "jdData": [{"name": "旧京东.xlsx", "url": "https://example/old"}],
            },
        )
        with (
            patch("service.jobs.load_configs", return_value=configs),
            patch("service.jobs.dingtalk_table.mark_status"),
            patch("service.jobs.dingtalk_table.fetch_record", return_value=record),
            patch("service.jobs._ensure_local_upload_folder", return_value="shared-folder"),
            patch(
                "service.jobs.dingtalk_table.clear_link_fields",
                side_effect=RuntimeError("旧结果链接清空失败，请刷新钉钉 AI 表后重试。"),
            ),
            patch("service.jobs.generate_delivery_tables") as generate,
            patch("service.jobs.dingtalk_docs.upload_file") as upload,
            patch("service.jobs.dingtalk_table.mark_failed") as mark_failed,
        ):
            with self.assertRaisesRegex(RuntimeError, "旧结果链接清空失败"):
                run_generate_delivery_tables("record")

        generate.assert_not_called()
        upload.assert_not_called()
        self.assertIn("旧结果链接清空失败", mark_failed.call_args.args[2])

    def test_local_xlsx_is_uploaded_and_written_back_before_use(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            source = Path(tmp) / "外卖数据.xlsx"
            source.write_bytes(b"xlsx")
            record = dingtalk_table.TaskRecord(
                "record",
                {
                    "brand": {"name": "爷爷不泡茶"},
                    "productName": {"value": ["栀子龙眼椰", "奇香青柠"]},
                    "launchDate": "2026-07-17T00:00:00+08:00",
                    "deliveryData": str(source),
                },
            )
            configs = {"dingtalk_docs": {}, "report_rules": {"outputDirectory": str(Path(tmp) / "outputs")}}
            with (
                patch(
                    "service.jobs.dingtalk_docs.upload_file",
                    return_value="https://alidocs.dingtalk.com/i/nodes/raw-node",
                ) as upload_file,
                patch("service.jobs.dingtalk_table.update_attachment_fields") as update_fields,
                patch("service.jobs._ensure_local_upload_folder", return_value="created-folder"),
            ):
                folder_id = _ensure_dingtalk_input_attachment(configs, record, "deliveryData", "外卖数据", None)
            cached_bytes = (Path(tmp) / "outputs" / "record" / "raw_data" / "外卖数据.xlsx").read_bytes()

        self.assertEqual(folder_id, "created-folder")
        upload_file.assert_called_once_with(
            {},
            Path(tmp) / "outputs" / "record" / "raw_data" / "外卖数据.xlsx",
            "created-folder",
        )
        update_fields.assert_called_once_with(
            configs,
            "record",
            {"deliveryData": ("外卖数据.xlsx", "https://alidocs.dingtalk.com/i/nodes/raw-node")},
        )
        self.assertEqual(
            record.cells["deliveryData"],
            [{"url": "https://alidocs.dingtalk.com/i/nodes/raw-node", "name": "外卖数据.xlsx"}],
        )
        self.assertEqual(cached_bytes, b"xlsx")

    def test_dingtalk_input_never_uses_local_filename_lookup(self) -> None:
        value = [{"name": "外卖数据.xlsx", "url": "https://alidocs.dingtalk.com/i/nodes/raw-node"}]
        record = dingtalk_table.TaskRecord("record", {"deliveryData": value})
        with (
            patch("service.jobs._single_local_xlsx") as local_lookup,
            patch("service.jobs._ensure_local_upload_folder") as ensure_folder,
            patch(
                "service.jobs.dingtalk_docs.folder_id_for_linked_node",
                return_value="shared-folder",
            ),
        ):
            folder_id = _ensure_dingtalk_input_attachment(
                {"dingtalk_docs": {}}, record, "deliveryData", "外卖数据", "shared-folder"
            )

        self.assertEqual(folder_id, "shared-folder")
        local_lookup.assert_not_called()
        ensure_folder.assert_not_called()

    def test_linked_delivery_input_is_copied_into_the_dated_target_folder(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            downloaded = root / "霸王茶姬-20260718.xlsx"
            downloaded.write_bytes(b"xlsx")
            configs = {
                "dingtalk_docs": {},
                "report_rules": {"outputDirectory": str(root / "outputs")},
            }
            record = dingtalk_table.TaskRecord(
                "record",
                {
                    "deliveryData": [
                        {
                            "name": downloaded.name,
                            "url": "https://alidocs.dingtalk.com/i/nodes/raw-node",
                        }
                    ]
                },
            )
            with (
                patch(
                    "service.jobs.dingtalk_docs.folder_id_for_linked_node",
                    return_value="old-undated-folder",
                ),
                patch(
                    "service.jobs.dingtalk_docs.download_linked_file",
                    return_value=downloaded,
                ) as download,
                patch(
                    "service.jobs.dingtalk_docs.upload_file",
                    return_value="https://docs/destination-copy",
                ) as upload,
                patch("service.jobs.dingtalk_table.update_attachment_fields") as update,
            ):
                folder_id = _ensure_dingtalk_input_attachment(
                    configs,
                    record,
                    "deliveryData",
                    "外卖数据",
                    "dated-folder",
                )

        self.assertEqual(folder_id, "dated-folder")
        download.assert_called_once()
        upload.assert_called_once_with({}, downloaded, "dated-folder")
        update.assert_called_once_with(
            configs,
            "record",
            {"deliveryData": (downloaded.name, "https://docs/destination-copy")},
        )
        self.assertEqual(
            record.cells["deliveryData"],
            [{"url": "https://docs/destination-copy", "name": downloaded.name}],
        )

    def test_linked_product_menu_is_copied_and_written_back_as_link(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            downloaded = root / "霸王茶姬-20260718-产品清单.xlsx"
            downloaded.write_bytes(b"xlsx")
            configs = {
                "dingtalk_docs": {},
                "field_mapping": {
                    "fields": {"productMenu": {"fieldId": "menu", "cellType": "link"}}
                },
                "report_rules": {"outputDirectory": str(root / "outputs")},
            }
            record = dingtalk_table.TaskRecord(
                "record",
                {
                    "productMenu": {
                        "text": downloaded.name,
                        "link": "https://alidocs.dingtalk.com/i/nodes/menu-node",
                    }
                },
            )
            with (
                patch(
                    "service.jobs.dingtalk_docs.folder_id_for_linked_node",
                    return_value="old-folder",
                ),
                patch(
                    "service.jobs.dingtalk_docs.download_linked_file",
                    return_value=downloaded,
                ),
                patch(
                    "service.jobs.dingtalk_docs.upload_file",
                    return_value="https://alidocs.dingtalk.com/i/nodes/new-menu-node",
                ) as upload,
                patch("service.jobs.dingtalk_table.update_link_fields") as update,
            ):
                folder_id = _ensure_dingtalk_input_attachment(
                    configs, record, "productMenu", "产品清单", "dated-folder"
                )

        self.assertEqual(folder_id, "dated-folder")
        upload.assert_called_once_with({}, downloaded, "dated-folder")
        update.assert_called_once_with(
            configs,
            "record",
            {
                "productMenu": (
                    downloaded.name,
                    "https://alidocs.dingtalk.com/i/nodes/new-menu-node",
                )
            },
        )
        self.assertEqual(
            record.cells["productMenu"],
            {
                "text": downloaded.name,
                "link": "https://alidocs.dingtalk.com/i/nodes/new-menu-node",
            },
        )

    def test_ai_table_local_attachment_is_downloaded_promoted_and_written_back(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            record = dingtalk_table.TaskRecord(
                "record",
                {
                    "brand": {"name": "爷爷不泡茶"},
                    "productName": {"value": ["栀子龙眼椰", "奇香青柠"]},
                    "launchDate": "2026-07-17T00:00:00+08:00",
                    "deliveryData": [
                        {
                            "filename": "爷爷不泡茶-20260717.xlsx",
                            "resourceId": "table-resource",
                            "url": "https://example/table-attachment.xlsx",
                            "resourceUrl": "/core/api/resources/table-resource/detail",
                            "type": "xls",
                        }
                    ],
                },
            )
            configs = {
                "dingtalk_docs": {},
                "report_rules": {"outputDirectory": str(root / "outputs")},
            }
            response = MagicMock()
            response.__enter__.return_value.read.return_value = b"latest-xlsx"
            with (
                patch("service.jobs.urllib.request.urlopen", return_value=response) as urlopen,
                patch("service.jobs._ensure_local_upload_folder", return_value="task-folder") as ensure_folder,
                patch(
                    "service.jobs.dingtalk_docs.upload_file",
                    return_value="https://alidocs.dingtalk.com/i/nodes/new-node",
                ) as upload_file,
                patch("service.jobs.dingtalk_table.update_attachment_fields") as update_fields,
            ):
                folder_id = _ensure_dingtalk_input_attachment(
                    configs, record, "deliveryData", "外卖数据", None
                )
            cached = root / "outputs" / "record" / "raw_data" / "爷爷不泡茶-20260717.xlsx"
            cached_bytes = cached.read_bytes()

        self.assertEqual(folder_id, "task-folder")
        self.assertEqual(cached_bytes, b"latest-xlsx")
        self.assertEqual(urlopen.call_args.args[0].full_url, "https://example/table-attachment.xlsx")
        ensure_folder.assert_called_once_with(configs, record)
        upload_file.assert_called_once_with({}, cached, "task-folder")
        update_fields.assert_called_once_with(
            configs,
            "record",
            {
                "deliveryData": (
                    "爷爷不泡茶-20260717.xlsx",
                    "https://alidocs.dingtalk.com/i/nodes/new-node",
                )
            },
        )

    def test_ai_table_local_attachment_validation(self) -> None:
        valid = {
            "filename": "外卖数据.xlsx",
            "resourceId": "resource",
            "resourceUrl": "https://example/data.xlsx",
        }
        with self.assertRaisesRegex(RuntimeError, "仅支持一个"):
            _single_ai_table_attachment([valid, valid], "外卖数据")
        with self.assertRaisesRegex(RuntimeError, "只支持 xlsx"):
            _single_ai_table_attachment(
                [{**valid, "filename": "外卖数据.csv"}], "外卖数据"
            )
        with self.assertRaisesRegex(RuntimeError, "缺少可下载 URL"):
            _single_ai_table_attachment(
                [{"filename": "外卖数据.xlsx", "resourceId": "resource"}], "外卖数据"
            )
        self.assertEqual(
            _single_ai_table_attachment(
                [
                    {
                        "filename": "阿嬷手作-20260717.xlsx",
                        "resourceId": "resource",
                        "url": "https://example/signed.xlsx",
                        "resourceUrl": "/core/api/resources/resource/detail",
                    }
                ],
                "外卖数据",
            ),
            ("阿嬷手作-20260717.xlsx", "https://example/signed.xlsx"),
        )

    def test_local_upload_folder_name_uses_brand_and_normalized_products(self) -> None:
        record = dingtalk_table.TaskRecord(
            "record",
            {
                "brand": {"id": "brand", "name": "爷爷不泡茶"},
                "productName": {"refFieldType": "text", "value": ["栀子龙眼椰, 奇香青柠", "兰香，青柠"]},
                "launchDate": "2026-07-17",
            },
        )

        self.assertEqual(
            _task_folder_name(record),
            "爷爷不泡茶：栀子龙眼椰、奇香青柠、兰香、青柠 20260717",
        )

    def test_local_upload_folder_uses_report_date_and_rejects_missing_metadata(self) -> None:
        configs = {"dingtalk_docs": {"localUploadRootFolderId": "root"}}
        complete = dingtalk_table.TaskRecord(
            "record",
            {
                "brand": {"name": "爷爷不泡茶"},
                "productName": {"value": ["栀子龙眼椰", "奇香青柠"]},
                "launchDate": "2026-07-17T00:00:00+08:00",
            },
        )
        with (
            patch(
                "service.jobs.dingtalk_docs.ensure_local_upload_month_folder",
                return_value=("month-folder", "https://example/month-folder"),
            ) as ensure_month,
            patch("service.jobs.dingtalk_docs.child_folders", return_value=[]),
            patch(
                "service.jobs.dingtalk_docs.ensure_child_folder",
                return_value=("task-folder", "https://example/task-folder"),
            ) as ensure_folder,
        ):
            self.assertEqual(_ensure_local_upload_folder(configs, complete), "task-folder")
        ensure_month.assert_called_once_with(configs["dingtalk_docs"], 2026, 7)
        ensure_folder.assert_called_once_with(
            configs["dingtalk_docs"],
            "month-folder",
            "爷爷不泡茶：栀子龙眼椰、奇香青柠 20260717",
        )

        for cells, message in (
            ({"brand": {"name": "爷爷不泡茶"}, "productName": {"value": ["新品"]}}, "报告日期"),
            ({"launchDate": "2026-07-17", "productName": {"value": ["新品"]}}, "竞品品牌"),
            ({"launchDate": "2026-07-17", "brand": {"name": "爷爷不泡茶"}}, "关注新品"),
        ):
            with self.subTest(message=message):
                with self.assertRaisesRegex(RuntimeError, message):
                    _ensure_local_upload_folder(configs, dingtalk_table.TaskRecord("record", cells))

    def test_delivery_reuses_dated_social_folder_with_same_products(self) -> None:
        configs = {"dingtalk_docs": {}}
        record = dingtalk_table.TaskRecord(
            "record",
            {
                "brand": {"name": "霸王茶姬"},
                "productName": {"value": ["糯青山柠檬奶", "雾红尘柠檬奶"]},
                "launchDate": "2026-07-18",
            },
        )
        with (
            patch(
                "service.jobs.dingtalk_docs.ensure_local_upload_month_folder",
                return_value=("month", "https://example/month"),
            ),
            patch(
                "service.jobs.dingtalk_docs.child_folders",
                return_value=[
                    ("shared", "霸王茶姬：雾红尘柠檬奶、糯青山柠檬奶 20260718")
                ],
            ),
            patch("service.jobs.dingtalk_docs.ensure_child_folder") as ensure_child,
        ):
            folder_id = _ensure_local_upload_folder(configs, record)

        self.assertEqual(folder_id, "shared")
        ensure_child.assert_not_called()

    def test_local_upload_task_folder_reuses_year_month_and_task_hierarchy(self) -> None:
        config = {"localUploadRootFolderName": "原始文件：竞品新品跟踪反馈"}
        root = {"nodeId": "root", "name": "原始文件：竞品新品跟踪反馈", "nodeType": "folder"}
        with (
            patch("service.dingtalk_docs.list_root_nodes", return_value=[root]),
            patch(
                "service.dingtalk_docs.ensure_child_folder",
                side_effect=[
                    ("year", "https://example/year"),
                    ("month", "https://example/month"),
                    ("task", "https://example/task"),
                ],
            ) as ensure_child,
        ):
            actual = dingtalk_docs.ensure_local_upload_task_folder(
                config, 2026, 7, "爷爷不泡茶：栀子龙眼椰、奇香青柠"
            )

        self.assertEqual(actual, ("task", "https://example/task"))
        self.assertEqual(
            ensure_child.call_args_list,
            [
                call(config, "root", "2026年"),
                call(config, "year", "2026年7月"),
                call(config, "month", "爷爷不泡茶：栀子龙眼椰、奇香青柠"),
            ],
        )

    def test_local_upload_root_must_already_exist(self) -> None:
        with (
            patch("service.dingtalk_docs.list_root_nodes", return_value=[]),
            patch("service.dingtalk_docs.call_docs_tool", return_value={"documents": []}),
        ):
            with self.assertRaisesRegex(RuntimeError, "未找到钉钉文档归档根目录"):
                dingtalk_docs.local_upload_root_folder_id({})

    def test_local_upload_root_can_be_found_in_knowledge_base_search(self) -> None:
        expected = {
            "nodeId": "archive-root",
            "name": "原始文件：竞品新品跟踪反馈",
            "nodeType": "folder",
            "extension": "folder",
        }
        with (
            patch("service.dingtalk_docs.list_root_nodes", return_value=[]),
            patch(
                "service.dingtalk_docs.call_docs_tool",
                return_value={"documents": [expected]},
            ) as search,
        ):
            actual = dingtalk_docs.local_upload_root_folder_id({})

        self.assertEqual(actual, "archive-root")
        search.assert_called_once_with(
            {},
            "search_documents",
            {"keyword": "原始文件：竞品新品跟踪反馈", "pageSize": 30},
        )

    def test_local_input_rejects_directory_non_xlsx_and_multiple_files(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            text_file = root / "input.csv"
            first = root / "first.xlsx"
            second = root / "second.xlsx"
            text_file.write_text("data")
            first.write_bytes(b"one")
            second.write_bytes(b"two")

            with self.assertRaisesRegex(RuntimeError, "本机目录"):
                _single_local_xlsx(str(root), "外卖数据")
            with self.assertRaisesRegex(RuntimeError, "只支持 xlsx"):
                _single_local_xlsx(str(text_file), "外卖数据")
            with self.assertRaisesRegex(RuntimeError, "多个本机文件"):
                _single_local_xlsx([str(first), str(second)], "外卖数据")

    def test_delivery_input_does_not_fall_back_to_all_data_field(self) -> None:
        record = dingtalk_table.TaskRecord(
            "record",
            {"allData": [{"name": "旧数据.xlsx", "url": "https://alidocs.dingtalk.com/i/nodes/old-node"}]},
        )
        with self.assertRaisesRegex(RuntimeError, "外卖数据.*缺少"):
            _task_delivery_input_dir({"report_rules": {"outputDirectory": "/tmp"}}, record)

    def test_attachment_name_is_supported_when_downloading_linked_xlsx(self) -> None:
        value = [{"name": "产品清单.xlsx", "url": "https://alidocs.dingtalk.com/i/nodes/menu-node"}]
        with tempfile.TemporaryDirectory() as tmp:
            expected = Path(tmp) / "产品清单.xlsx"
            with patch("service.dingtalk_docs.download_file", return_value=expected) as download_file:
                actual = dingtalk_docs.download_linked_file({}, value, Path(tmp))

        self.assertEqual(actual, expected)
        self.assertEqual(download_file.call_args.args[1]["nodeId"], "menu-node")
        self.assertEqual(download_file.call_args.args[1]["name"], "产品清单.xlsx")

    def test_download_file_supports_resource_url_list_and_signed_headers(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            response = MagicMock()
            response.__enter__.return_value.read.return_value = self._xlsx_bytes()

            def call_docs_tool(config, tool_name, payload):
                if tool_name == "get_document_info":
                    return {"workspaceId": "space", "updateTime": 1}
                return {"resourceUrl": ["https://example/file.xlsx"], "headers": {"X-Signature": "signed"}}

            with (
                patch(
                    "service.dingtalk_docs.call_docs_tool",
                    side_effect=call_docs_tool,
                ),
                patch("service.dingtalk_docs.urllib.request.urlopen", return_value=response) as urlopen,
            ):
                output = dingtalk_docs.download_file(
                    {},
                    {"nodeId": "menu-node", "name": "产品清单.xlsx", "extension": "xlsx"},
                    Path(tmp),
                )

            request = urlopen.call_args.args[0]
            self.assertEqual(request.full_url, "https://example/file.xlsx")
            self.assertEqual(request.get_header("X-signature"), "signed")
            self.assertTrue(output.is_file())

    def test_download_file_reports_missing_mcp_download_url(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            with patch("service.dingtalk_docs.call_docs_tool", return_value={"success": True}):
                with self.assertRaisesRegex(
                    dingtalk_docs.DownloadUrlUnavailableError,
                    "钉钉文档 MCP 未返回可下载 URL",
                ):
                    dingtalk_docs.download_file(
                        {},
                        {"nodeId": "remote-node", "name": "数据.xlsx", "extension": "xlsx"},
                        Path(tmp),
                    )

    def test_versioned_cache_requires_matching_node_update_time_filename_and_hash(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            output = Path(tmp) / "数据.xlsx"
            output.write_bytes(self._xlsx_bytes())
            dingtalk_docs._write_download_cache(
                output, "node-id", {"updateTime": 123}
            )
            with (
                patch(
                    "service.dingtalk_docs.get_document_info",
                    return_value={"workspaceId": "space", "updateTime": 123},
                ),
                patch("service.dingtalk_docs.call_docs_tool") as mcp,
            ):
                actual = dingtalk_docs.download_file(
                    {},
                    {"nodeId": "node-id", "name": "数据.xlsx", "extension": "xlsx"},
                    Path(tmp),
                )

        self.assertEqual(actual, output)
        mcp.assert_not_called()

    def test_versioned_cache_is_invalidated_when_remote_update_time_changes(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            output = Path(tmp) / "数据.xlsx"
            output.write_bytes(self._xlsx_bytes())
            dingtalk_docs._write_download_cache(
                output, "node-id", {"updateTime": 123}
            )
            response = MagicMock()
            response.__enter__.return_value.read.return_value = self._xlsx_bytes()

            def call_docs_tool(config, tool_name, payload):
                return {"resourceUrl": "https://example/new.xlsx"}

            with (
                patch(
                    "service.dingtalk_docs.get_document_info",
                    return_value={"workspaceId": "space", "updateTime": 124},
                ),
                patch("service.dingtalk_docs.call_docs_tool", side_effect=call_docs_tool) as mcp,
                patch("service.dingtalk_docs.urllib.request.urlopen", return_value=response),
            ):
                actual = dingtalk_docs.download_file(
                    {},
                    {"nodeId": "node-id", "name": "数据.xlsx", "extension": "xlsx"},
                    Path(tmp),
                )

            manifest = json.loads(
                dingtalk_docs._cache_manifest_path(output).read_text(encoding="utf-8")
            )

        self.assertEqual(actual, output)
        self.assertEqual(mcp.call_args.args[1], "download_file")
        self.assertEqual(manifest["updateTime"], "124")

    def test_upload_explicitly_keeps_office_file_as_downloadable_attachment(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            source = Path(tmp) / "产品清单.xlsx"
            source.write_bytes(b"xlsx")
            response = MagicMock(status=200)
            response.__enter__.return_value = response
            calls: list[tuple[str, dict]] = []

            def call_docs_tool(config, tool_name, payload):
                calls.append((tool_name, payload))
                if tool_name == "get_file_upload_info":
                    return {"uploadKey": "upload-key", "resourceUrl": "https://example/upload", "headers": {}}
                return {"nodeId": "menu-node", "name": source.name}

            with (
                patch("service.dingtalk_docs.delete_existing_file") as delete_existing,
                patch("service.dingtalk_docs.call_docs_tool", side_effect=call_docs_tool),
                patch("service.dingtalk_docs.urllib.request.urlopen", return_value=response),
            ):
                dingtalk_docs.upload_file({}, source, "folder")

        delete_existing.assert_called_once_with({}, "folder", source.name)
        commit_payload = next(payload for tool, payload in calls if tool == "commit_uploaded_file")
        self.assertIs(commit_payload["convertToOnlineDoc"], False)

    def test_overwrite_removes_exact_and_numbered_duplicate_files(self) -> None:
        nodes = [
            {"nodeId": "exact", "name": "产品清单", "extension": "xlsx"},
            {"nodeId": "copy-1", "name": "产品清单(1)", "extension": "xlsx"},
            {"nodeId": "copy-2", "name": "产品清单(2).xlsx", "extension": "xlsx"},
            {"nodeId": "other", "name": "其他文件", "extension": "xlsx"},
        ]
        with (
            patch("service.dingtalk_docs.list_nodes", return_value=nodes),
            patch("service.dingtalk_docs.call_docs_tool") as mcp,
        ):
            dingtalk_docs.delete_existing_file({}, "folder", "产品清单.xlsx")

        deleted_ids = [
            item.args[2]["nodeId"]
            for item in mcp.call_args_list
            if item.args[1] == "delete_document"
        ]
        self.assertEqual(deleted_ids, ["exact", "copy-1", "copy-2"])

    def test_dingtalk_link_takes_priority_over_stale_same_named_local_file(self) -> None:
        value = [{"name": "产品清单.xlsx", "url": "https://alidocs.dingtalk.com/i/nodes/menu-node"}]
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            stale = root / "desktop" / "产品清单.xlsx"
            latest = root / "current-task" / "产品清单.xlsx"
            with (
                patch("service.dingtalk_docs.local_file_from_attachment_name", return_value=stale) as local_lookup,
                patch("service.dingtalk_docs.download_file", return_value=latest) as download_file,
            ):
                actual = dingtalk_docs.download_linked_file({}, value, latest.parent)

        self.assertEqual(actual, latest)
        self.assertEqual(download_file.call_args.args[1]["nodeId"], "menu-node")
        self.assertEqual(download_file.call_args.args[2], latest.parent)
        local_lookup.assert_not_called()

    def test_annotation_relies_on_centralized_download_and_version_cache(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            expected = root / "record" / "产品清单.xlsx"
            configs = {"dingtalk_docs": {}, "report_rules": {"outputDirectory": str(root)}}
            record = dingtalk_table.TaskRecord(
                "record",
                {"productMenu": [{"name": "产品清单.xlsx", "url": "https://alidocs.dingtalk.com/i/nodes/menu-node"}]},
            )
            with patch(
                "service.jobs.dingtalk_docs.download_linked_file", return_value=expected
            ) as download:
                actual = _task_annotation_path(configs, record)

        self.assertEqual(actual, expected)
        self.assertEqual(download.call_args.args[2], root / "record")

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

    def test_empty_platform_detection_keeps_zero_sales_products(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            outputs = {
                "meituanData": root / "美团.xlsx",
                "elemeData": root / "饿了么.xlsx",
                "jdData": root / "京东.xlsx",
            }
            self._write_delivery_output(outputs["meituanData"], "零销量商品", 0)
            self._write_delivery_output(outputs["elemeData"])
            self._write_delivery_output(outputs["jdData"])

            empty = _empty_delivery_platforms(outputs)

        self.assertEqual(empty, ["饿了么", "京东"])

    def test_delivery_completion_messages_list_empty_platforms_without_suffixes(self) -> None:
        self.assertEqual(_delivery_completion_message([]), "外卖数据统计完毕")
        self.assertEqual(_delivery_completion_message(["京东"]), "外卖数据统计完毕，京东无数据")
        self.assertEqual(_delivery_completion_message(["美团", "京东"]), "外卖数据统计完毕，美团、京东无数据")
        self.assertEqual(
            _delivery_completion_message(["美团", "饿了么", "京东"]),
            "外卖数据统计完毕，美团、饿了么、京东无数据",
        )

    def test_product_menu_keeps_progress_steps_and_uses_review_completion_message(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            output = root / "茉莉奶白-20260710-产品清单.xlsx"
            output.write_bytes(b"menu")
            configs = {
                "dingtalk": {},
                "dingtalk_docs": {},
                "field_mapping": {"fields": {}},
                "report_rules": {"outputDirectory": str(root)},
            }
            record = dingtalk_table.TaskRecord(
                "record",
                {"deliveryData": [{"name": "外卖数据.xlsx", "url": "https://alidocs.dingtalk.com/i/nodes/raw-node"}]},
            )
            events: list[str] = []

            def prepare(record_id, input_dir, output_dir, progress_callback):
                progress_callback("2/4 正在提取产品清单")
                progress_callback("3/4 正在标记上新不满32天的产品")
                return output

            with (
                patch("service.jobs.load_configs", return_value=configs),
                patch("service.jobs.dingtalk_table.mark_status"),
                patch("service.jobs.dingtalk_table.fetch_record", return_value=record),
                patch("service.jobs._task_delivery_input_dir", return_value=root / "raw"),
                patch("service.jobs._ensure_dingtalk_input_attachment", return_value="shared-folder"),
                patch("service.jobs.prepare_product_menu", side_effect=prepare),
                patch("service.jobs._ensure_local_upload_folder", return_value="shared-folder"),
                patch("service.jobs.dingtalk_docs.upload_file", return_value="https://example/menu"),
                patch("service.jobs.dingtalk_table.mark_links"),
                patch(
                    "service.jobs.dingtalk_table.clear_link_fields",
                    side_effect=lambda *args: events.append("clear"),
                ) as clear_fields,
                patch("service.jobs.dingtalk_table.update_feedback") as update_feedback,
            ):
                result = run_prepare_product_menu("record")

        self.assertTrue(result["ok"])
        clear_fields.assert_called_once_with(configs, "record", ("productMenu",))
        self.assertEqual(events, ["clear"])
        self.assertEqual(
            [call.args[2] for call in update_feedback.call_args_list],
            [
                "1/4 正在读取外卖数据",
                "2/4 正在提取产品清单",
                "3/4 正在标记上新不满32天的产品",
                "4/4 正在上传到钉钉文档",
                "已提取产品清单，请确认在售不满30日的新品上新日期标注无遗漏后，再进行外卖数据统计",
            ],
        )

    def test_delivery_job_uploads_only_nonempty_results_and_leaves_empty_field_blank(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            annotation_path = root / "产品清单.xlsx"
            annotation_path.write_bytes(b"annotation")
            outputs = {
                "meituanData": root / "茉莉奶白-20260710-美团.xlsx",
                "elemeData": root / "茉莉奶白-20260710-饿了么.xlsx",
                "jdData": root / "茉莉奶白-20260710-京东.xlsx",
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
                    "deliveryData": [
                        {"name": "外卖数据.xlsx", "url": "https://alidocs.dingtalk.com/i/nodes/raw-node"}
                    ],
                    "productMenu": [
                        {"name": "产品清单.xlsx", "url": "https://alidocs.dingtalk.com/i/nodes/menu-node"}
                    ],
                },
            )
            events: list[str] = []

            with (
                patch("service.jobs.load_configs", return_value=configs),
                patch("service.jobs.dingtalk_table.mark_status"),
                patch("service.jobs.dingtalk_table.fetch_record", return_value=record),
                patch("service.jobs._ensure_local_upload_folder", return_value="shared-folder"),
                patch(
                    "service.jobs.dingtalk_docs.folder_id_for_linked_node",
                    return_value="shared-folder",
                ),
                patch("service.jobs._task_delivery_input_dir", return_value=root / "raw"),
                patch("service.jobs._task_annotation_path", return_value=annotation_path),
                patch(
                    "service.jobs.generate_delivery_tables",
                    side_effect=lambda *args: events.append("generate") or outputs,
                ),
                patch("service.jobs._empty_delivery_platforms", return_value=["京东"]),
                patch(
                    "service.jobs.dingtalk_docs.upload_file",
                    side_effect=lambda config, path, folder_id: f"https://example/{path.stem}",
                ) as upload_file,
                patch("service.jobs._filter_links_for_existing_fields", side_effect=lambda configs, links: links),
                patch(
                    "service.jobs.dingtalk_table.clear_link_fields",
                    side_effect=lambda *args: events.append("clear"),
                ) as clear_fields,
                patch("service.jobs.dingtalk_table.mark_links", side_effect=lambda *args: events.append("links")) as mark_links,
                patch(
                    "service.jobs.dingtalk_table.update_feedback",
                    side_effect=lambda configs, record_id, message: events.append(message),
                ) as update_feedback,
            ):
                result = run_generate_delivery_tables("record")
            jd_file_was_removed = not outputs["jdData"].exists()

        self.assertTrue(result["ok"])
        clear_fields.assert_called_once_with(configs, "record", ("meituanData", "elemeData", "jdData"))
        self.assertLess(events.index("clear"), events.index("generate"))
        self.assertEqual(upload_file.call_count, 2)
        self.assertEqual([item.args[2] for item in upload_file.call_args_list], ["shared-folder"] * 2)
        self.assertEqual(
            [item.args[1].name for item in upload_file.call_args_list],
            ["茉莉奶白-20260710-美团.xlsx", "茉莉奶白-20260710-饿了么.xlsx"],
        )
        self.assertTrue(jd_file_was_removed)
        written_links = mark_links.call_args.args[2]
        self.assertEqual(set(written_links), {"meituanData", "elemeData"})
        self.assertEqual(set(result["outputs"]), {"meituanData", "elemeData"})
        self.assertEqual(result["emptyPlatforms"], ["京东"])
        self.assertEqual(
            [call.args[2] for call in update_feedback.call_args_list],
            [
                "1/3 正在下载外卖数据和产品清单",
                "2/3 正在生成外卖数表",
                "3/3 正在上传外卖数表到钉钉文档",
                "外卖数据统计完毕，京东无数据",
            ],
        )
        self.assertLess(events.index("links"), events.index("外卖数据统计完毕，京东无数据"))

    def test_delivery_job_with_all_platforms_empty_uploads_nothing_and_still_completes(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            annotation_path = root / "产品清单.xlsx"
            annotation_path.write_bytes(b"annotation")
            outputs = {
                "meituanData": root / "茉莉奶白-20260710-美团.xlsx",
                "elemeData": root / "茉莉奶白-20260710-饿了么.xlsx",
                "jdData": root / "茉莉奶白-20260710-京东.xlsx",
            }
            for path in outputs.values():
                path.write_bytes(b"header-only")
            configs = {
                "dingtalk": {},
                "dingtalk_docs": {},
                "field_mapping": {"fields": {}},
                "report_rules": {"outputDirectory": str(root)},
            }
            record = dingtalk_table.TaskRecord(
                "record",
                {
                    "deliveryData": [
                        {"name": "外卖数据.xlsx", "url": "https://alidocs.dingtalk.com/i/nodes/raw-node"}
                    ],
                    "productMenu": [
                        {"name": "产品清单.xlsx", "url": "https://alidocs.dingtalk.com/i/nodes/menu-node"}
                    ],
                },
            )
            with (
                patch("service.jobs.load_configs", return_value=configs),
                patch("service.jobs.dingtalk_table.mark_status"),
                patch("service.jobs.dingtalk_table.fetch_record", return_value=record),
                patch("service.jobs._ensure_local_upload_folder", return_value="shared-folder"),
                patch(
                    "service.jobs.dingtalk_docs.folder_id_for_linked_node",
                    return_value="shared-folder",
                ),
                patch("service.jobs.dingtalk_table.clear_link_fields"),
                patch("service.jobs._task_delivery_input_dir", return_value=root / "raw"),
                patch("service.jobs._task_annotation_path", return_value=annotation_path),
                patch("service.jobs.generate_delivery_tables", return_value=outputs),
                patch(
                    "service.jobs._empty_delivery_platforms",
                    return_value=["美团", "饿了么", "京东"],
                ),
                patch("service.jobs.dingtalk_docs.upload_file") as upload_file,
                patch("service.jobs._filter_links_for_existing_fields") as filter_links,
                patch("service.jobs.dingtalk_table.mark_links") as mark_links,
                patch("service.jobs.dingtalk_table.update_feedback") as update_feedback,
            ):
                result = run_generate_delivery_tables("record")
            files_were_removed = all(not path.exists() for path in outputs.values())

        self.assertTrue(result["ok"])
        self.assertEqual(result["outputs"], {})
        self.assertEqual(result["emptyPlatforms"], ["美团", "饿了么", "京东"])
        self.assertTrue(files_were_removed)
        upload_file.assert_not_called()
        filter_links.assert_not_called()
        mark_links.assert_called_once_with(configs, "record", {}, "外卖数据已生成")
        self.assertEqual(
            update_feedback.call_args_list[-1].args[2],
            "外卖数据统计完毕，美团、饿了么、京东无数据",
        )

    def test_delivery_failure_keeps_cleared_result_fields_and_does_not_delete_remote_files(self) -> None:
        configs = {
            "dingtalk": {},
            "dingtalk_docs": {},
            "field_mapping": {"fields": {}},
            "report_rules": {},
        }
        record = dingtalk_table.TaskRecord(
            "record",
            {
                "deliveryData": [
                    {"name": "外卖数据.xlsx", "url": "https://alidocs.dingtalk.com/i/nodes/raw-node"}
                ],
                "productMenu": [
                    {"name": "产品清单.xlsx", "url": "https://alidocs.dingtalk.com/i/nodes/menu-node"}
                ],
            },
        )
        with (
            patch("service.jobs.load_configs", return_value=configs),
            patch("service.jobs.dingtalk_table.mark_status"),
            patch("service.jobs.dingtalk_table.fetch_record", return_value=record),
            patch("service.jobs._ensure_local_upload_folder", return_value="shared-folder"),
            patch(
                "service.jobs.dingtalk_docs.folder_id_for_linked_node",
                return_value="shared-folder",
            ),
            patch("service.jobs.dingtalk_table.clear_link_fields") as clear_fields,
            patch("service.jobs._task_delivery_input_dir", side_effect=RuntimeError("download failed")),
            patch("service.jobs.dingtalk_table.mark_failed"),
            patch("service.jobs.dingtalk_table.mark_links") as mark_links,
            patch("service.jobs.dingtalk_docs.delete_existing_file") as delete_file,
        ):
            with self.assertRaisesRegex(RuntimeError, "download failed"):
                run_generate_delivery_tables("record")

        clear_fields.assert_called_once_with(configs, "record", ("meituanData", "elemeData", "jdData"))
        mark_links.assert_not_called()
        delete_file.assert_not_called()


if __name__ == "__main__":
    unittest.main()
