from __future__ import annotations

from datetime import date
from pathlib import Path
import tempfile
import unittest
from unittest.mock import ANY, patch

from openpyxl import load_workbook

from scripts.social.processing import (
    PlatformSummary,
    build_social_feedback_workbook,
    summarize_social_rows,
)
from service import dingtalk_table
from service.jobs import (
    _consumer_feedback_configs,
    _ensure_social_archive_folder,
    _prepare_social_input,
    _wait_for_ai_table_attachment_url,
    run_generate_consumer_feedback_tables,
)


class SocialFeedbackStatisticsTests(unittest.TestCase):
    def test_tags_can_count_on_both_sides_while_users_follow_sentiment(self) -> None:
        rows = [
            ["主贴id", "情感识别", "评价1-好/差评", "评价1-对应标签", "评价2-好/差评", "评价2-对应标签"],
            ["p1", "正向", "好评", "好喝", "差评", "太甜"],
            ["p2", "负向", "差评", "太甜", "", ""],
            ["p3", "中性", "好评", "清爽", "", ""],
        ]
        summary = summarize_social_rows("weibo", "微博", rows)
        self.assertEqual(summary.positive_tags, (("好喝", 1), ("清爽", 1)))
        self.assertEqual(summary.negative_tags, (("太甜", 2),))
        self.assertEqual(summary.positive_users, 1)
        self.assertEqual(summary.negative_users, 1)

    def test_missing_required_columns_is_explicit(self) -> None:
        with self.assertRaisesRegex(RuntimeError, "情感识别"):
            summarize_social_rows("weibo", "微博", [["主贴id"], ["p1"]])

    def test_real_export_shape_supports_five_evaluation_groups(self) -> None:
        header = ["品牌", "主贴id", "产品名称", "搜索关键词", "用户名", "内容", "发布时间", "链接", "情感识别"]
        for index in range(1, 6):
            header.extend([f"评价{index}-好/差评", f"评价{index}-对应标签"])
        row = ["茉莉奶白", "p1", "青芒黄皮冰茶", "关键词", "用户", "内容", "2026-07-09", "链接", "正向"]
        for index in range(1, 6):
            row.extend(["好评", f"标签{index}"])
        summary = summarize_social_rows("douyin", "抖音", [header, row])
        self.assertEqual(sum(count for _, count in summary.positive_tags), 5)
        self.assertEqual(summary.positive_users, 1)

    def test_workbook_matches_single_product_summary_layout(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            output = build_social_feedback_workbook(
                brand="测试品牌",
                product="测试新品",
                start_date="6.10",
                end_date="7.9",
                summaries=[
                    PlatformSummary("weibo", "微博", (("好喝", 2),), (("太甜", 1),), 3, 1),
                    PlatformSummary("xiaohongshu", "小红书", (), (), 0, 0),
                    PlatformSummary("douyin", "抖音", (("清爽", 1),), (), 1, 0),
                ],
                output_dir=Path(temporary),
            )
            workbook = load_workbook(output, data_only=False)
            try:
                sheet = workbook["测试新品"]
                self.assertEqual(sheet["A1"].value, "测试新品 6.10-7.9 第三方平台评价反馈")
                self.assertEqual(sheet["F1"].value, 0.8)
                self.assertEqual([sheet[f"F{row}"].value for row in range(2, 5)], [5, 4, 1])
                self.assertNotEqual(sheet["F1"].data_type, "f")
                self.assertIn("小红书好评", [cell.value for cell in sheet["A"]])
                self.assertIn("/", [cell.value for cell in sheet["A"]])
                self.assertIn("A1:D1", [str(item) for item in sheet.merged_cells.ranges])
                self.assertIn("$A$1:$F$", str(sheet.print_area))
                self.assertEqual(sheet.page_setup.fitToWidth, 1)
                self.assertEqual(sheet.page_setup.fitToHeight, 0)
                self.assertEqual(sheet.column_dimensions["A"].width, 27)
            finally:
                workbook.close()


class SocialFeedbackJobTests(unittest.TestCase):
    def test_consumer_table_mapping_uses_current_social_fields(self) -> None:
        configs = _consumer_feedback_configs({"dingtalk": {"tableId": "old"}})
        fields = configs["field_mapping"]["fields"]
        self.assertEqual(configs["dingtalk"]["tableId"], "hvcu7Bw")
        self.assertEqual(fields["launchDate"]["fieldId"], "mKUEya0")
        self.assertEqual(fields["day30"]["fieldId"], "SSrz1N1")
        self.assertEqual(fields["report"]["fieldId"], "amWTton")
        self.assertEqual(fields["feedback"]["fieldId"], "j8IgB7P")
        self.assertNotIn("allData", fields)

    def test_local_attachment_url_is_polled_from_ai_table(self) -> None:
        initial = dingtalk_table.TaskRecord(
            "record",
            {"weibo": [{"filename": "微博.xlsx", "resourceId": "r1"}]},
        )
        refreshed = dingtalk_table.TaskRecord(
            "record",
            {
                "weibo": [
                    {
                        "filename": "微博.xlsx",
                        "resourceId": "r1",
                        "resourceUrl": "https://example.test/weibo.xlsx",
                    }
                ]
            },
        )
        with patch("service.jobs.dingtalk_table.fetch_record", return_value=refreshed) as fetch:
            result = _wait_for_ai_table_attachment_url(
                {}, initial, "weibo", "微博", attempts=2, delay_seconds=0
            )
        self.assertEqual(result, ("微博.xlsx", "https://example.test/weibo.xlsx"))
        fetch.assert_called_once_with({}, "record")

    def test_social_input_never_searches_service_machine_paths(self) -> None:
        record = dingtalk_table.TaskRecord("record", {"weibo": "/tmp/微博.xlsx"})
        with patch("service.jobs._single_local_xlsx") as local_lookup:
            with self.assertRaisesRegex(RuntimeError, "只支持钉钉文档节点"):
                _prepare_social_input({}, record, "weibo", "微博", None)
        local_lookup.assert_not_called()

    def test_social_archive_reuses_folder_containing_brand_and_product(self) -> None:
        configs = {"dingtalk_docs": {}}
        record = dingtalk_table.TaskRecord(
            "record",
            {"brand": {"name": "茉莉奶白"}, "productName": "青芒黄皮冰茶"},
        )
        with (
            patch(
                "service.jobs.dingtalk_docs.ensure_local_upload_month_folder",
                return_value=("month", "https://example/month"),
            ) as ensure_month,
            patch(
                "service.jobs.dingtalk_docs.child_folders",
                return_value=[
                    ("matched", "茉莉奶白：青芒黄皮冰茶、青芒香柚橄榄"),
                    ("other", "瑞幸：其他新品"),
                ],
            ),
            patch("service.jobs.dingtalk_table.fetch_records") as fetch_records,
        ):
            folder_id = _ensure_social_archive_folder(
                configs, record, date(2026, 7, 9)
            )

        self.assertEqual(folder_id, "matched")
        ensure_month.assert_called_once_with({}, 2026, 7)
        fetch_records.assert_not_called()

    def test_social_archive_groups_same_brand_and_same_day30_products(self) -> None:
        configs = {"dingtalk": {}, "dingtalk_docs": {}}
        current = dingtalk_table.TaskRecord(
            "current",
            {
                "brand": {"name": "茉莉奶白"},
                "productName": "青芒黄皮冰茶",
                "day30": date(2026, 7, 9),
            },
        )
        records = [
            current,
            dingtalk_table.TaskRecord(
                "same",
                {
                    "brand": {"name": "茉莉奶白"},
                    "productName": "青芒香柚橄榄",
                    "day30": date(2026, 7, 9),
                },
            ),
            dingtalk_table.TaskRecord(
                "other-day",
                {
                    "brand": {"name": "茉莉奶白"},
                    "productName": "不应进入",
                    "day30": date(2026, 7, 10),
                },
            ),
        ]
        with (
            patch(
                "service.jobs.dingtalk_docs.ensure_local_upload_month_folder",
                return_value=("month", "https://example/month"),
            ),
            patch("service.jobs.dingtalk_docs.child_folders", return_value=[]),
            patch("service.jobs.dingtalk_table.fetch_records", return_value=records),
            patch(
                "service.jobs.dingtalk_docs.ensure_child_folder",
                return_value=("created", "https://example/created"),
            ) as ensure_child,
        ):
            folder_id = _ensure_social_archive_folder(
                configs, current, date(2026, 7, 9)
            )

        self.assertEqual(folder_id, "created")
        ensure_child.assert_called_once_with(
            {}, "month", "茉莉奶白：青芒黄皮冰茶、青芒香柚橄榄"
        )

    def test_fetch_records_reads_all_ai_table_pages(self) -> None:
        configs = _consumer_feedback_configs(
            {"dingtalk": {"baseId": "base", "tableId": "ignored"}}
        )
        brand_field = configs["field_mapping"]["fields"]["brand"]["fieldId"]
        with patch(
            "service.dingtalk_table.call_table_tool",
            side_effect=[
                {
                    "data": {
                        "records": [
                            {"recordId": "r1", "cells": {brand_field: {"name": "茉莉奶白"}}}
                        ],
                        "nextCursor": "next",
                    }
                },
                {
                    "data": {
                        "records": [
                            {"recordId": "r2", "cells": {brand_field: {"name": "瑞幸"}}}
                        ]
                    }
                },
            ],
        ) as query:
            records = dingtalk_table.fetch_records(configs)

        self.assertEqual([record.record_id for record in records], ["r1", "r2"])
        self.assertEqual(records[0].cells["brand"], {"name": "茉莉奶白"})
        self.assertEqual(query.call_args_list[0].args[2]["limit"], 100)
        self.assertNotIn("cursor", query.call_args_list[0].args[2])
        self.assertEqual(query.call_args_list[1].args[2]["cursor"], "next")

    def test_job_clears_report_runs_three_steps_and_writes_url(self) -> None:
        record = dingtalk_table.TaskRecord(
            "record",
            {
                "brand": {"name": "茉莉奶白"},
                "productName": "青芒黄皮冰茶",
                "launchDate": date(2026, 6, 10),
                "day30": date(2026, 7, 9),
                "weibo": [{"url": "https://alidocs.dingtalk.com/i/nodes/weibo"}],
                "xiaohongshu": None,
                "douyin": None,
            },
        )
        output = Path("/tmp/茉莉奶白-青芒黄皮冰茶-社媒评论统计.xlsx")
        prepare_results = iter(
            [("folder", Path("/tmp/weibo.xlsx")), ("folder", None), ("folder", None)]
        )
        with (
            patch("service.jobs.load_configs", return_value={"dingtalk": {}, "dingtalk_docs": {}}),
            patch("service.jobs.dingtalk_table.fetch_record", return_value=record),
            patch("service.jobs.dingtalk_table.clear_link_fields") as clear,
            patch("service.jobs.dingtalk_table.update_feedback") as feedback,
            patch("service.jobs._ensure_social_archive_folder", return_value="folder"),
            patch("service.jobs._prepare_social_input", side_effect=lambda *args: next(prepare_results)),
            patch("service.jobs.generate_consumer_feedback_tables", return_value=output) as generate,
            patch("service.jobs.dingtalk_docs.upload_file", return_value="https://docs/result") as upload,
            patch("service.jobs.dingtalk_table.mark_links") as mark_links,
        ):
            result = run_generate_consumer_feedback_tables("record")

        self.assertTrue(result["ok"])
        clear.assert_called_once()
        generate.assert_called_once()
        upload.assert_called_once_with({}, output, "folder")
        mark_links.assert_called_once_with(
            ANY,
            "record",
            {"report": (output.name, "https://docs/result")},
        )
        self.assertEqual(
            [item.args[2] for item in feedback.call_args_list],
            [
                "1/3 正在下载并归档社媒数据",
                "2/3 正在生成社媒评论统计表",
                "3/3 正在上传统计表到钉钉文档",
                "已生成社媒评论统计表",
            ],
        )


if __name__ == "__main__":
    unittest.main()
