from __future__ import annotations

import json
import sys
import unittest
from datetime import date
from pathlib import Path
from subprocess import CompletedProcess
from unittest.mock import patch


ROOT = Path(__file__).resolve().parents[1]
SCRIPTS_DIR = ROOT / "scripts"
if str(SCRIPTS_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPTS_DIR))

import generate_weekly_report as weekly_report  # noqa: E402
import render_report_image as report_image  # noqa: E402


def command_result(payload: dict, returncode: int = 0) -> CompletedProcess[str]:
    return CompletedProcess(
        args=["mcporter"],
        returncode=returncode,
        stdout=json.dumps(payload, ensure_ascii=False),
        stderr="",
    )


def timeout_result() -> CompletedProcess[str]:
    return command_result(
        {
            "success": False,
            "error": "请求超时",
            "message": "请求超时。服务响应较慢，请稍后重试",
            "code": "TIMEOUT_ERROR",
            "retryable": True,
        }
    )


def nested_calc_not_ready_result() -> CompletedProcess[str]:
    return command_result(
        {
            "success": False,
            "error": {
                "retryable": True,
                "code": "InvalidRequest.DataNotReady",
                "type": "SYSTEM_ERROR",
                "message": "Calc fields are not ready for current version",
            },
        }
    )


class DingTalkQueryRetryTests(unittest.TestCase):
    def test_query_retries_json_containing_replacement_character(self) -> None:
        corrupted = command_result(
            {
                "status": "success",
                "data": {
                    "records": [
                        {
                            "recordId": "rec-1",
                            "cells": {"sellingPoint": "Q弹\ufffd\ufffd西米"},
                        }
                    ]
                },
            }
        )
        success = command_result(
            {
                "status": "success",
                "data": {
                    "records": [
                        {
                            "recordId": "rec-1",
                            "cells": {"sellingPoint": "Q弹小西米"},
                        }
                    ]
                },
            }
        )
        with patch.object(
            weekly_report.subprocess,
            "run",
            side_effect=[corrupted, success],
        ) as run_mock:
            with patch.object(weekly_report.time, "sleep") as sleep_mock:
                result = weekly_report.call_dingtalk_tool(
                    {}, "query_records", {"limit": 20}
                )

        self.assertEqual(
            result["data"]["records"][0]["cells"]["sellingPoint"],
            "Q弹小西米",
        )
        self.assertEqual(run_mock.call_count, 2)
        sleep_mock.assert_not_called()
        repair_args = json.loads(run_mock.call_args.args[0][-1])
        self.assertEqual(repair_args["recordIds"], ["rec-1"])

    def test_query_retries_truncated_json_and_uses_safe_mcporter_environment(self) -> None:
        truncated = CompletedProcess(
            args=[],
            returncode=0,
            stdout='{"status":"success","data":{"records":[',
            stderr="",
        )
        success = command_result(
            {
                "status": "success",
                "data": {"records": [{"recordId": "rec-1", "cells": {}}]},
            }
        )
        with patch.object(
            weekly_report.subprocess,
            "run",
            side_effect=[truncated, success],
        ) as run_mock:
            with patch.object(weekly_report.time, "sleep") as sleep_mock:
                result = weekly_report.call_dingtalk_tool(
                    {}, "query_records", {"limit": 20}
                )

        self.assertEqual(result["data"]["records"][0]["recordId"], "rec-1")
        self.assertEqual(run_mock.call_count, 2)
        self.assertEqual(run_mock.call_args.kwargs["encoding"], "utf-8")
        self.assertEqual(run_mock.call_args.kwargs["errors"], "replace")
        self.assertEqual(
            run_mock.call_args.kwargs["env"]["MCPORTER_NO_FORCE_EXIT"], "1"
        )
        sleep_mock.assert_called_once_with(1.0)

    def test_query_retries_retryable_failure_then_returns_records(self) -> None:
        success = command_result(
            {
                "status": "success",
                "data": {"records": [{"recordId": "rec-1", "cells": {}}]},
            }
        )
        with patch.object(weekly_report.subprocess, "run", side_effect=[timeout_result(), success]) as run_mock:
            with patch.object(weekly_report.time, "sleep") as sleep_mock:
                result = weekly_report.call_dingtalk_tool({}, "query_records", {"limit": 20})

        self.assertEqual(result["data"]["records"][0]["recordId"], "rec-1")
        self.assertEqual(run_mock.call_count, 2)
        sleep_mock.assert_called_once_with(1.0)

    def test_query_raises_after_three_retryable_failures(self) -> None:
        with patch.object(
            weekly_report.subprocess,
            "run",
            side_effect=[timeout_result(), timeout_result(), timeout_result()],
        ) as run_mock:
            with patch.object(weekly_report.time, "sleep") as sleep_mock:
                with self.assertRaisesRegex(RuntimeError, "TIMEOUT_ERROR"):
                    weekly_report.call_dingtalk_tool({}, "query_records", {"limit": 20})

        self.assertEqual(run_mock.call_count, 3)
        self.assertEqual(sleep_mock.call_count, 2)

    def test_query_retries_nested_calc_not_ready_error_with_backoff(self) -> None:
        success = command_result(
            {
                "status": "success",
                "data": {"records": [{"recordId": "rec-1", "cells": {}}]},
            }
        )
        with patch.object(
            weekly_report.subprocess,
            "run",
            side_effect=[
                nested_calc_not_ready_result(),
                nested_calc_not_ready_result(),
                nested_calc_not_ready_result(),
                success,
            ],
        ) as run_mock:
            with patch.object(weekly_report.time, "sleep") as sleep_mock:
                result = weekly_report.call_dingtalk_tool({}, "query_records", {"limit": 20})

        self.assertEqual(result["data"]["records"][0]["recordId"], "rec-1")
        self.assertEqual(run_mock.call_count, 4)
        self.assertEqual([call.args[0] for call in sleep_mock.call_args_list], [1.0, 2.0, 4.0])

    def test_query_reports_nested_error_fields_after_calc_retry_limit(self) -> None:
        with patch.object(
            weekly_report.subprocess,
            "run",
            side_effect=[nested_calc_not_ready_result()] * 6,
        ) as run_mock:
            with patch.object(weekly_report.time, "sleep") as sleep_mock:
                with self.assertRaisesRegex(
                    RuntimeError,
                    "InvalidRequest.DataNotReady: Calc fields are not ready",
                ):
                    weekly_report.call_dingtalk_tool({}, "query_records", {"limit": 20})

        self.assertEqual(run_mock.call_count, 6)
        self.assertEqual(
            [call.args[0] for call in sleep_mock.call_args_list],
            [1.0, 2.0, 4.0, 8.0, 10.0],
        )

    def test_second_page_timeout_does_not_return_partial_records(self) -> None:
        first_page = command_result(
            {
                "status": "success",
                "data": {
                    "records": [{"recordId": "rec-1", "cells": {}}],
                    "nextCursor": "cursor-1",
                },
            }
        )
        with patch.object(
            weekly_report.subprocess,
            "run",
            side_effect=[first_page, timeout_result(), timeout_result(), timeout_result()],
        ):
            with patch.object(weekly_report.time, "sleep"):
                with self.assertRaisesRegex(RuntimeError, "TIMEOUT_ERROR"):
                    weekly_report.query_records(
                        {"dingtalk": {"baseId": "base", "tableId": "table"}},
                        {"launchDate": "date-field"},
                        date(2026, 7, 11),
                        date(2026, 7, 17),
                    )

    def test_successful_empty_query_remains_valid(self) -> None:
        empty = command_result({"status": "success", "data": {"records": []}})
        with patch.object(weekly_report.subprocess, "run", return_value=empty) as run_mock:
            result = weekly_report.call_dingtalk_tool({}, "query_records", {"limit": 20})

        self.assertEqual(result["data"]["records"], [])
        run_mock.assert_called_once()

    def test_idempotent_update_retries_timeout_then_succeeds(self) -> None:
        success = command_result({"status": "success", "data": {}})
        with patch.object(weekly_report.subprocess, "run", side_effect=[timeout_result(), success]) as run_mock:
            with patch.object(weekly_report.time, "sleep") as sleep_mock:
                result = weekly_report.call_dingtalk_tool(
                    {},
                    "update_records",
                    {"records": [{"recordId": "rec-1", "cells": {"status": "等待1-2分钟"}}]},
                )

        self.assertEqual(result["status"], "success")
        self.assertEqual(run_mock.call_count, 2)
        sleep_mock.assert_called_once_with(1.0)

    def test_idempotent_update_raises_after_three_timeouts(self) -> None:
        with patch.object(
            weekly_report.subprocess,
            "run",
            side_effect=[timeout_result(), timeout_result(), timeout_result()],
        ) as run_mock:
            with patch.object(weekly_report.time, "sleep") as sleep_mock:
                with self.assertRaisesRegex(RuntimeError, "TIMEOUT_ERROR"):
                    weekly_report.call_dingtalk_tool(
                        {},
                        "update_records",
                        {"records": [{"recordId": "rec-1", "cells": {"status": "等待1-2分钟"}}]},
                    )

        self.assertEqual(run_mock.call_count, 3)
        self.assertEqual(sleep_mock.call_count, 2)

    def test_non_idempotent_tool_is_not_retried(self) -> None:
        with patch.object(weekly_report.subprocess, "run", return_value=timeout_result()) as run_mock:
            with self.assertRaisesRegex(RuntimeError, "TIMEOUT_ERROR"):
                weekly_report.call_dingtalk_tool({}, "create_records", {"records": []})

        run_mock.assert_called_once()


class ResponsiveHtmlTests(unittest.TestCase):
    def test_long_screenshot_is_kept_below_preview_texture_limit(self) -> None:
        from PIL import Image

        source = Image.new("RGB", (1344, 20000), "white")
        constrained = report_image.constrain_output_image_height(source, 16000)

        self.assertEqual(constrained.size, (1075, 16000))
        source.close()
        constrained.close()

    def test_short_screenshot_is_not_resampled(self) -> None:
        from PIL import Image

        source = Image.new("RGB", (1344, 12000), "white")
        constrained = report_image.constrain_output_image_height(source, 16000)

        self.assertIs(constrained, source)
        source.close()

    def test_screenshot_chunks_verify_each_scroll_position(self) -> None:
        source = Path(report_image.__file__).read_text(encoding="utf-8")

        self.assertIn("window.scrollTo(0, targetY)", source)
        self.assertIn("周报截图滚动位置异常", source)

    def test_price_without_slash_stays_on_one_line(self) -> None:
        rendered = report_image.render_price_html("18.48元(21元)", True)

        self.assertNotIn("<br>", rendered)
        self.assertNotIn("<wbr>", rendered)
        self.assertIn('<span class="strike">21元</span>', rendered)
        self.assertEqual(
            report_image.safe_price_width_segments("18.48元(21元)"),
            ["18.48元(21元)"],
        )
        self.assertEqual(
            report_image.estimate_price_lines("18.48元(21元)", 72, 12),
            1,
        )

    def test_price_with_slash_keeps_explicit_slash_break(self) -> None:
        rendered = report_image.render_price_html(
            "12.64元(16元)(中杯)/15.01元(19元)(大杯)", True
        )

        self.assertIn("/<br>", rendered)

    def test_html_uses_width_only_breakpoints_and_mobile_vertical_pan(self) -> None:
        source = Path(report_image.__file__).read_text(encoding="utf-8")

        self.assertIn("@media (min-width: 768px)", source)
        self.assertIn("@media (max-width: 767px)", source)
        self.assertIn("touch-action: pan-y pinch-zoom", source)
        self.assertIn("panX = dragStartPanX + dx / currentScale", source)
        self.assertIn("th, td {{ border-width: 1.5px; }}", source)
        self.assertIn("window.matchMedia('(max-width: 767px)')", source)
        self.assertNotIn("touch-action: none", source)
        self.assertNotIn("@media (hover:", source)

    def test_summary_rows_use_one_compact_minimum_height(self) -> None:
        layout = report_image.load_image_layout()
        record = weekly_report.ReportRecord(
            record_id="record",
            brand="茶百道",
            product_name="虎斑橘拿铁",
            launch_date=date(2026, 7, 30),
            price="12.8元(21元)",
            category="咖啡",
            series="",
            selling_point="",
            ingredients="",
            image_urls=[],
            remark="/",
        )

        summary = report_image.build_summary_html([record], ["茶百道"], layout)
        source = Path(report_image.__file__).read_text(encoding="utf-8")

        self.assertIn("<tbody>\n<tr>", summary)
        self.assertNotIn("line-count-", summary)
        self.assertIn(
            ".summary-table tbody > tr > td {{ height: {summary_row_height}px; }}",
            source,
        )
        self.assertNotIn(".summary-table tr.line-count-", source)


class RemarkRuleTests(unittest.TestCase):
    def setUp(self) -> None:
        self.rules = weekly_report.load_configs()["report_rules"]

    def test_plain_launch_is_not_shown_in_remark(self) -> None:
        self.assertEqual(
            weekly_report.compute_remark({"returnStatus": ["上新"]}, self.rules),
            "/",
        )

    def test_each_visible_launch_type_is_shown(self) -> None:
        for label in ("回归", "经典回归"):
            with self.subTest(label=label):
                self.assertEqual(
                    weekly_report.compute_remark(
                        {"returnStatus": [label]}, self.rules
                    ),
                    label,
                )

    def test_series_launch_is_hidden_and_classic_return_suppresses_return(self) -> None:
        self.assertEqual(
            weekly_report.compute_remark(
                {
                    "returnStatus": ["上新", "系列上新", "回归", "经典回归"],
                    "collaboration": "品牌IP",
                    "remark": "限时",
                },
                self.rules,
            ),
            "经典回归，品牌IP联名，限时",
        )

    def test_series_launch_is_hidden_when_return_is_present(self) -> None:
        self.assertEqual(
            weekly_report.compute_remark(
                {"returnStatus": ["系列上新", "回归"]}, self.rules
            ),
            "回归",
        )

    def test_series_launch_alone_is_not_shown(self) -> None:
        self.assertEqual(
            weekly_report.compute_remark(
                {"returnStatus": ["系列上新"]}, self.rules
            ),
            "/",
        )


if __name__ == "__main__":
    unittest.main()
