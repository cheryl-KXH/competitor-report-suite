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


class DingTalkQueryRetryTests(unittest.TestCase):
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


if __name__ == "__main__":
    unittest.main()
