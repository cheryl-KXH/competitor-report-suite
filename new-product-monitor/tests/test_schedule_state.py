from __future__ import annotations

import sys
import unittest
from pathlib import Path
from unittest.mock import patch


ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from service import dingtalk_table  # noqa: E402


def configs() -> dict:
    return {
        "report_rules": {
            "weeklyReportTable": {
                "tableId": "weekly-table",
                "fields": {
                    "reportUrl": "report-url",
                    "business": "business",
                    "startDate": "start-date",
                    "endDate": "end-date",
                    "year": "year",
                    "week": "week",
                    "status": "status",
                    "feedbackMessage": "feedback",
                    "generatedAt": "generated-at",
                },
                "statusOptionIds": {
                    "等待1-2分钟": "running-option",
                    "已生成": "success-option",
                },
            }
        }
    }


class ScheduleStateTests(unittest.TestCase):
    def test_mark_running_clears_old_report_link_before_generation(self) -> None:
        with patch.object(dingtalk_table, "update_schedule_cells") as update_mock:
            dingtalk_table.mark_running(configs(), "record-w29")

        cells = update_mock.call_args.args[2]
        self.assertEqual(cells["report-url"], "")
        self.assertEqual(cells["generated-at"], "")
        self.assertEqual(cells["feedback"], "")
        self.assertEqual(cells["status"]["name"], "等待1-2分钟")

    def test_mark_success_writes_new_report_link_after_upload(self) -> None:
        with patch.object(dingtalk_table, "update_schedule_cells") as update_mock:
            dingtalk_table.mark_success(
                configs(),
                "record-w29",
                "茶坊26W29_20260717",
                "https://example.invalid/new-report",
                "生成记录 3 条",
            )

        cells = update_mock.call_args.args[2]
        self.assertEqual(
            cells["report-url"],
            {
                "text": "茶坊26W29_20260717",
                "link": "https://example.invalid/new-report",
            },
        )
        self.assertEqual(cells["status"]["name"], "已生成")


if __name__ == "__main__":
    unittest.main()
