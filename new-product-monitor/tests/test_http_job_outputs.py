from __future__ import annotations

import sys
import tempfile
import unittest
from datetime import date
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch


ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from service import jobs  # noqa: E402


class HttpJobOutputTests(unittest.TestCase):
    def test_command_line_default_output_stays_in_project_outputs(self) -> None:
        output_root = jobs.weekly_outputs.resolve_output_root(None, {"outputDirectory": "outputs"})

        self.assertEqual(output_root, ROOT / "outputs")

    def test_http_job_uploads_temporary_outputs_then_removes_them(self) -> None:
        schedule = SimpleNamespace(
            business="茶坊",
            year=2026,
            week="W29",
            start=date(2026, 7, 11),
            end=date(2026, 7, 17),
        )
        captured: dict[str, Path] = {}

        def generate_outputs(**kwargs):
            output_root = Path(kwargs["output_dir"])
            self.assertNotEqual(output_root, ROOT / "outputs")
            report_dir = output_root / kwargs["stem"]
            report_dir.mkdir(parents=True)
            output_paths = []
            for suffix in (".xlsx", ".html", ".png"):
                path = report_dir / f"{kwargs['stem']}{suffix}"
                path.write_bytes(b"test")
                output_paths.append(path)
            captured["output_root"] = output_root
            return SimpleNamespace(
                report_dir=report_dir,
                output_paths=output_paths,
                record_count=3,
                missing_field_count=0,
                image_issue_count=0,
                data_quality_warnings=[],
            )

        def upload_outputs(_config, _business, _year, _stem, output_paths, **_kwargs):
            self.assertTrue(all(path.exists() for path in output_paths))
            return "https://example.invalid/report"

        with tempfile.TemporaryDirectory() as temp_parent:
            with patch.object(jobs.tempfile, "tempdir", temp_parent):
                with patch.object(jobs, "load_configs", return_value={"report_rules": {}}):
                    with patch.object(jobs.dingtalk_table, "mark_running"):
                        with patch.object(jobs.dingtalk_table, "fetch_schedule_row", return_value=schedule):
                            with patch.object(jobs.excel_report, "parse_business", return_value="茶坊"):
                                with patch.object(jobs.weekly_outputs, "generate_weekly_outputs", side_effect=generate_outputs):
                                    with patch.object(jobs.dingtalk_docs, "load_docs_config", return_value={}):
                                        with patch.object(jobs.dingtalk_docs, "upload_report_directory", side_effect=upload_outputs):
                                            with patch.object(jobs.dingtalk_table, "mark_success"):
                                                result = jobs.run_schedule_job("record-w29")

            self.assertFalse(captured["output_root"].exists())

        self.assertIsNone(result["reportDir"])
        self.assertEqual(
            result["outputFiles"],
            [
                "茶坊26W29_20260717.xlsx",
                "茶坊26W29_20260717.html",
                "茶坊26W29_20260717.png",
            ],
        )
        self.assertEqual(result["recordCount"], 3)


if __name__ == "__main__":
    unittest.main()
