from __future__ import annotations

import json
import subprocess
import unittest
from unittest.mock import AsyncMock, patch

from fastapi import HTTPException
from starlette.requests import Request

from gateway import app as gateway_app


def request_with_record_id(record_id: str) -> Request:
    body = json.dumps({"recordId": record_id}).encode("utf-8")

    async def receive() -> dict[str, object]:
        return {"type": "http.request", "body": body, "more_body": False}

    return Request(
        {
            "type": "http",
            "asgi": {"version": "3.0"},
            "http_version": "1.1",
            "method": "POST",
            "scheme": "http",
            "path": "/finalize-report",
            "raw_path": b"/finalize-report",
            "query_string": b"",
            "headers": [(b"content-type", b"application/json")],
            "client": ("test", 1),
            "server": ("test", 80),
        },
        receive,
    )


class GatewayRootTests(unittest.TestCase):
    def test_weekly_root_uses_current_suite_project(self) -> None:
        self.assertEqual(gateway_app.WEEKLY_ROOT.name, "new-product-monitor")
        self.assertTrue(gateway_app.WEEKLY_ROOT.is_dir())

    def test_sync_job_runner_extracts_structured_result(self) -> None:
        payload = {
            "ok": True,
            "status": "completed",
            "recordId": "record-123",
        }
        completed = subprocess.CompletedProcess(
            args=[],
            returncode=0,
            stdout=(
                "progress output\n"
                + "__REPORT_JOB_RESULT__"
                + json.dumps(payload)
                + "\n"
            ),
            stderr="",
        )
        with patch.object(gateway_app.subprocess, "run", return_value=completed):
            result = gateway_app._run_python_job_result(
                gateway_app.FEEDBACK_ROOT,
                "run_finalize_report",
                "record-123",
            )

        self.assertEqual(result, payload)


class GatewayFinalizeTests(unittest.IsolatedAsyncioTestCase):
    async def test_finalize_waits_for_completed_job_result(self) -> None:
        result = {
            "ok": True,
            "status": "completed",
            "recordId": "record-123",
            "message": "原始文件归档完成",
            "reportUrl": "https://example/report",
            "folderUrl": "https://example/folder",
            "archivedFileCount": 2,
            "archivedFileNames": "a.xlsx、b.xlsx",
        }
        with patch.object(
            gateway_app.asyncio, "to_thread", new=AsyncMock(return_value=result)
        ) as to_thread:
            response = await gateway_app.finalize_report(
                request_with_record_id("record-123")
            )

        self.assertTrue(response.ok)
        self.assertEqual(response.status, "completed")
        self.assertEqual(response.archivedFileCount, 2)
        to_thread.assert_awaited_once_with(
            gateway_app._run_python_job_result,
            gateway_app.FEEDBACK_ROOT,
            "run_finalize_report",
            "record-123",
        )

    async def test_finalize_returns_http_500_when_job_fails(self) -> None:
        with patch.object(
            gateway_app.asyncio,
            "to_thread",
            new=AsyncMock(side_effect=RuntimeError("原始文件归档失败：上传失败")),
        ):
            with self.assertRaises(HTTPException) as raised:
                await gateway_app.finalize_report(
                    request_with_record_id("record-123")
                )

        self.assertEqual(raised.exception.status_code, 500)
        self.assertIn("上传失败", raised.exception.detail)


if __name__ == "__main__":
    unittest.main()
