from __future__ import annotations

import json
import unittest
from unittest.mock import AsyncMock, patch

from fastapi import HTTPException
from starlette.requests import Request

from gateway import app as gateway_app
from service import app as service_app


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


def completed_result() -> dict[str, object]:
    return {
        "ok": True,
        "status": "completed",
        "recordId": "record-123",
        "message": "原始文件归档完成",
        "reportUrl": "https://example/report",
        "folderUrl": "https://example/folder",
        "archivedFileCount": 2,
        "archivedFileNames": "a.xlsx、b.xlsx",
    }


class DirectServiceFinalizeTests(unittest.IsolatedAsyncioTestCase):
    async def test_direct_service_returns_completed_result(self) -> None:
        result = completed_result()
        with patch.object(
            service_app.asyncio,
            "to_thread",
            new=AsyncMock(return_value=result),
        ) as to_thread:
            response = await service_app.finalize_report(
                request_with_record_id("record-123")
            )

        self.assertEqual(response.status, "completed")
        self.assertEqual(response.folderUrl, "https://example/folder")
        to_thread.assert_awaited_once_with(service_app.run_finalize_report, "record-123")

    async def test_direct_service_maps_failure_to_http_500(self) -> None:
        with patch.object(
            service_app.asyncio,
            "to_thread",
            new=AsyncMock(side_effect=RuntimeError("原始文件归档失败：回写失败")),
        ):
            with self.assertRaises(HTTPException) as raised:
                await service_app.finalize_report(
                    request_with_record_id("record-123")
                )

        self.assertEqual(raised.exception.status_code, 500)
        self.assertIn("回写失败", raised.exception.detail)


class ProjectGatewayFinalizeTests(unittest.IsolatedAsyncioTestCase):
    async def test_project_gateway_waits_for_completed_result(self) -> None:
        result = completed_result()
        with patch.object(
            gateway_app.asyncio,
            "to_thread",
            new=AsyncMock(return_value=result),
        ) as to_thread:
            response = await gateway_app.finalize_report(
                request_with_record_id("record-123")
            )

        self.assertEqual(response.status, "completed")
        self.assertEqual(response.archivedFileNames, "a.xlsx、b.xlsx")
        to_thread.assert_awaited_once_with(
            gateway_app._run_python_job_result,
            gateway_app.FEEDBACK_ROOT,
            "run_finalize_report",
            "record-123",
        )
