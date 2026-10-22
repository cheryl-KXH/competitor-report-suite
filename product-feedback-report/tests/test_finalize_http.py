from __future__ import annotations

import json
import unittest
from unittest.mock import patch

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


class DirectServiceFinalizeTests(unittest.IsolatedAsyncioTestCase):
    async def test_direct_service_accepts_background_job(self) -> None:
        with patch.object(service_app, "_submit_finalize", return_value=True) as submit:
            response = await service_app.finalize_report(
                request_with_record_id("record-123")
            )

        self.assertEqual(response.status, "accepted")
        submit.assert_called_once_with("record-123")

    async def test_direct_service_reports_duplicate_as_running(self) -> None:
        with patch.object(service_app, "_submit_finalize", return_value=False):
            response = await service_app.finalize_report(
                request_with_record_id("record-123")
            )

        self.assertEqual(response.status, "running")
        self.assertIn("请勿重复提交", response.message)


class ProjectGatewayFinalizeTests(unittest.IsolatedAsyncioTestCase):
    async def test_project_gateway_accepts_background_job(self) -> None:
        with patch.object(gateway_app, "_submit_finalize", return_value=True) as submit:
            response = await gateway_app.finalize_report(
                request_with_record_id("record-123")
            )

        self.assertEqual(response.status, "accepted")
        submit.assert_called_once_with("record-123")
