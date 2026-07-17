from __future__ import annotations

import json
import os
import re
import subprocess
import sys
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from typing import Any

from fastapi import FastAPI, HTTPException, Request
from pydantic import BaseModel


ROOT = Path(__file__).resolve().parents[1]


def _project_root(env_name: str, suite_name: str, sibling_name: str) -> Path:
    configured = os.getenv(env_name, "").strip()
    if configured:
        return Path(configured).expanduser().resolve()
    if (ROOT / suite_name).exists():
        return ROOT / suite_name
    if ROOT.name == suite_name:
        return ROOT
    return ROOT.parent / sibling_name


WEEKLY_ROOT = _project_root(
    "WEEKLY_REPORT_ROOT",
    "competitor-new-product-monitor_new",
    "competitor-new-product-monitor_new",
)
FEEDBACK_ROOT = _project_root(
    "FEEDBACK_REPORT_ROOT",
    "product-feedback-report",
    "product-feedback-report",
)


class AcceptedResponse(BaseModel):
    ok: bool
    status: str
    recordId: str
    message: str = ""


app = FastAPI(title="Competitor Report Gateway")
executor = ThreadPoolExecutor(max_workers=3)


def require_secret(secret: str | None) -> None:
    expected = os.getenv("REPORT_SERVICE_SECRET", "").strip()
    if expected and secret != expected:
        raise HTTPException(status_code=401, detail="invalid secret")


def _pick_record_id(payload: Any, query: dict[str, str]) -> tuple[str, str | None]:
    record_id = ""
    secret = query.get("secret")
    if isinstance(payload, dict):
        for key in ("recordId", "recordID", "record_id", "记录ID", "记录id"):
            value = payload.get(key)
            if value:
                record_id = str(value).strip()
                break
        if payload.get("secret"):
            secret = str(payload["secret"]).strip()
    elif isinstance(payload, str):
        match = re.search(r"[A-Za-z0-9_-]{8,}", payload)
        if match:
            record_id = match.group(0)
    if not record_id:
        record_id = str(query.get("recordId") or query.get("record_id") or query.get("recordID") or "").strip()
    return record_id, secret


async def parse_request(request: Request) -> tuple[str, str | None]:
    raw_body = (await request.body()).decode("utf-8", errors="ignore").strip()
    payload: Any = None
    if raw_body:
        try:
            payload = json.loads(raw_body)
        except json.JSONDecodeError:
            payload = raw_body
    record_id, secret = _pick_record_id(payload, dict(request.query_params))
    if not record_id:
        raise HTTPException(status_code=400, detail="missing recordId")
    return record_id, secret


def _run_python_job(project_root: Path, function_name: str, record_id: str) -> None:
    code = f"from service.jobs import {function_name}; print({function_name}({record_id!r}))"
    result = subprocess.run(
        [sys.executable, "-c", code],
        cwd=project_root,
        check=False,
        capture_output=True,
        text=True,
    )
    if result.stdout:
        print(result.stdout)
    if result.returncode != 0:
        print(result.stderr or result.stdout)


def _submit(project_root: Path, function_name: str, record_id: str) -> None:
    executor.submit(_run_python_job, project_root, function_name, record_id)


@app.get("/health")
def health() -> dict[str, Any]:
    return {
        "ok": True,
        "weeklyRoot": str(WEEKLY_ROOT),
        "feedbackRoot": str(FEEDBACK_ROOT),
    }


@app.post("/generate-weekly-report", response_model=AcceptedResponse)
async def generate_weekly_report(request: Request) -> AcceptedResponse:
    record_id, secret = await parse_request(request)
    require_secret(secret)
    _submit(WEEKLY_ROOT, "run_schedule_job", record_id)
    return AcceptedResponse(ok=True, status="accepted", recordId=record_id, message="周报任务已接收。")


@app.post("/prepare-product-menu", response_model=AcceptedResponse)
async def prepare_product_menu(request: Request) -> AcceptedResponse:
    record_id, secret = await parse_request(request)
    require_secret(secret)
    _submit(FEEDBACK_ROOT, "run_prepare_product_menu", record_id)
    return AcceptedResponse(ok=True, status="accepted", recordId=record_id, message="产品清单任务已接收。")


@app.post("/generate-data-tables", response_model=AcceptedResponse)
async def generate_data_tables(request: Request) -> AcceptedResponse:
    record_id, secret = await parse_request(request)
    require_secret(secret)
    _submit(FEEDBACK_ROOT, "run_generate_data_tables", record_id)
    return AcceptedResponse(ok=True, status="accepted", recordId=record_id, message="数表生成任务已接收。")


@app.post("/generate-report", response_model=AcceptedResponse)
async def generate_report(request: Request) -> AcceptedResponse:
    record_id, secret = await parse_request(request)
    require_secret(secret)
    _submit(FEEDBACK_ROOT, "run_generate_report", record_id)
    return AcceptedResponse(ok=True, status="accepted", recordId=record_id, message="报告生成任务已接收。")

