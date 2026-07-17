from __future__ import annotations

import json
import os
import re
from concurrent.futures import ThreadPoolExecutor

from fastapi import FastAPI, HTTPException, Request

from service.jobs import run_schedule_job
from service.schemas import GenerateWeeklyReportResponse


app = FastAPI(title="Competitor Weekly Report Service")
executor = ThreadPoolExecutor(max_workers=2)


def require_secret(secret: str | None) -> None:
    expected = os.getenv("REPORT_SERVICE_SECRET", "").strip()
    if expected and secret != expected:
        raise HTTPException(status_code=401, detail="invalid secret")


def _run_and_log(record_id: str) -> None:
    try:
        result = run_schedule_job(record_id)
        print(f"weekly report generated: {result}")
    except Exception as exc:
        print(f"weekly report failed for {record_id}: {exc}")


async def parse_generate_request(request: Request) -> tuple[str, str | None]:
    raw_body = (await request.body()).decode("utf-8", errors="ignore").strip()
    payload = None
    if raw_body:
        try:
            payload = json.loads(raw_body)
        except json.JSONDecodeError:
            payload = raw_body

    record_id = ""
    secret = request.query_params.get("secret")
    if isinstance(payload, dict):
        for key in ("recordId", "recordID", "record_id", "记录ID", "记录id"):
            value = payload.get(key)
            if value:
                record_id = str(value).strip()
                break
        if payload.get("secret"):
            secret = str(payload.get("secret")).strip()
    elif isinstance(payload, str):
        match = re.search(r"[A-Za-z0-9]{8,}", payload)
        if match:
            record_id = match.group(0)

    if not record_id:
        record_id = str(request.query_params.get("recordId") or request.query_params.get("record_id") or "").strip()
    if not record_id:
        raise HTTPException(status_code=400, detail="missing recordId")
    return record_id, secret


@app.get("/health")
def health() -> dict[str, bool]:
    return {"ok": True}


@app.post("/generate-weekly-report", response_model=GenerateWeeklyReportResponse)
async def generate_weekly_report(request: Request) -> GenerateWeeklyReportResponse:
    record_id, secret = await parse_generate_request(request)
    require_secret(secret)
    executor.submit(_run_and_log, record_id)
    return GenerateWeeklyReportResponse(
        ok=True,
        status="accepted",
        recordId=record_id,
        message="任务已接收，等待1-2分钟。",
    )
