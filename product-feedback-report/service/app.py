from __future__ import annotations

import json
import os
import re
from concurrent.futures import ThreadPoolExecutor
from typing import Any

from fastapi import FastAPI, HTTPException, Request

from service.jobs import (
    run_finalize_report,
    run_generate_consumer_feedback_tables,
    run_generate_delivery_tables,
    run_generate_report,
    run_prepare_product_menu,
)
from service.schemas import AcceptedResponse, HealthResponse


app = FastAPI(title="Product Feedback Report Service")
executor = ThreadPoolExecutor(max_workers=2)


def require_secret(secret: str | None) -> None:
    expected = os.getenv("REPORT_SERVICE_SECRET", "").strip()
    if expected and secret != expected:
        raise HTTPException(status_code=401, detail="invalid secret")


def _pick_record_id(payload: Any, query: dict[str, str]) -> tuple[str, str | None]:
    record_id = ""
    secret = query.get("secret")
    if isinstance(payload, dict):
        for key in ("recordId", "recordID", "record_id", "记录ID", "记录id"):
            if payload.get(key):
                record_id = str(payload[key]).strip()
                break
        secret = str(payload.get("secret") or secret or "").strip() or None
    elif isinstance(payload, str):
        match = re.search(r"[A-Za-z0-9_-]{8,}", payload)
        if match:
            record_id = match.group(0)
    if not record_id:
        for key in ("recordId", "recordID", "record_id"):
            if query.get(key):
                record_id = str(query[key]).strip()
                break
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


def _submit(fn, record_id: str) -> None:
    def run() -> None:
        try:
            print(fn(record_id))
        except Exception as exc:
            print(f"job failed for {record_id}: {type(exc).__name__}: {exc}")

    executor.submit(run)


@app.get("/health", response_model=HealthResponse)
def health() -> HealthResponse:
    return HealthResponse(ok=True)


@app.post("/prepare-product-menu", response_model=AcceptedResponse)
async def prepare_product_menu(request: Request) -> AcceptedResponse:
    record_id, secret = await parse_request(request)
    require_secret(secret)
    _submit(run_prepare_product_menu, record_id)
    return AcceptedResponse(ok=True, status="accepted", recordId=record_id, message="产品清单任务已接收。")


@app.post("/generate-delivery-tables", response_model=AcceptedResponse)
async def generate_delivery_tables(request: Request) -> AcceptedResponse:
    record_id, secret = await parse_request(request)
    require_secret(secret)
    _submit(run_generate_delivery_tables, record_id)
    return AcceptedResponse(ok=True, status="accepted", recordId=record_id, message="外卖数据统计任务已接收。")


@app.post("/generate-consumer-feedback", response_model=AcceptedResponse)
async def generate_consumer_feedback(request: Request) -> AcceptedResponse:
    record_id, secret = await parse_request(request)
    require_secret(secret)
    _submit(run_generate_consumer_feedback_tables, record_id)
    return AcceptedResponse(ok=True, status="accepted", recordId=record_id, message="消费者反馈统计任务已接收。")


@app.post("/generate-report", response_model=AcceptedResponse)
async def generate_report(request: Request) -> AcceptedResponse:
    record_id, secret = await parse_request(request)
    require_secret(secret)
    _submit(run_generate_report, record_id)
    return AcceptedResponse(ok=True, status="accepted", recordId=record_id, message="报告生成任务已接收。")


@app.post("/finalize-report", response_model=AcceptedResponse)
async def finalize_report(request: Request) -> AcceptedResponse:
    record_id, secret = await parse_request(request)
    require_secret(secret)
    _submit(run_finalize_report, record_id)
    return AcceptedResponse(
        ok=True,
        status="accepted",
        recordId=record_id,
        message="原始文件归档任务已接收。",
    )
