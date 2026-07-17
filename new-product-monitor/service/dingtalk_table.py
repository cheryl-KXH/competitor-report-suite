from __future__ import annotations

import sys
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[1]
SCRIPTS_DIR = ROOT / "scripts"
if str(SCRIPTS_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPTS_DIR))

import generate_weekly_report as excel_report  # noqa: E402


@dataclass(frozen=True)
class ScheduleRow:
    record_id: str
    business: str
    start: date
    end: date
    year: int
    week: str


def schedule_config(configs: dict[str, dict[str, Any]]) -> dict[str, Any]:
    config = configs["report_rules"].get("weeklyReportTable", {})
    fields = config.get("fields", {})
    required = ("tableId",)
    missing = [key for key in required if not config.get(key)]
    missing.extend(f"fields.{key}" for key in ("reportUrl", "business", "startDate", "endDate", "year", "week", "status", "generatedAt") if not fields.get(key))
    if missing:
        raise RuntimeError("weeklyReportTable 配置不完整：" + "、".join(missing))
    return config


def _field_ids(configs: dict[str, dict[str, Any]]) -> dict[str, str]:
    return schedule_config(configs).get("fields", {})


def _extract_select_name(raw: Any) -> str:
    if isinstance(raw, dict):
        return str(raw.get("name") or "").strip()
    return str(raw or "").strip()


def _extract_formula_or_date(raw: Any) -> date:
    if isinstance(raw, dict) and "value" in raw:
        value = raw.get("value")
        if isinstance(value, list):
            raw = value[0] if value else ""
        else:
            raw = value
    parsed = excel_report.parse_date_value(raw)
    if not parsed:
        raise RuntimeError(f"无法解析日期字段：{raw}")
    return parsed


def _extract_year(raw: Any, fallback: date) -> int:
    if isinstance(raw, dict) and "value" in raw:
        value = raw.get("value")
        raw = value[0] if isinstance(value, list) and value else value
    try:
        return int(str(raw).strip())
    except (TypeError, ValueError):
        return fallback.year


def _parse_schedule_record(configs: dict[str, dict[str, Any]], record: dict[str, Any], record_id: str) -> ScheduleRow:
    fields = _field_ids(configs)
    cells = record.get("cells", {})
    business = _extract_select_name(cells.get(fields["business"]))
    start = _extract_formula_or_date(cells.get(fields["startDate"]))
    end = _extract_formula_or_date(cells.get(fields["endDate"]))
    if start > end:
        raise RuntimeError(f"开始时间晚于截止时间：{start} > {end}")
    week = str(cells.get(fields["week"]) or "").strip()
    if not business:
        raise RuntimeError("状态机行缺少业务")
    if not week:
        raise RuntimeError("状态机行缺少周次")
    return ScheduleRow(
        record_id=record_id,
        business=business,
        start=start,
        end=end,
        year=_extract_year(cells.get(fields["year"]), end),
        week=week,
    )


def fetch_schedule_row(configs: dict[str, dict[str, Any]], record_id: str) -> ScheduleRow:
    dingtalk = configs["dingtalk"]
    weekly = schedule_config(configs)
    fields = _field_ids(configs)
    requested_field_ids = list(dict.fromkeys(fields.values()))
    direct_payload = {
        "baseId": dingtalk["baseId"],
        "tableId": weekly["tableId"],
        "recordIds": [record_id],
        "fieldIds": requested_field_ids,
    }
    direct_data = excel_report.call_dingtalk_tool(dingtalk, "query_records", direct_payload)
    for record in direct_data.get("data", {}).get("records", []) or []:
        if str(record.get("recordId") or "") == record_id:
            return _parse_schedule_record(configs, record, record_id)

    cursor = None
    seen_cursors: set[str] = set()
    while True:
        payload: dict[str, Any] = {
            "baseId": dingtalk["baseId"],
            "tableId": weekly["tableId"],
            "limit": 100,
            "fieldIds": requested_field_ids,
        }
        if cursor:
            payload["cursor"] = cursor
        data = excel_report.call_dingtalk_tool(dingtalk, "query_records", payload)
        records = data.get("data", {}).get("records", []) or []
        for record in records:
            if str(record.get("recordId") or "") != record_id:
                continue
            return _parse_schedule_record(configs, record, record_id)
        next_cursor = data.get("data", {}).get("nextCursor")
        if not next_cursor or next_cursor in seen_cursors or not records:
            break
        seen_cursors.add(next_cursor)
        cursor = next_cursor
    raise RuntimeError(f"未找到状态机记录：{record_id}")


def status_cell(configs: dict[str, dict[str, Any]], status_name: str) -> dict[str, str]:
    config = schedule_config(configs)
    option_id = config.get("statusOptionIds", {}).get(status_name)
    if not option_id:
        raise RuntimeError(f"weeklyReportTable.statusOptionIds 缺少状态：{status_name}")
    return {"id": option_id, "name": status_name}


def report_url_cell(name: str, url: str) -> dict[str, str]:
    return {"text": name, "link": url}


def update_schedule_cells(configs: dict[str, dict[str, Any]], record_id: str, cells: dict[str, Any]) -> None:
    dingtalk = configs["dingtalk"]
    weekly = schedule_config(configs)
    payload = {
        "baseId": dingtalk["baseId"],
        "tableId": weekly["tableId"],
        "records": [{"recordId": record_id, "cells": cells}],
    }
    excel_report.call_dingtalk_tool(dingtalk, "update_records", payload)


def feedback_field_id(configs: dict[str, dict[str, Any]]) -> str:
    fields = _field_ids(configs)
    return fields.get("feedbackMessage") or fields.get("errorMessage") or ""


def mark_running(configs: dict[str, dict[str, Any]], record_id: str) -> None:
    fields = _field_ids(configs)
    cells = {
        fields["status"]: status_cell(configs, "等待1-2分钟"),
        fields["generatedAt"]: "",
    }
    feedback_field = feedback_field_id(configs)
    if feedback_field:
        cells[feedback_field] = ""
    update_schedule_cells(configs, record_id, cells)


def update_feedback(configs: dict[str, dict[str, Any]], record_id: str, message: str) -> None:
    feedback_field = feedback_field_id(configs)
    if not feedback_field:
        return
    update_schedule_cells(configs, record_id, {feedback_field: message[:5000]})


def mark_success(
    configs: dict[str, dict[str, Any]],
    record_id: str,
    report_name: str,
    report_url: str,
    feedback: str = "",
    status_name: str = "已生成",
) -> None:
    fields = _field_ids(configs)
    cells = {
        fields["status"]: status_cell(configs, status_name),
        fields["reportUrl"]: report_url_cell(report_name, report_url),
        fields["generatedAt"]: datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }
    feedback_field = feedback_field_id(configs)
    if feedback_field:
        cells[feedback_field] = feedback
    update_schedule_cells(configs, record_id, cells)


def mark_failed(configs: dict[str, dict[str, Any]], record_id: str, message: str) -> None:
    fields = _field_ids(configs)
    cells = {
        fields["status"]: status_cell(configs, "生成失败"),
    }
    feedback_field = feedback_field_id(configs)
    if feedback_field:
        cells[feedback_field] = message[:5000]
    update_schedule_cells(configs, record_id, cells)
