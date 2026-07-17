from __future__ import annotations

import json
import os
import subprocess
from dataclasses import dataclass
from datetime import datetime
from typing import Any


@dataclass
class TaskRecord:
    record_id: str
    cells: dict[str, Any]


def _selector(config: dict[str, Any], tool_name: str) -> list[str]:
    url = str(config.get("streamableHttpUrl") or os.getenv("DINGTALK_MCP_URL") or "").strip()
    if url:
        return ["mcporter", "call", url, tool_name]
    server = str(config.get("serverName") or "dingtalk-ai-table").strip()
    return ["mcporter", "call", f"{server}.{tool_name}"]


def parse_mcporter_json(stdout: str) -> dict[str, Any]:
    text = stdout.strip()
    if not text:
        return {}
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        start = text.find("{")
        end = text.rfind("}")
        if start >= 0 and end > start:
            return json.loads(text[start : end + 1])
        raise


def call_table_tool(config: dict[str, Any], tool_name: str, payload: dict[str, Any]) -> dict[str, Any]:
    cmd = _selector(config, tool_name) + ["--args", json.dumps(payload, ensure_ascii=False)]
    try:
        result = subprocess.run(cmd, check=False, capture_output=True, text=True)
    except FileNotFoundError as exc:
        raise RuntimeError("找不到 mcporter，请先安装并配置 dingtalk-ai-table。") from exc
    if result.returncode != 0:
        raise RuntimeError(f"钉钉 AI 表调用失败：{result.stderr.strip() or result.stdout.strip()}")
    data = parse_mcporter_json(result.stdout)
    if data.get("status") not in (None, "success", "ok"):
        raise RuntimeError(f"钉钉 AI 表返回失败：{json.dumps(data.get('error', data), ensure_ascii=False)}")
    return data


def field_defs(configs: dict[str, dict[str, Any]]) -> dict[str, dict[str, Any]]:
    return configs.get("field_mapping", {}).get("fields", {})


def field_id(configs: dict[str, dict[str, Any]], key: str) -> str:
    return str(field_defs(configs).get(key, {}).get("fieldId") or "").strip()


def _cells_by_logical_key(configs: dict[str, dict[str, Any]], cells: dict[str, Any]) -> dict[str, Any]:
    mapped: dict[str, Any] = {}
    for key, meta in field_defs(configs).items():
        fid = str(meta.get("fieldId") or "").strip()
        names = [str(name) for name in meta.get("fieldNames", [])]
        if fid and fid in cells:
            mapped[key] = cells[fid]
            continue
        for name in names:
            if name in cells:
                mapped[key] = cells[name]
                break
    return mapped


def fetch_record(configs: dict[str, dict[str, Any]], record_id: str) -> TaskRecord:
    dingtalk = configs["dingtalk"]
    base_id = str(dingtalk.get("baseId") or "").strip()
    table_id = str(dingtalk.get("tableId") or "").strip()
    if not base_id or not table_id:
        raise RuntimeError("config/dingtalk.json 缺少 baseId 或 tableId。")
    requested_field_ids = [meta["fieldId"] for meta in field_defs(configs).values() if meta.get("fieldId")]
    payload: dict[str, Any] = {"baseId": base_id, "tableId": table_id, "recordIds": [record_id]}
    if requested_field_ids:
        payload["fieldIds"] = requested_field_ids
    data = call_table_tool(dingtalk, "query_records", payload)
    records = data.get("data", {}).get("records", []) or []
    for record in records:
        if str(record.get("recordId") or "") == record_id:
            return TaskRecord(record_id=record_id, cells=_cells_by_logical_key(configs, record.get("cells", {})))
    raise RuntimeError(f"未找到记录：{record_id}")


def link_cell(name: str, url: str) -> dict[str, str]:
    return {"text": name, "link": url}


def _update_cells(configs: dict[str, dict[str, Any]], record_id: str, logical_cells: dict[str, Any]) -> None:
    dingtalk = configs["dingtalk"]
    base_id = str(dingtalk.get("baseId") or "").strip()
    table_id = str(dingtalk.get("tableId") or "").strip()
    cells: dict[str, Any] = {}
    for key, value in logical_cells.items():
        fid = field_id(configs, key)
        if fid:
            cells[fid] = value
    if not cells:
        return
    payload = {"baseId": base_id, "tableId": table_id, "records": [{"recordId": record_id, "cells": cells}]}
    call_table_tool(dingtalk, "update_records", payload)


def mark_status(configs: dict[str, dict[str, Any]], record_id: str, status: str, feedback: str = "") -> None:
    cells: dict[str, Any] = {"status": status}
    if feedback:
        cells["feedback"] = feedback[:5000]
    _update_cells(configs, record_id, cells)


def update_feedback(configs: dict[str, dict[str, Any]], record_id: str, message: str) -> None:
    _update_cells(configs, record_id, {"feedback": message[:5000]})


def mark_failed(configs: dict[str, dict[str, Any]], record_id: str, message: str) -> None:
    status = configs.get("report_rules", {}).get("statuses", {}).get("failed", "生成失败")
    mark_status(configs, record_id, status, message)


def mark_links(configs: dict[str, dict[str, Any]], record_id: str, links: dict[str, tuple[str, str]], status: str | None = None) -> None:
    cells: dict[str, Any] = {}
    attachment_keys = {"productMenu", "meituanData", "elemeData", "jdData"}
    for key, (name, url) in links.items():
        if key in attachment_keys:
            cells[key] = [{"url": url, "name": name}]
        else:
            cells[key] = link_cell(name, url)
    if status:
        cells["status"] = status
    cells["generatedAt"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    _update_cells(configs, record_id, cells)
