from __future__ import annotations

import json
import mimetypes
import os
import re
import subprocess
import sys
import time
import urllib.request
from pathlib import Path
from typing import Any, Callable

ROOT = Path(__file__).resolve().parents[1]
CONFIG_DIR = ROOT / "config"
SCRIPTS_DIR = ROOT / "scripts"
if str(SCRIPTS_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPTS_DIR))

import generate_weekly_report as excel_report  # noqa: E402


def load_docs_config() -> dict[str, Any]:
    path = CONFIG_DIR / "dingtalk_docs.json"
    config: dict[str, Any] = {}
    if path.exists():
        config = excel_report.load_json(path)
    config.setdefault("serverName", "dingtalk-docs")
    config.setdefault("streamableHttpUrl", "")
    config.setdefault("workspaceId", "")
    config.setdefault("businessRootFolderIds", {})
    config.setdefault("rootFolderId", "")
    config.setdefault("rootFolderPath", [])
    env_url = os.getenv("DINGTALK_DOCS_MCP_URL", "").strip()
    if env_url:
        config["streamableHttpUrl"] = env_url
    env_workspace_id = os.getenv("DINGTALK_DOCS_WORKSPACE_ID", "").strip()
    if env_workspace_id:
        config["workspaceId"] = env_workspace_id
    env_root_folder_id = os.getenv("DINGTALK_DOCS_ROOT_FOLDER_ID", "").strip()
    if env_root_folder_id:
        config["rootFolderId"] = env_root_folder_id
    return config


def mcporter_selector(config: dict[str, Any], tool_name: str) -> list[str]:
    url = str(config.get("streamableHttpUrl") or "").strip()
    if url:
        return ["mcporter", "call", url, tool_name]
    server = str(config.get("serverName") or "dingtalk-docs").strip()
    return ["mcporter", "call", f"{server}.{tool_name}"]


def call_docs_tool(config: dict[str, Any], tool_name: str, payload: dict[str, Any]) -> dict[str, Any]:
    cmd = mcporter_selector(config, tool_name) + ["--args", json.dumps(payload, ensure_ascii=False)]
    try:
        result = subprocess.run(cmd, check=False, capture_output=True, text=True)
    except FileNotFoundError as exc:
        raise RuntimeError("找不到 mcporter，请先安装并配置钉钉文档 MCP。") from exc
    if result.returncode != 0:
        raise RuntimeError(f"钉钉文档 MCP 调用失败：{result.stderr.strip() or result.stdout.strip()}")
    try:
        data = excel_report.parse_mcporter_json(result.stdout)
    except json.JSONDecodeError as exc:
        raise RuntimeError(f"钉钉文档 MCP 返回不是 JSON：{result.stdout[:1000]}") from exc
    if data.get("status") not in (None, "success", "ok"):
        raise RuntimeError(f"钉钉文档 MCP 返回失败：{json.dumps(data.get('error', data), ensure_ascii=False)}")
    return data


def _data_payload(result: dict[str, Any]) -> Any:
    data = result.get("data") if "data" in result else result
    if isinstance(data, dict) and "result" in data:
        return data["result"]
    return data


def _node_id(node: dict[str, Any]) -> str:
    for key in ("nodeId", "id", "dentryUuid", "fileId"):
        value = node.get(key)
        if value:
            return str(value)
    return ""


def _node_name(node: dict[str, Any]) -> str:
    for key in ("name", "title", "nodeName", "fileName"):
        value = node.get(key)
        if value:
            return str(value)
    return ""


def _node_url(node: dict[str, Any]) -> str:
    for key in ("docUrl", "url", "nodeUrl", "resourceUrl", "webUrl"):
        value = node.get(key)
        if value:
            return str(value)
    node_id = _node_id(node)
    return f"https://alidocs.dingtalk.com/i/nodes/{node_id}" if node_id else ""


def _iter_nodes(data: Any) -> list[dict[str, Any]]:
    if isinstance(data, list):
        return [item for item in data if isinstance(item, dict)]
    if not isinstance(data, dict):
        return []
    for key in ("nodes", "items", "files", "documents", "records", "list"):
        value = data.get(key)
        if isinstance(value, list):
            return [item for item in value if isinstance(item, dict)]
    return []


def list_nodes(config: dict[str, Any], folder_id: str | None = None) -> list[dict[str, Any]]:
    workspace_id = str(config.get("workspaceId") or "").strip()
    nodes: list[dict[str, Any]] = []
    page_token = None
    while True:
        payload: dict[str, Any] = {"pageSize": 50}
        if folder_id:
            payload["folderId"] = folder_id
        elif workspace_id:
            payload["workspaceId"] = workspace_id
        if page_token:
            payload["pageToken"] = page_token
        data = _data_payload(call_docs_tool(config, "list_nodes", payload))
        nodes.extend(_iter_nodes(data))
        page_token = data.get("nextPageToken") if isinstance(data, dict) else None
        if not page_token:
            break
    return nodes


def find_child_by_name(config: dict[str, Any], parent_id: str | None, name: str) -> dict[str, Any] | None:
    for node in list_nodes(config, parent_id):
        if _node_name(node) == name:
            return node
    return None


def create_folder(config: dict[str, Any], name: str, parent_id: str | None = None) -> dict[str, Any]:
    payload: dict[str, Any] = {"name": name}
    if parent_id:
        payload["folderId"] = parent_id
    elif config.get("workspaceId"):
        payload["workspaceId"] = config["workspaceId"]
    data = _data_payload(call_docs_tool(config, "create_folder", payload))
    if isinstance(data, dict):
        return data
    raise RuntimeError(f"创建文件夹返回异常：{data}")


def ensure_child_folder(config: dict[str, Any], parent_id: str | None, name: str) -> tuple[str, str]:
    existing = find_child_by_name(config, parent_id, name)
    node = existing or create_folder(config, name, parent_id)
    node_id = _node_id(node)
    if not node_id:
        raise RuntimeError(f"未获取到文件夹节点 ID：{name}")
    return node_id, _node_url(node)


def ensure_target_root(config: dict[str, Any]) -> tuple[str | None, str]:
    root_folder_id = str(config.get("rootFolderId") or "").strip()
    if root_folder_id:
        return root_folder_id, f"https://alidocs.dingtalk.com/i/nodes/{root_folder_id}" if not root_folder_id.startswith("http") else root_folder_id
    folder_id = None
    folder_url = ""
    for name in config.get("rootFolderPath") or []:
        folder_id, folder_url = ensure_child_folder(config, folder_id, str(name))
    if folder_id:
        return folder_id, folder_url
    if not config.get("workspaceId"):
        raise RuntimeError("config/dingtalk_docs.json 需要配置 rootFolderId、rootFolderPath+workspaceId 或 DINGTALK_DOCS_MCP_URL 对应默认位置")
    return None, str(config.get("workspaceId"))


def ensure_business_root(config: dict[str, Any], business: str) -> tuple[str | None, str]:
    root_folder_id = str((config.get("businessRootFolderIds") or {}).get(business) or "").strip()
    if root_folder_id:
        return root_folder_id, (
            f"https://alidocs.dingtalk.com/i/nodes/{root_folder_id}"
            if not root_folder_id.startswith("http")
            else root_folder_id
        )
    return ensure_target_root(config)


def normalize_year_folder_name(year: int | str) -> str:
    year_text = str(year).strip()
    if not year_text.isdigit() or len(year_text) != 4:
        raise RuntimeError(f"年份文件夹必须是四位数字：{year}")
    return year_text


def ensure_report_folder(config: dict[str, Any], business: str, year: int | str, report_stem: str) -> tuple[str, str]:
    root_id, _ = ensure_business_root(config, business)
    year_id, _ = ensure_child_folder(config, root_id, normalize_year_folder_name(year))
    return ensure_child_folder(config, year_id, report_stem)


def delete_child_folder_if_exists(config: dict[str, Any], parent_id: str, name: str) -> None:
    existing = find_child_by_name(config, parent_id, name)
    if not existing:
        return
    node_id = _node_id(existing)
    if not node_id:
        raise RuntimeError(f"找到同名文件夹但未获取到节点 ID：{name}")
    call_docs_tool(config, "delete_document", {"nodeId": node_id})
    for _ in range(10):
        if not find_child_by_name(config, parent_id, name):
            return
        time.sleep(0.5)
    raise RuntimeError(f"删除旧周报文件夹后仍能查询到同名文件夹：{name}")


def replace_report_folder(config: dict[str, Any], business: str, year: int | str, report_stem: str) -> tuple[str, str]:
    root_id, _ = ensure_business_root(config, business)
    year_id, _ = ensure_child_folder(config, root_id, normalize_year_folder_name(year))
    delete_child_folder_if_exists(config, year_id, report_stem)
    return ensure_child_folder(config, year_id, report_stem)


def delete_named_children(config: dict[str, Any], folder_id: str, names: set[str]) -> None:
    for node in list_nodes(config, folder_id):
        if _node_name(node) not in names:
            continue
        node_id = _node_id(node)
        if node_id:
            call_docs_tool(config, "delete_document", {"nodeId": node_id})


def delete_report_file_variants(config: dict[str, Any], folder_id: str, paths: list[Path]) -> None:
    patterns = [
        re.compile(rf"^{re.escape(path.stem)}(?:\(\d+\))?{re.escape(path.suffix)}$")
        for path in paths
    ]
    for node in list_nodes(config, folder_id):
        name = _node_name(node)
        if not any(pattern.match(name) for pattern in patterns):
            continue
        node_id = _node_id(node)
        if node_id:
            call_docs_tool(config, "delete_document", {"nodeId": node_id})


def _upload_info_value(data: Any, *keys: str) -> Any:
    if isinstance(data, dict):
        for key in keys:
            if key in data:
                return data[key]
        for value in data.values():
            found = _upload_info_value(value, *keys)
            if found:
                return found
    return None


def upload_file(config: dict[str, Any], folder_id: str, path: Path) -> dict[str, Any]:
    mime_type = mimetypes.guess_type(path.name)[0] or "application/octet-stream"
    size = path.stat().st_size
    payload = {
        "fileName": path.name,
        "fileSize": size,
        "mimeType": mime_type,
        "folderId": folder_id,
    }
    if config.get("workspaceId"):
        payload["workspaceId"] = config["workspaceId"]
    data = _data_payload(call_docs_tool(config, "get_file_upload_info", payload))
    upload_key = _upload_info_value(data, "uploadKey")
    upload_url = _upload_info_value(data, "uploadUrl", "resourceUrl", "url")
    headers = _upload_info_value(data, "headers") or {}
    if isinstance(upload_url, list):
        upload_url = upload_url[0] if upload_url else ""
    if not upload_key or not upload_url:
        raise RuntimeError(f"上传凭证缺少 uploadKey 或 uploadUrl：{data}")
    request_headers = {str(k): str(v) for k, v in headers.items()} if isinstance(headers, dict) else {}
    request_headers["Content-Type"] = ""
    request_headers["Content-Length"] = str(size)
    request = urllib.request.Request(str(upload_url), data=path.read_bytes(), method="PUT", headers=request_headers)
    with urllib.request.urlopen(request, timeout=300) as response:
        if response.status != 200:
            raise RuntimeError(f"上传文件失败：HTTP {response.status}")
    commit_payload = {
        "uploadKey": upload_key,
        "name": path.name,
        "fileSize": size,
        "folderId": folder_id,
        "convertToOnlineDoc": False,
    }
    if config.get("workspaceId"):
        commit_payload["workspaceId"] = config["workspaceId"]
    committed = _data_payload(call_docs_tool(config, "commit_uploaded_file", commit_payload))
    return committed if isinstance(committed, dict) else {"result": committed}


def upload_report_directory(
    config: dict[str, Any],
    business: str,
    year: int | str,
    report_stem: str,
    paths: list[Path],
    progress_callback: Callable[[str], None] | None = None,
) -> str:
    folder_id, folder_url = replace_report_folder(config, business, year, report_stem)
    if progress_callback:
        progress_callback("5/5 正在上传至钉钉文档")
    for path in paths:
        started_at = time.perf_counter()
        upload_file(config, folder_id, path)
        print(f"耗时：上传 {path.name} ({path.stat().st_size / 1024 / 1024:.1f}MB) {time.perf_counter() - started_at:.1f}s", file=sys.stderr)
    return folder_url or f"https://alidocs.dingtalk.com/i/nodes/{folder_id}"
