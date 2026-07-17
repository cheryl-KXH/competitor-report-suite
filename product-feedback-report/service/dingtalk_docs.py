from __future__ import annotations

import json
import mimetypes
import os
import re
import shutil
import subprocess
import time
import urllib.request
from pathlib import Path
from typing import Any

from service import dingtalk_table


def _selector(config: dict[str, Any], tool_name: str) -> list[str]:
    url = str(config.get("streamableHttpUrl") or os.getenv("DINGTALK_DOCS_MCP_URL") or "").strip()
    if url:
        return ["mcporter", "call", url, tool_name]
    server = str(config.get("serverName") or "dingtalk-docs").strip()
    return ["mcporter", "call", f"{server}.{tool_name}"]


def call_docs_tool(config: dict[str, Any], tool_name: str, payload: dict[str, Any]) -> dict[str, Any]:
    cmd = _selector(config, tool_name) + ["--args", json.dumps(payload, ensure_ascii=False)]
    try:
        result = subprocess.run(cmd, check=False, capture_output=True, text=True)
    except FileNotFoundError as exc:
        raise RuntimeError("找不到 mcporter，请先安装并配置钉钉文档 MCP。") from exc
    if result.returncode != 0:
        raise RuntimeError(f"钉钉文档 MCP 调用失败：{result.stderr.strip() or result.stdout.strip()}")
    data = dingtalk_table.parse_mcporter_json(result.stdout)
    if data.get("status") not in (None, "success", "ok"):
        raise RuntimeError(f"钉钉文档 MCP 返回失败：{json.dumps(data.get('error', data), ensure_ascii=False)}")
    return data


def extract_link(value: Any) -> str:
    if isinstance(value, dict):
        for key in ("link", "url", "docUrl", "nodeUrl", "resourceUrl", "webUrl"):
            if value.get(key):
                return str(value[key])
        if value.get("text"):
            return str(value["text"])
    if isinstance(value, list) and value:
        return extract_link(value[0])
    return str(value or "").strip()


def local_path_from_link(value: Any) -> Path | None:
    text = extract_link(value)
    if not text:
        return None
    path = Path(text)
    if path.exists():
        return path
    if text.startswith("file://"):
        path = Path(text.replace("file://", "", 1))
        return path if path.exists() else None
    return None


def local_folder_from_attachment_name(value: Any) -> Path | None:
    candidates = value if isinstance(value, list) else [value]
    search_roots = [Path.home() / "Desktop", Path.cwd()]
    for item in candidates:
        if not isinstance(item, dict):
            continue
        name = str(item.get("filename") or item.get("name") or "").strip()
        if not name:
            continue
        for root in search_roots:
            direct = root / name
            if direct.is_dir():
                return direct
    return None


def _linked_filename(value: Any) -> str:
    if not isinstance(value, dict):
        return ""
    return str(value.get("filename") or value.get("fileName") or value.get("name") or "").strip()


def local_file_from_attachment_name(value: Any) -> Path | None:
    candidates = value if isinstance(value, list) else [value]
    search_roots = [Path.home() / "Desktop", Path.cwd()]
    for item in candidates:
        if not isinstance(item, dict):
            continue
        name = _linked_filename(item)
        if not name:
            continue
        for root in search_roots:
            direct = root / name
            if direct.is_file():
                return direct
    return None


def node_id_from_url(value: str) -> str:
    match = re.search(r"/nodes/([^/?#]+)", value)
    return match.group(1) if match else ""


def _data_payload(result: dict[str, Any]) -> Any:
    if "data" in result:
        return result["data"]
    return result


def _iter_nodes(payload: Any) -> list[dict[str, Any]]:
    if isinstance(payload, list):
        return [item for item in payload if isinstance(item, dict)]
    if not isinstance(payload, dict):
        return []
    for key in ("nodes", "items", "files", "documents", "records", "list"):
        value = payload.get(key)
        if isinstance(value, list):
            return [item for item in value if isinstance(item, dict)]
    return []


def _node_id(node: dict[str, Any]) -> str:
    for key in ("nodeId", "id", "dentryUuid", "fileId"):
        value = node.get(key)
        if value:
            return str(value)
    return ""


def _node_name(node: dict[str, Any]) -> str:
    for key in ("name", "title", "nodeName", "fileName", "filename"):
        value = node.get(key)
        if value:
            return str(value)
    return _node_id(node)


def _node_url(node: dict[str, Any]) -> str:
    for key in ("docUrl", "url", "nodeUrl", "resourceUrl", "webUrl"):
        value = node.get(key)
        if value:
            return str(value)
    node_id = _node_id(node)
    return f"https://alidocs.dingtalk.com/i/nodes/{node_id}" if node_id else ""


def _download_url(payload: Any) -> str:
    if isinstance(payload, dict):
        for key in ("downloadUrl", "url", "resourceUrl", "fileUrl"):
            value = payload.get(key)
            if value and str(value).lower() != "null":
                return str(value)
        for value in payload.values():
            found = _download_url(value)
            if found:
                return found
    if isinstance(payload, str):
        match = re.search(r"https?://[^\s)）]+", payload)
        if match and match.group(0).lower() != "null":
            return match.group(0)
    return ""


def _info_value(payload: Any, *keys: str) -> Any:
    if isinstance(payload, dict):
        for key in keys:
            if payload.get(key):
                return payload[key]
        for value in payload.values():
            found = _info_value(value, *keys)
            if found:
                return found
    return None


def _safe_filename(name: str, extension: str = "") -> str:
    clean = re.sub(r'[\\/:*?"<>|]+', "_", name).strip() or "未命名文件"
    if extension and not clean.lower().endswith(f".{extension.lower()}"):
        clean = f"{clean}.{extension}"
    return clean


def list_nodes(config: dict[str, Any], folder_id: str) -> list[dict[str, Any]]:
    nodes: list[dict[str, Any]] = []
    page_token = None
    while True:
        payload: dict[str, Any] = {"folderId": folder_id, "pageSize": 50}
        if page_token:
            payload["pageToken"] = page_token
        result = call_docs_tool(config, "list_nodes", payload)
        data = _data_payload(result)
        nodes.extend(_iter_nodes(data))
        page_token = data.get("nextPageToken") if isinstance(data, dict) else None
        if not page_token:
            break
    return nodes


def list_root_nodes(config: dict[str, Any]) -> list[dict[str, Any]]:
    nodes: list[dict[str, Any]] = []
    page_token = None
    while True:
        payload: dict[str, Any] = {"pageSize": 50}
        if config.get("workspaceId"):
            payload["workspaceId"] = config["workspaceId"]
        if page_token:
            payload["pageToken"] = page_token
        result = call_docs_tool(config, "list_nodes", payload)
        data = _data_payload(result)
        nodes.extend(_iter_nodes(data))
        page_token = data.get("nextPageToken") if isinstance(data, dict) else None
        if not page_token:
            break
    return nodes


def find_child_by_name(config: dict[str, Any], parent_id: str | None, name: str) -> dict[str, Any] | None:
    nodes = list_nodes(config, parent_id) if parent_id else list_root_nodes(config)
    for node in nodes:
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
    node = find_child_by_name(config, parent_id, name) or create_folder(config, name, parent_id)
    node_id = _node_id(node)
    if not node_id:
        raise RuntimeError(f"未获取到文件夹节点 ID：{name}")
    return node_id, _node_url(node)


def ensure_target_root(config: dict[str, Any]) -> tuple[str | None, str]:
    root_folder_id = str(config.get("rootFolderId") or "").strip()
    if root_folder_id:
        return root_folder_id, f"https://alidocs.dingtalk.com/i/nodes/{root_folder_id}"
    folder_id = None
    folder_url = ""
    for name in config.get("rootFolderPath") or []:
        folder_id, folder_url = ensure_child_folder(config, folder_id, str(name))
    return folder_id, folder_url


def get_document_info(config: dict[str, Any], node_id: str) -> dict[str, Any]:
    data = _data_payload(call_docs_tool(config, "get_document_info", {"nodeId": node_id}))
    return data if isinstance(data, dict) else {}


def folder_id_for_linked_node(config: dict[str, Any], value: Any) -> str | None:
    candidates = value if isinstance(value, list) else [value]
    for item in candidates:
        node_id = node_id_from_url(extract_link(item))
        if not node_id:
            continue
        info = get_document_info(config, node_id)
        if str(info.get("nodeType") or "").lower() == "folder":
            return node_id
        folder_id = _info_value(info, "folderId", "parentId", "parentNodeId")
        if folder_id:
            return str(folder_id)
    return None


def _node_filename(node: dict[str, Any]) -> str:
    extension = str(node.get("extension") or "").strip().lstrip(".")
    return _safe_filename(_node_name(node), extension)


def delete_existing_file(config: dict[str, Any], folder_id: str | None, filename: str) -> None:
    if not folder_id:
        return
    expected = filename.strip()
    expected_stem = Path(expected).stem
    duplicate_pattern = re.compile(rf"^{re.escape(expected_stem)}\(\d+\)$")
    for node in list_nodes(config, folder_id):
        node_id = _node_id(node)
        if not node_id:
            continue
        node_filename = _node_filename(node)
        node_name = _node_name(node)
        node_stem = Path(node_filename).stem
        if node_filename == expected or node_name in {expected, expected_stem} or duplicate_pattern.match(node_name) or duplicate_pattern.match(node_stem):
            temp_name = f"待删除-{int(time.time())}-{node_id[:8]}-{node_filename}"
            call_docs_tool(config, "rename_document", {"nodeId": node_id, "newName": temp_name[:255]})
            call_docs_tool(config, "delete_document", {"nodeId": node_id})


def download_file(config: dict[str, Any], node: dict[str, Any], output_dir: Path) -> Path:
    node_id = _node_id(node)
    if not node_id:
        raise RuntimeError(f"文件节点缺少 nodeId：{node}")
    extension = str(node.get("extension") or "").strip().lstrip(".")
    output_path = output_dir / _safe_filename(_node_name(node), extension)
    result = call_docs_tool(config, "download_file", {"nodeId": node_id})
    url = _download_url(result)
    if not url:
        raise RuntimeError(
            f"钉钉文档 MCP 未返回可下载 URL：{_node_name(node)} ({node_id})。"
            "请确认该文件是可下载附件，或更新钉钉文档 MCP 后重试。"
        )
    output_dir.mkdir(parents=True, exist_ok=True)
    request = urllib.request.Request(url)
    with urllib.request.urlopen(request, timeout=120) as response:
        output_path.write_bytes(response.read())
    return output_path


def download_folder(config: dict[str, Any], folder_id: str, output_dir: Path) -> Path:
    output_dir.mkdir(parents=True, exist_ok=True)
    downloaded = 0
    for node in list_nodes(config, folder_id):
        node_type = str(node.get("nodeType") or "").lower()
        extension = str(node.get("extension") or "").lower()
        node_id = _node_id(node)
        if node.get("hasChildren") or node_type == "folder":
            child_dir = output_dir / _safe_filename(_node_name(node))
            download_folder(config, node_id, child_dir)
            continue
        if extension != "xlsx":
            continue
        download_file(config, node, output_dir)
        downloaded += 1
    if downloaded == 0 and not any(output_dir.rglob("*.xlsx")):
        raise RuntimeError(f"钉钉文件夹内未找到可下载 xlsx 文件：{folder_id}")
    return output_dir


def download_linked_folder(config: dict[str, Any], value: Any, output_dir: Path) -> Path | None:
    local = local_path_from_link(value)
    if local and local.is_dir():
        return local
    named_local = local_folder_from_attachment_name(value)
    if named_local:
        return named_local
    named_file = local_file_from_attachment_name(value)
    if named_file:
        output_dir.mkdir(parents=True, exist_ok=True)
        target = output_dir / named_file.name
        if named_file.resolve() != target.resolve():
            shutil.copy2(named_file, target)
        return output_dir
    candidates = value if isinstance(value, list) else [value]
    for item in candidates:
        link = extract_link(item)
        node_id = node_id_from_url(link)
        if node_id:
            filename = _linked_filename(item)
            if filename.lower().endswith(".xlsx"):
                download_file(
                    config,
                    {"nodeId": node_id, "name": filename, "extension": Path(filename).suffix.lstrip(".")},
                    output_dir,
                )
                return output_dir
            info = get_document_info(config, node_id)
            node_type = str(_info_value(info, "nodeType", "type") or "").lower()
            if node_type == "folder":
                return download_folder(config, node_id, output_dir)
            info_name = str(_info_value(info, "name", "title", "nodeName", "fileName", "filename") or filename)
            extension = str(_info_value(info, "extension") or Path(info_name).suffix.lstrip(".")).lower()
            if extension == "xlsx":
                download_file(
                    config,
                    {"nodeId": node_id, "name": info_name or f"{node_id}.xlsx", "extension": extension},
                    output_dir,
                )
                return output_dir
            raise RuntimeError(f"钉钉链接既不是文件夹，也不是 xlsx 文件：{link}")
    return None


def download_linked_file(config: dict[str, Any], value: Any, output_dir: Path) -> Path | None:
    local = local_path_from_link(value)
    if local and local.is_file():
        return local
    named_local = local_file_from_attachment_name(value)
    if named_local:
        return named_local
    candidates = value if isinstance(value, list) else [value]
    for item in candidates:
        link = extract_link(item)
        node_id = node_id_from_url(link)
        if not node_id:
            continue
        filename = _linked_filename(item)
        info: dict[str, Any] = {}
        if not filename.lower().endswith(".xlsx"):
            info = get_document_info(config, node_id)
            node_type = str(_info_value(info, "nodeType", "type") or "").lower()
            if node_type == "folder":
                return None
            filename = str(_info_value(info, "name", "title", "nodeName", "fileName", "filename") or filename)
        extension = str(_info_value(info, "extension") or Path(filename).suffix.lstrip(".")).lower()
        if extension != "xlsx":
            raise RuntimeError(f"钉钉链接不是 xlsx 文件：{link}")
        return download_file(
            config,
            {"nodeId": node_id, "name": filename or f"{node_id}.xlsx", "extension": extension},
            output_dir,
        )
    return None


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


def upload_file(config: dict[str, Any], path: Path, folder_id: str | None = None) -> str:
    if not path.exists():
        raise RuntimeError(f"待上传文件不存在：{path}")
    if not folder_id:
        folder_id, _ = ensure_target_root(config)
    delete_existing_file(config, folder_id, path.name)
    payload: dict[str, Any] = {
        "fileName": path.name,
        "fileSize": path.stat().st_size,
        "mimeType": mimetypes.guess_type(path.name)[0] or "application/octet-stream",
    }
    if folder_id:
        payload["folderId"] = folder_id
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
    request_headers["Content-Length"] = str(path.stat().st_size)
    request = urllib.request.Request(str(upload_url), data=path.read_bytes(), method="PUT", headers=request_headers)
    with urllib.request.urlopen(request, timeout=300) as response:
        if response.status != 200:
            raise RuntimeError(f"上传文件失败：HTTP {response.status}")
    commit_payload: dict[str, Any] = {
        "uploadKey": upload_key,
        "name": path.name,
        "fileSize": path.stat().st_size,
    }
    if folder_id:
        commit_payload["folderId"] = folder_id
    if config.get("workspaceId"):
        commit_payload["workspaceId"] = config["workspaceId"]
    committed = _data_payload(call_docs_tool(config, "commit_uploaded_file", commit_payload))
    if isinstance(committed, dict):
        return _node_url(committed)
    return str(committed)
