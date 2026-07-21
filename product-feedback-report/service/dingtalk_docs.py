from __future__ import annotations

import hashlib
import json
import mimetypes
import os
import re
import shutil
import subprocess
import time
import urllib.error
import urllib.request
import zipfile
from pathlib import Path
from typing import Any

from service import dingtalk_table


DEFAULT_LOCAL_UPLOAD_ROOT_FOLDER_NAME = "原始文件：竞品新品跟踪反馈"


class DownloadUrlUnavailableError(RuntimeError):
    pass


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


def node_url(node_id: str) -> str:
    clean_id = str(node_id or "").strip()
    if not clean_id:
        raise RuntimeError("钉钉文档节点 ID 为空，无法生成访问链接。")
    return f"https://alidocs.dingtalk.com/i/nodes/{clean_id}"


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
            found = _download_url(value)
            if found:
                return found
        for value in payload.values():
            found = _download_url(value)
            if found:
                return found
    if isinstance(payload, list):
        for value in payload:
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


def _file_sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _cache_manifest_path(path: Path) -> Path:
    return path.with_name(f".{path.name}.dingtalk-cache.json")


def _remote_version(info: dict[str, Any]) -> str:
    value = _info_value(info, "updateTime")
    return str(value) if value is not None else ""


def _cached_download_path(
    output_path: Path, node_id: str, info: dict[str, Any]
) -> Path | None:
    version = _remote_version(info)
    manifest_path = _cache_manifest_path(output_path)
    if not version or not output_path.is_file() or not manifest_path.is_file():
        return None
    try:
        manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return None
    if (
        str(manifest.get("nodeId") or "") != node_id
        or str(manifest.get("updateTime") or "") != version
        or str(manifest.get("filename") or "") != output_path.name
    ):
        return None
    expected_hash = str(manifest.get("sha256") or "")
    return output_path if expected_hash and _file_sha256(output_path) == expected_hash else None


def _write_download_cache(
    path: Path, node_id: str, info: dict[str, Any]
) -> None:
    version = _remote_version(info)
    if not version or not path.is_file():
        return
    manifest = {
        "nodeId": node_id,
        "updateTime": version,
        "filename": path.name,
        "sha256": _file_sha256(path),
    }
    manifest_path = _cache_manifest_path(path)
    temporary = manifest_path.with_name(f".{manifest_path.name}.tmp")
    temporary.write_text(json.dumps(manifest, ensure_ascii=False), encoding="utf-8")
    temporary.replace(manifest_path)


def _validate_downloaded_file(path: Path) -> None:
    if not path.is_file() or path.stat().st_size == 0:
        raise RuntimeError("下载结果为空。")
    if path.suffix.lower() != ".xlsx":
        return
    if not zipfile.is_zipfile(path):
        raise RuntimeError("下载结果不是有效的 xlsx 文件。")
    with zipfile.ZipFile(path) as workbook:
        names = set(workbook.namelist())
    if "[Content_Types].xml" not in names or "xl/workbook.xml" not in names:
        raise RuntimeError("下载结果缺少 xlsx 工作簿结构。")


def _download_binary(
    url: str, headers: dict[str, str], output_path: Path, channel: str
) -> Path:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    temporary = output_path.with_name(f".{output_path.name}.download")
    try:
        request = urllib.request.Request(url, headers=headers)
        try:
            with urllib.request.urlopen(request, timeout=120) as response:
                content = response.read()
        except urllib.error.HTTPError as exc:
            raise RuntimeError(f"{channel}签名下载失败（HTTP {exc.code}）。") from exc
        except urllib.error.URLError as exc:
            raise RuntimeError(f"{channel}签名下载网络失败。") from exc
        temporary.write_bytes(content)
        _validate_downloaded_file(temporary)
        temporary.replace(output_path)
        return output_path
    finally:
        if temporary.exists():
            temporary.unlink()


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


def local_upload_root_folder_id(config: dict[str, Any]) -> str:
    configured = str(config.get("localUploadRootFolderId") or "").strip()
    if configured:
        return node_id_from_url(configured) or configured
    name = str(
        config.get("localUploadRootFolderName") or DEFAULT_LOCAL_UPLOAD_ROOT_FOLDER_NAME
    ).strip()
    if not name:
        raise RuntimeError("本地上传归档根目录名称不能为空。")
    matches = [node for node in list_root_nodes(config) if _node_name(node) == name]
    if not matches:
        search_result = call_docs_tool(
            config, "search_documents", {"keyword": name, "pageSize": 30}
        )
        matches = [
            node
            for node in _iter_nodes(_data_payload(search_result))
            if _node_name(node) == name
            and (
                str(node.get("nodeType") or "").lower() == "folder"
                or str(node.get("extension") or "").lower() == "folder"
            )
        ]
    if not matches:
        raise RuntimeError(
            f"未找到钉钉文档归档根目录“{name}”，请配置 localUploadRootFolderId。"
        )
    if len(matches) > 1:
        raise RuntimeError(
            f"钉钉文档根目录存在多个同名“{name}”，请配置 localUploadRootFolderId 明确指定。"
        )
    node_id = _node_id(matches[0])
    if not node_id:
        raise RuntimeError(f"归档根目录“{name}”缺少节点 ID。")
    return node_id


def ensure_local_upload_task_folder(
    config: dict[str, Any], year: int, month: int, task_name: str
) -> tuple[str, str]:
    month_id, _ = ensure_local_upload_month_folder(config, year, month)
    clean_name = task_name.strip()
    if not clean_name:
        raise RuntimeError("任务文件夹名称不能为空。")
    return ensure_child_folder(config, month_id, clean_name)


def ensure_local_upload_month_folder(
    config: dict[str, Any], year: int, month: int
) -> tuple[str, str]:
    if year < 1000 or year > 9999:
        raise RuntimeError(f"报告日期年份无效：{year}")
    if month < 1 or month > 12:
        raise RuntimeError(f"报告日期月份无效：{month}")
    root_id = local_upload_root_folder_id(config)
    year_id, _ = ensure_child_folder(config, root_id, f"{year}年")
    return ensure_child_folder(config, year_id, f"{year}年{month}月")


def child_folders(config: dict[str, Any], parent_id: str) -> list[tuple[str, str]]:
    """返回直接子文件夹的 (node_id, name)。"""
    folders: list[tuple[str, str]] = []
    for node in list_nodes(config, parent_id):
        node_type = str(node.get("nodeType") or "").lower()
        extension = str(node.get("extension") or "").lower()
        if node_type != "folder" and extension != "folder":
            continue
        node_id = _node_id(node)
        name = _node_name(node)
        if node_id and name:
            folders.append((node_id, name))
    return folders


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
    info: dict[str, Any] = {}
    try:
        info = get_document_info(config, node_id)
    except Exception:
        pass
    cached = _cached_download_path(output_path, node_id, info)
    if cached:
        return cached

    result = call_docs_tool(config, "download_file", {"nodeId": node_id})
    url = _download_url(result)
    if not url:
        raise DownloadUrlUnavailableError(
            f"钉钉文档 MCP 未返回可下载 URL：{_node_name(node)} ({node_id})。"
            "请确认该文件是可下载附件，或更新钉钉文档 MCP 后重试。"
        )
    headers = _info_value(result, "headers") or {}
    request_headers = (
        {str(key): str(value) for key, value in headers.items()}
        if isinstance(headers, dict)
        else {}
    )
    downloaded = _download_binary(url, request_headers, output_path, "MCP")
    _write_download_cache(downloaded, node_id, info)
    return downloaded


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
    return None


def download_linked_file(config: dict[str, Any], value: Any, output_dir: Path) -> Path | None:
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

    # A DingTalk node link is the source of truth. Only fall back to local files
    # when the cell does not contain a downloadable DingTalk link; otherwise a
    # stale same-named file on Desktop can silently override the latest version.
    local = local_path_from_link(value)
    if local and local.is_file():
        return local
    named_local = local_file_from_attachment_name(value)
    if named_local:
        return named_local
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


def _cache_uploaded_file(
    config: dict[str, Any], path: Path, committed: dict[str, Any]
) -> None:
    node_id = _node_id(committed)
    if not node_id:
        return
    info = committed
    if not _remote_version(info):
        try:
            info = get_document_info(config, node_id)
        except Exception:
            return
    _write_download_cache(path, node_id, info)


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
        "convertToOnlineDoc": False,
    }
    if folder_id:
        commit_payload["folderId"] = folder_id
    if config.get("workspaceId"):
        commit_payload["workspaceId"] = config["workspaceId"]
    committed = _data_payload(call_docs_tool(config, "commit_uploaded_file", commit_payload))
    if isinstance(committed, dict):
        _cache_uploaded_file(config, path, committed)
        return _node_url(committed)
    return str(committed)
