from __future__ import annotations

from datetime import date, datetime
import re
import traceback
import sys
import shutil
import time
import urllib.request
from pathlib import Path
from typing import Any

from openpyxl import load_workbook

from service import dingtalk_docs, dingtalk_table
from service.config import load_configs, output_root
from scripts.delivery.generate_tables import generate_delivery_tables
from scripts.delivery.prepare_product_menu import prepare_product_menu
from scripts.reporting.generate_report import generate_report
from scripts.social.generate_tables import generate_consumer_feedback_tables


CONSUMER_FEEDBACK_TABLE_ID = "hvcu7Bw"
DELIVERY_OUTPUT_PLATFORMS = (
    ("meituanData", "美团"),
    ("elemeData", "饿了么"),
    ("jdData", "京东"),
)
DELIVERY_OUTPUT_FIELDS = ("meituanData", "elemeData", "jdData")
SOCIAL_INPUT_PLATFORMS = (
    ("weibo", "微博"),
    ("xiaohongshu", "小红书"),
    ("douyin", "抖音"),
)
SOCIAL_INPUT_FIELDS = tuple(key for key, _ in SOCIAL_INPUT_PLATFORMS)
AI_TABLE_ATTACHMENT_URL_ATTEMPTS = 10
AI_TABLE_ATTACHMENT_URL_DELAY_SECONDS = 1.0


def _status(configs: dict[str, dict[str, Any]], key: str, fallback: str) -> str:
    return configs.get("report_rules", {}).get("statuses", {}).get(key, fallback)


def _task_input_dir(configs: dict[str, dict[str, Any]], record: dingtalk_table.TaskRecord) -> Path:
    value = record.cells.get("deliveryData") or record.cells.get("allData")
    raw_dir = output_root(configs) / record.record_id / "raw_data"
    linked_dir = dingtalk_docs.download_linked_folder(configs.get("dingtalk_docs", {}), value, raw_dir)
    if linked_dir and linked_dir.is_dir():
        return linked_dir
    raise RuntimeError("输入数据字段为空，或不是可识别的本机文件/钉钉文档附件。")


def _task_delivery_input_dir(configs: dict[str, dict[str, Any]], record: dingtalk_table.TaskRecord) -> Path:
    value = record.cells.get("deliveryData")
    if not _has_dingtalk_node_link(value):
        raise RuntimeError("“外卖数据”字段缺少可识别的钉钉 xlsx 附件。")
    raw_dir = output_root(configs) / record.record_id / "raw_data"
    linked_dir = dingtalk_docs.download_linked_folder(
        configs.get("dingtalk_docs", {}), value, raw_dir
    )
    if linked_dir and linked_dir.is_dir():
        return linked_dir
    raise RuntimeError("“外卖数据”字段不是可下载的钉钉 xlsx 附件。")


def _task_docs_folder_id(
    configs: dict[str, dict[str, Any]],
    record: dingtalk_table.TaskRecord,
    keys: tuple[str, ...] = ("deliveryData", "productMenu", *DELIVERY_OUTPUT_FIELDS),
) -> str | None:
    docs_config = configs.get("dingtalk_docs", {})
    for key in keys:
        value = record.cells.get(key)
        if not value:
            continue
        folder_id = dingtalk_docs.folder_id_for_linked_node(docs_config, value)
        if folder_id:
            return folder_id
    return None


def _value_items(value: Any) -> list[Any]:
    if isinstance(value, list):
        return [item for item in value if item]
    return [value] if value else []


def _has_dingtalk_node_link(value: Any) -> bool:
    return any(
        dingtalk_docs.node_id_from_url(dingtalk_docs.extract_link(item))
        for item in _value_items(value)
    )


def _local_item_name(item: Any) -> str:
    if isinstance(item, dict):
        for key in ("filename", "fileName", "name"):
            if item.get(key):
                return str(item[key]).strip()
    return ""


def _single_ai_table_attachment(
    value: Any, field_label: str, *, require_url: bool = True
) -> tuple[str, str] | None:
    items = _value_items(value)
    attachments = [item for item in items if isinstance(item, dict) and item.get("resourceId")]
    if not attachments:
        return None
    if len(items) != 1 or len(attachments) != 1:
        raise RuntimeError(f"“{field_label}”字段仅支持一个本地上传的 xlsx 附件。")
    item = attachments[0]
    name = _local_item_name(item)
    if not name:
        raise RuntimeError(f"“{field_label}”本地上传附件缺少文件名。")
    if Path(name).name != name:
        raise RuntimeError(f"“{field_label}”本地上传附件文件名无效：{name}")
    if Path(name).suffix.lower() != ".xlsx":
        raise RuntimeError(f"“{field_label}”字段只支持 xlsx 文件：{name}")
    # The current AI Table MCP returns the signed download URL in `url` while
    # `resourceUrl` can be a relative metadata endpoint such as /core/api/....
    resource_url = item.get("url") or item.get("resourceUrl")
    if isinstance(resource_url, list):
        resource_url = resource_url[0] if resource_url else ""
    url = str(resource_url or "").strip()
    if require_url and not url.startswith(("http://", "https://")):
        raise RuntimeError(f"“{field_label}”本地上传附件缺少可下载 URL：{name}")
    return name, url if url.startswith(("http://", "https://")) else ""


def _wait_for_ai_table_attachment_url(
    configs: dict[str, dict[str, Any]],
    record: dingtalk_table.TaskRecord,
    field_key: str,
    field_label: str,
    *,
    attempts: int = AI_TABLE_ATTACHMENT_URL_ATTEMPTS,
    delay_seconds: float = AI_TABLE_ATTACHMENT_URL_DELAY_SECONDS,
) -> tuple[str, str] | None:
    for attempt in range(attempts):
        attachment = _single_ai_table_attachment(
            record.cells.get(field_key), field_label, require_url=False
        )
        if attachment is None:
            return None
        if attachment[1]:
            return attachment
        if attempt + 1 < attempts:
            if delay_seconds > 0:
                time.sleep(delay_seconds)
            refreshed = dingtalk_table.fetch_record(configs, record.record_id)
            record.cells[field_key] = refreshed.cells.get(field_key)
    name = _single_ai_table_attachment(
        record.cells.get(field_key), field_label, require_url=False
    )[0]
    raise RuntimeError(
        f"“{field_label}”本地上传附件在等待后仍缺少可下载 URL：{name}。"
        "请稍后重试，或确认钉钉 AI 表 MCP 能返回附件 resourceUrl。"
    )


def _download_ai_table_attachment(
    configs: dict[str, dict[str, Any]],
    record: dingtalk_table.TaskRecord,
    field_key: str,
    field_label: str,
    attachment: tuple[str, str],
) -> Path:
    name, url = attachment
    task_root = output_root(configs) / record.record_id
    if field_key == "deliveryData":
        target_dir = task_root / "raw_data"
    elif field_key in SOCIAL_INPUT_FIELDS:
        target_dir = task_root / "social_inputs" / field_key
    else:
        target_dir = task_root
    target_dir.mkdir(parents=True, exist_ok=True)
    target = target_dir / name
    temporary = target.with_name(f".{target.name}.download")
    try:
        request = urllib.request.Request(url)
        with urllib.request.urlopen(request, timeout=120) as response:
            content = response.read()
        if not content:
            raise RuntimeError(f"“{field_label}”本地上传附件下载结果为空：{name}")
        temporary.write_bytes(content)
        temporary.replace(target)
    except Exception as exc:
        if isinstance(exc, RuntimeError):
            raise
        raise RuntimeError(f"“{field_label}”本地上传附件下载失败：{name}：{exc}") from exc
    finally:
        if temporary.exists():
            temporary.unlink()
    return target


def _single_local_xlsx(value: Any, field_label: str) -> Path:
    items = _value_items(value)
    if not items:
        raise RuntimeError(f"“{field_label}”字段为空，请先放入一个 xlsx 文件。")
    candidates: dict[Path, Path] = {}
    search_roots = (Path.home() / "Desktop", Path.cwd())
    for item in items:
        link = dingtalk_docs.extract_link(item)
        if dingtalk_docs.node_id_from_url(link):
            continue
        if link.startswith(("http://", "https://")):
            raise RuntimeError(f"“{field_label}”字段包含无法识别的非钉钉链接：{link}")
        raw_path = link.replace("file://", "", 1) if link.startswith("file://") else link
        path = Path(raw_path) if raw_path else None
        explicit_path = bool(
            path
            and (
                link.startswith("file://")
                or path.is_absolute()
                or path.parent != Path(".")
            )
        )
        if explicit_path and path and path.exists():
            resolved = path.resolve()
            candidates[resolved] = path
            continue
        name = _local_item_name(item) or raw_path
        if not name:
            continue
        for root in search_roots:
            direct = root / name
            if direct.exists():
                candidates[direct.resolve()] = direct
    if not candidates:
        raise RuntimeError(f"“{field_label}”字段未找到可上传的本机 xlsx 文件。")
    if len(candidates) != 1:
        names = "、".join(sorted(path.name for path in candidates.values()))
        raise RuntimeError(f"“{field_label}”字段匹配到多个本机文件：{names}。仅支持单个 xlsx。")
    path = next(iter(candidates.values()))
    if path.is_dir():
        raise RuntimeError(f"“{field_label}”字段指向本机目录，仅支持单个 xlsx 文件。")
    if path.suffix.lower() != ".xlsx":
        raise RuntimeError(f"“{field_label}”字段只支持 xlsx 文件：{path.name}")
    return path


def _task_brand(record: dingtalk_table.TaskRecord) -> str:
    value = record.cells.get("brand")
    if isinstance(value, dict):
        value = value.get("name") or value.get("value") or value.get("text")
    if isinstance(value, list):
        value = value[0] if value else ""
        if isinstance(value, dict):
            value = value.get("name") or value.get("value") or value.get("text")
    return str(value or "").strip()


def _task_products(record: dingtalk_table.TaskRecord) -> list[str]:
    def text_values(value: Any) -> list[str]:
        if isinstance(value, dict):
            for key in ("value", "name", "text"):
                if key in value:
                    return text_values(value[key])
            return []
        if isinstance(value, list):
            values: list[str] = []
            for item in value:
                values.extend(text_values(item))
            return values
        text = str(value or "").strip()
        return re.split(r"[,，、]+", text) if text else []

    products: list[str] = []
    seen: set[str] = set()
    for item in text_values(record.cells.get("productName")):
        product = item.strip()
        if product and product not in seen:
            seen.add(product)
            products.append(product)
    return products


def _task_report_date(record: dingtalk_table.TaskRecord) -> date:
    value = record.cells.get("launchDate")
    if isinstance(value, dict):
        value = value.get("value") or value.get("date") or value.get("text")
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    text = str(value or "").strip()
    if not text:
        raise RuntimeError("“报告日期”为空，无法创建钉钉归档目录。")
    try:
        return datetime.fromisoformat(text.replace("Z", "+00:00")).date()
    except ValueError as exc:
        raise RuntimeError(f"“报告日期”格式无效：{text}") from exc


def _task_folder_name(record: dingtalk_table.TaskRecord) -> str:
    brand = _task_brand(record)
    if not brand:
        raise RuntimeError("“竞品品牌”为空，无法创建钉钉归档目录。")
    products = _task_products(record)
    if not products:
        raise RuntimeError("“关注新品”为空，无法创建钉钉归档目录。")
    clean_brand = re.sub(r'[\\/:*?"<>|\r\n]+', "_", brand).strip()
    clean_products = [
        cleaned
        for item in products
        if (cleaned := re.sub(r'[\\/:*?"<>|\r\n]+', "_", item).strip())
    ]
    if not clean_brand or not clean_products:
        raise RuntimeError("“竞品品牌”或“关注新品”无法生成有效的钉钉文件夹名称。")
    return f"{clean_brand}：{'、'.join(clean_products)}"


def _ensure_local_upload_folder(
    configs: dict[str, dict[str, Any]], record: dingtalk_table.TaskRecord
) -> str:
    report_date = _task_report_date(record)
    folder_id, _ = dingtalk_docs.ensure_local_upload_task_folder(
        configs.get("dingtalk_docs", {}),
        report_date.year,
        report_date.month,
        _task_folder_name(record),
    )
    return folder_id


def _clean_folder_text(value: str) -> str:
    return re.sub(r'[\\/:*?"<>|\r\n]+', "_", value).strip()


def _same_social_day_products(
    configs: dict[str, dict[str, Any]],
    current_record: dingtalk_table.TaskRecord,
    brand: str,
    day30: date,
) -> list[str]:
    products: list[str] = []
    seen: set[str] = set()
    for record in dingtalk_table.fetch_records(configs):
        if _task_brand(record) != brand:
            continue
        try:
            record_day30 = _social_task_date(record, "day30", "30日")
        except RuntimeError:
            continue
        if record_day30 != day30:
            continue
        for product in _task_products(record):
            clean_product = _clean_folder_text(product)
            if clean_product and clean_product not in seen:
                seen.add(clean_product)
                products.append(clean_product)
    for product in _task_products(current_record):
        clean_product = _clean_folder_text(product)
        if clean_product and clean_product not in seen:
            seen.add(clean_product)
            products.append(clean_product)
    return products


def _ensure_social_archive_folder(
    configs: dict[str, dict[str, Any]],
    record: dingtalk_table.TaskRecord,
    day30: date,
) -> str:
    """
    社媒结果按 30 日所在月归档，并优先复用同品牌、同产品文件夹。
    """
    brand = _clean_folder_text(_task_brand(record))
    products = [_clean_folder_text(item) for item in _task_products(record)]
    products = [item for item in products if item]
    if not brand:
        raise RuntimeError("“品牌”为空，无法创建社媒归档目录。")
    if not products:
        raise RuntimeError("“新品”为空，无法创建社媒归档目录。")

    docs_config = configs.get("dingtalk_docs", {})
    month_id, _ = dingtalk_docs.ensure_local_upload_month_folder(
        docs_config, day30.year, day30.month
    )
    matches = [
        (folder_id, name)
        for folder_id, name in dingtalk_docs.child_folders(docs_config, month_id)
        if brand in name and all(product in name for product in products)
    ]
    exact_name = f"{brand}：{'、'.join(products)}"
    exact_matches = [item for item in matches if item[1] == exact_name]
    if len(exact_matches) == 1:
        return exact_matches[0][0]
    if len(matches) == 1:
        return matches[0][0]
    if len(matches) > 1:
        names = "、".join(name for _, name in matches)
        raise RuntimeError(
            f"同一月份存在多个同时包含品牌和产品名的文件夹：{names}。"
        )

    grouped_products = _same_social_day_products(configs, record, _task_brand(record), day30)
    if not grouped_products:
        raise RuntimeError("未找到同品牌、同一“30日”的新品记录。")
    folder_name = f"{brand}：{'、'.join(grouped_products)}"
    folder_id, _ = dingtalk_docs.ensure_child_folder(docs_config, month_id, folder_name)
    return folder_id


def _ensure_dingtalk_input_attachment(
    configs: dict[str, dict[str, Any]],
    record: dingtalk_table.TaskRecord,
    field_key: str,
    field_label: str,
    folder_id: str | None,
) -> str | None:
    value = record.cells.get(field_key)
    if _has_dingtalk_node_link(value):
        items = _value_items(value)
        node_count = sum(
            bool(dingtalk_docs.node_id_from_url(dingtalk_docs.extract_link(item)))
            for item in items
        )
        if len(items) != 1 or node_count != 1:
            raise RuntimeError(f"“{field_label}”字段仅支持一个钉钉附件或文件夹链接。")
        return folder_id or dingtalk_docs.folder_id_for_linked_node(configs.get("dingtalk_docs", {}), value)
    ai_table_attachment = _single_ai_table_attachment(value, field_label)
    local_path = (
        _download_ai_table_attachment(
            configs, record, field_key, field_label, ai_table_attachment
        )
        if ai_table_attachment
        else _single_local_xlsx(value, field_label)
    )
    if not folder_id:
        folder_id = _ensure_local_upload_folder(configs, record)
    local_path = _cache_promoted_input(configs, record, field_key, local_path)
    url = dingtalk_docs.upload_file(configs.get("dingtalk_docs", {}), local_path, folder_id)
    dingtalk_table.update_attachment_fields(configs, record.record_id, {field_key: (local_path.name, url)})
    linked_value = dingtalk_table.attachment_cell(local_path.name, url)
    record.cells[field_key] = linked_value
    return folder_id


def _cache_promoted_input(
    configs: dict[str, dict[str, Any]], record: dingtalk_table.TaskRecord, field_key: str, source: Path
) -> Path:
    task_root = output_root(configs) / record.record_id
    target_dir = task_root / "raw_data" if field_key == "deliveryData" else task_root
    target_dir.mkdir(parents=True, exist_ok=True)
    target = target_dir / source.name
    if source.resolve() != target.resolve():
        shutil.copy2(source, target)
    return target


def _task_annotation_path(configs: dict[str, dict[str, Any]], record: dingtalk_table.TaskRecord) -> Path | None:
    value = record.cells.get("productMenu")
    if not _has_dingtalk_node_link(value):
        raise RuntimeError("“产品清单”字段缺少可识别的钉钉 xlsx 附件。")
    if value:
        downloaded = dingtalk_docs.download_linked_file(
            configs.get("dingtalk_docs", {}),
            value,
            output_root(configs) / record.record_id,
        )
        if downloaded:
            return downloaded
    candidates = sorted((output_root(configs) / record.record_id).glob("*产品清单.xlsx"))
    return candidates[-1] if candidates else None

def _consumer_feedback_configs(configs: dict[str, dict[str, Any]]) -> dict[str, dict[str, Any]]:
    scoped = {key: dict(value) for key, value in configs.items()}
    scoped["dingtalk"] = dict(configs["dingtalk"])
    scoped["dingtalk"]["tableId"] = CONSUMER_FEEDBACK_TABLE_ID
    scoped["field_mapping"] = {
        "fields": {
            "brand": {"fieldId": "01ZM8y7"},
            "productName": {"fieldId": "mHe1U1b"},
            "launchDate": {"fieldId": "mKUEya0"},
            "day30": {"fieldId": "SSrz1N1"},
            "dianping": {"fieldId": "Hkl2xDv"},
            "weibo": {"fieldId": "GGXHqyG"},
            "xiaohongshu": {"fieldId": "NvvlACX"},
            "douyin": {"fieldId": "CtZivbf"},
            "bilibili": {"fieldId": "KfgxCcf"},
            "report": {"fieldId": "amWTton"},
            "feedback": {"fieldId": "j8IgB7P"},
        }
    }
    return scoped


def _social_task_date(record: dingtalk_table.TaskRecord, key: str, label: str) -> date:
    value = record.cells.get(key)
    while isinstance(value, (dict, list)):
        if isinstance(value, list):
            value = value[0] if value else None
        else:
            value = value.get("value") or value.get("date") or value.get("text")
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, (int, float)):
        timestamp = float(value)
        if timestamp > 10_000_000_000:
            timestamp /= 1000
        return datetime.fromtimestamp(timestamp).date()
    text = str(value or "").strip()
    if not text:
        raise RuntimeError(f"“{label}”为空，无法生成社媒评论统计表。")
    normalized = text.replace("年", "-").replace("月", "-").replace("日", "").replace("/", "-")
    try:
        return datetime.fromisoformat(normalized.replace("Z", "+00:00")).date()
    except ValueError as exc:
        raise RuntimeError(f"“{label}”格式无效：{text}") from exc


def _download_linked_social_file(
    configs: dict[str, dict[str, Any]],
    record: dingtalk_table.TaskRecord,
    field_key: str,
    field_label: str,
) -> Path:
    directory = output_root(configs) / record.record_id / "social_inputs" / field_key
    downloaded = dingtalk_docs.download_linked_file(
        configs.get("dingtalk_docs", {}), record.cells.get(field_key), directory
    )
    if not downloaded or downloaded.suffix.lower() != ".xlsx":
        raise RuntimeError(f"“{field_label}”字段不是可下载的 xlsx 附件。")
    return downloaded


def _prepare_social_input(
    configs: dict[str, dict[str, Any]],
    record: dingtalk_table.TaskRecord,
    field_key: str,
    field_label: str,
    folder_id: str | None,
) -> tuple[str | None, Path | None]:
    value = record.cells.get(field_key)
    if not _value_items(value):
        return folder_id, None
    if _has_dingtalk_node_link(value):
        items = _value_items(value)
        node_count = sum(
            bool(dingtalk_docs.node_id_from_url(dingtalk_docs.extract_link(item)))
            for item in items
        )
        if len(items) != 1 or node_count != 1:
            raise RuntimeError(f"“{field_label}”字段仅支持一个钉钉 xlsx 附件。")
        resolved_folder = folder_id or dingtalk_docs.folder_id_for_linked_node(
            configs.get("dingtalk_docs", {}), value
        )
        return resolved_folder, _download_linked_social_file(
            configs, record, field_key, field_label
        )

    attachment = _wait_for_ai_table_attachment_url(
        configs, record, field_key, field_label
    )
    if attachment is None:
        raise RuntimeError(
            f"“{field_label}”字段只支持钉钉文档节点或 AI 表“+”上传的单个 xlsx 附件。"
        )
    local_path = _download_ai_table_attachment(
        configs, record, field_key, field_label, attachment
    )
    if not folder_id:
        folder_id = _ensure_local_upload_folder(configs, record)
    url = dingtalk_docs.upload_file(
        configs.get("dingtalk_docs", {}), local_path, folder_id
    )
    dingtalk_table.update_attachment_fields(
        configs, record.record_id, {field_key: (local_path.name, url)}
    )
    record.cells[field_key] = dingtalk_table.attachment_cell(local_path.name, url)
    return folder_id, local_path


def _existing_field_ids(configs: dict[str, dict[str, Any]]) -> set[str]:
    dingtalk = configs["dingtalk"]
    data = dingtalk_table.call_table_tool(
        dingtalk,
        "get_fields",
        {"baseId": dingtalk["baseId"], "tableId": dingtalk["tableId"]},
    )
    fields = data.get("data", {}).get("fields") or data.get("data") or []
    return {str(field.get("fieldId") or field.get("id")) for field in fields if isinstance(field, dict)}


def _filter_links_for_existing_fields(configs: dict[str, dict[str, Any]], links: dict[str, tuple[str, str]]) -> dict[str, tuple[str, str]]:
    existing = _existing_field_ids(configs)
    return {key: value for key, value in links.items() if dingtalk_table.field_id(configs, key) in existing}


def build_progress_callback(configs: dict[str, dict[str, Any]], record_id: str):
    def update(message: str) -> None:
        try:
            dingtalk_table.update_feedback(configs, record_id, message)
        except Exception as exc:
            print(f"WARNING: 进度回写失败：{message} {type(exc).__name__}: {exc}", file=sys.stderr)

    return update


def _empty_delivery_platforms(outputs: dict[str, Path]) -> list[str]:
    empty: list[str] = []
    for key, platform in DELIVERY_OUTPUT_PLATFORMS:
        path = outputs[key]
        wb = load_workbook(path, read_only=True, data_only=True)
        try:
            has_product = any(
                str(row[0] or "").strip()
                for row in wb.active.iter_rows(min_row=2, min_col=2, max_col=2, values_only=True)
            )
        finally:
            wb.close()
        if not has_product:
            empty.append(platform)
    return empty


def _delivery_completion_message(empty_platforms: list[str]) -> str:
    if not empty_platforms:
        return "已生成外卖数表"
    return f"已生成外卖数表，{'、'.join(empty_platforms)}无数据"


def run_prepare_product_menu(record_id: str) -> dict[str, Any]:
    configs = load_configs()
    progress_callback = build_progress_callback(configs, record_id)
    try:
        dingtalk_table.mark_status(configs, record_id, _status(configs, "productMenuRunning", "产品清单生成中"))
        record = dingtalk_table.fetch_record(configs, record_id)
        folder_id = _task_docs_folder_id(configs, record)
        dingtalk_table.clear_attachment_fields(configs, record_id, ("productMenu",))
        record.cells["productMenu"] = []
        progress_callback("1/4 正在读取外卖数据")
        folder_id = _ensure_dingtalk_input_attachment(configs, record, "deliveryData", "外卖数据", folder_id)
        input_dir = _task_delivery_input_dir(configs, record)
        output = prepare_product_menu(record_id, input_dir, output_root(configs) / record_id, progress_callback=progress_callback)
        progress_callback("4/4 正在上传到钉钉文档")
        url = dingtalk_docs.upload_file(configs.get("dingtalk_docs", {}), output, folder_id)
        dingtalk_table.mark_links(
            configs,
            record_id,
            {"productMenu": (output.name, url)},
            _status(configs, "waitingAnnotation", "待标注"),
        )
        progress_callback("已生成，请人工检查产品清单上新日期标注是否完整。")
        return {"ok": True, "stage": "prepare-product-menu", "recordId": record_id, "output": str(output), "url": url}
    except Exception as exc:
        message = f"{type(exc).__name__}: {exc}"
        try:
            progress_callback(f"生成失败：{message}")
            dingtalk_table.mark_failed(configs, record_id, message)
        except Exception:
            message += "\n\n回写失败状态也失败：\n" + traceback.format_exc()
        raise RuntimeError(message) from exc


def run_generate_delivery_tables(record_id: str) -> dict[str, Any]:
    configs = load_configs()
    progress_callback = build_progress_callback(configs, record_id)
    try:
        dingtalk_table.mark_status(configs, record_id, _status(configs, "dataTablesRunning", "外卖数据统计中"))
        record = dingtalk_table.fetch_record(configs, record_id)
        folder_id = _task_docs_folder_id(configs, record)
        dingtalk_table.clear_attachment_fields(configs, record_id, DELIVERY_OUTPUT_FIELDS)
        for key in DELIVERY_OUTPUT_FIELDS:
            record.cells[key] = []
        progress_callback("1/3 正在下载外卖数据和产品清单")
        folder_id = _ensure_dingtalk_input_attachment(configs, record, "deliveryData", "外卖数据", folder_id)
        folder_id = _ensure_dingtalk_input_attachment(configs, record, "productMenu", "产品清单", folder_id)
        if not folder_id:
            raise RuntimeError("无法确定钉钉目标文件夹，请确认“外卖数据”或“产品清单”字段中的附件有效。")
        input_dir = _task_delivery_input_dir(configs, record)
        annotation_path = _task_annotation_path(configs, record)
        if not annotation_path or not annotation_path.exists():
            raise RuntimeError("找不到产品清单及上新日期，请先点击“提取产品清单”并完成标注。")
        progress_callback("2/3 正在生成外卖数表")
        outputs = generate_delivery_tables(record_id, input_dir, annotation_path)
        empty_platforms = _empty_delivery_platforms(outputs)
        empty_keys = {
            key for key, platform in DELIVERY_OUTPUT_PLATFORMS if platform in empty_platforms
        }
        generated_outputs = {
            key: path for key, path in outputs.items() if key not in empty_keys
        }
        for key in empty_keys:
            path = outputs.get(key)
            if path and path.exists():
                path.unlink()
        progress_callback("3/3 正在上传外卖数表到钉钉文档")
        links: dict[str, tuple[str, str]] = {}
        for key, path in generated_outputs.items():
            links[key] = (path.name, dingtalk_docs.upload_file(configs.get("dingtalk_docs", {}), path, folder_id))
        writable_links = _filter_links_for_existing_fields(configs, links) if links else {}
        dingtalk_table.mark_links(
            configs,
            record_id,
            writable_links,
            _status(configs, "waitingReport", "外卖数据已生成"),
        )
        progress_callback(_delivery_completion_message(empty_platforms))
        return {
            "ok": True,
            "stage": "generate-delivery-tables",
            "recordId": record_id,
            "outputs": {key: str(path) for key, path in generated_outputs.items()},
            "emptyPlatforms": empty_platforms,
        }
    except Exception as exc:
        message = f"{type(exc).__name__}: {exc}"
        try:
            dingtalk_table.mark_failed(configs, record_id, message)
        except Exception:
            message += "\n\n回写失败状态也失败：\n" + traceback.format_exc()
        raise RuntimeError(message) from exc


def run_generate_consumer_feedback_tables(record_id: str) -> dict[str, Any]:
    configs = _consumer_feedback_configs(load_configs())
    progress_callback = build_progress_callback(configs, record_id)
    try:
        record = dingtalk_table.fetch_record(configs, record_id)
        brand = _task_brand(record)
        products = _task_products(record)
        if not brand:
            raise RuntimeError("“品牌”为空，无法生成社媒评论统计表。")
        if not products:
            raise RuntimeError("“新品”为空，无法生成社媒评论统计表。")
        product = "、".join(products)
        launch_date = _social_task_date(record, "launchDate", "上市日期")
        day30 = _social_task_date(record, "day30", "30日")
        if day30 < launch_date:
            raise RuntimeError("“30日”早于“上市日期”，无法生成社媒评论统计表。")

        dingtalk_table.clear_link_fields(configs, record_id, ("report",))
        record.cells["report"] = ""
        progress_callback("1/3 正在下载并归档社媒数据")
        folder_id = _ensure_social_archive_folder(configs, record, day30)
        platform_files: dict[str, Path] = {}
        for key, label in SOCIAL_INPUT_PLATFORMS:
            folder_id, path = _prepare_social_input(
                configs, record, key, label, folder_id
            )
            if path:
                platform_files[key] = path
        progress_callback("2/3 正在生成社媒评论统计表")
        output = generate_consumer_feedback_tables(
            record_id,
            platform_files,
            brand=brand,
            product=product,
            start_date=f"{launch_date.month}.{launch_date.day}",
            end_date=f"{day30.month}.{day30.day}",
        )
        progress_callback("3/3 正在上传统计表到钉钉文档")
        url = dingtalk_docs.upload_file(
            configs.get("dingtalk_docs", {}), output, folder_id
        )
        dingtalk_table.mark_links(
            configs, record_id, {"report": (output.name, url)}
        )
        progress_callback("已生成社媒评论统计表")
        return {
            "ok": True,
            "stage": "generate-consumer-feedback-tables",
            "recordId": record_id,
            "output": str(output),
            "url": url,
        }
    except Exception as exc:
        message = f"{type(exc).__name__}: {exc}"
        try:
            progress_callback(f"生成失败：{message}")
            dingtalk_table.mark_failed(configs, record_id, message)
        except Exception:
            message += "\n\n回写失败状态也失败：\n" + traceback.format_exc()
        raise RuntimeError(message) from exc


def run_generate_report(record_id: str) -> dict[str, Any]:
    configs = load_configs()
    try:
        dingtalk_table.mark_status(configs, record_id, _status(configs, "reportRunning", "报告生成中"))
        record = dingtalk_table.fetch_record(configs, record_id)
        brand = _task_brand(record) or "品牌"
        products = _task_products(record)
        output = generate_report(record_id, brand, products)
        url = dingtalk_docs.upload_file(configs.get("dingtalk_docs", {}), output, _task_docs_folder_id(configs, record))
        dingtalk_table.mark_links(configs, record_id, {"report": (output.name, url)}, _status(configs, "done", "已生成"))
        return {"ok": True, "stage": "generate-report", "recordId": record_id, "output": str(output), "url": url}
    except Exception as exc:
        message = f"{type(exc).__name__}: {exc}"
        try:
            dingtalk_table.mark_failed(configs, record_id, message)
        except Exception:
            message += "\n\n回写失败状态也失败：\n" + traceback.format_exc()
        raise RuntimeError(message) from exc
