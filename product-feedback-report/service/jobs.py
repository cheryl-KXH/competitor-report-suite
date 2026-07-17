from __future__ import annotations

import traceback
import sys
from pathlib import Path
from typing import Any

from service import dingtalk_docs, dingtalk_table
from service.config import load_configs, output_root
from scripts.generate_consumer_feedback_tables import generate_consumer_feedback_tables
from scripts.generate_delivery_tables import generate_delivery_tables
from scripts.generate_report import generate_report
from scripts.prepare_product_menu import prepare_product_menu


CONSUMER_FEEDBACK_TABLE_ID = "hvcu7Bw"


def _status(configs: dict[str, dict[str, Any]], key: str, fallback: str) -> str:
    return configs.get("report_rules", {}).get("statuses", {}).get(key, fallback)


def _task_input_dir(configs: dict[str, dict[str, Any]], record: dingtalk_table.TaskRecord) -> Path:
    value = record.cells.get("deliveryData") or record.cells.get("allData")
    raw_dir = output_root(configs) / record.record_id / "raw_data"
    linked_dir = dingtalk_docs.download_linked_folder(configs.get("dingtalk_docs", {}), value, raw_dir)
    if linked_dir and linked_dir.is_dir():
        return linked_dir
    raise RuntimeError("“外卖数据”字段为空，或不是可识别的本地文件夹/钉钉文档文件夹附件。")


def _task_docs_folder_id(configs: dict[str, dict[str, Any]], record: dingtalk_table.TaskRecord) -> str | None:
    docs_config = configs.get("dingtalk_docs", {})
    for key in ("deliveryData", "productMenu", "allData"):
        value = record.cells.get(key)
        if not value:
            continue
        folder_id = dingtalk_docs.folder_id_for_linked_node(docs_config, value)
        if folder_id:
            return folder_id
    return None


def _task_annotation_path(configs: dict[str, dict[str, Any]], record: dingtalk_table.TaskRecord) -> Path | None:
    value = record.cells.get("productMenu")
    local = dingtalk_docs.local_path_from_link(value)
    if local and local.is_file():
        return local
    if value:
        downloaded = dingtalk_docs.download_linked_file(
            configs.get("dingtalk_docs", {}),
            value,
            output_root(configs) / record.record_id / "product_menu",
        )
        if downloaded:
            return downloaded
    candidates = sorted((output_root(configs) / record.record_id).glob("*产品清单.xlsx"))
    return candidates[-1] if candidates else None


def _task_products(record: dingtalk_table.TaskRecord) -> list[str]:
    value = record.cells.get("productName")
    if isinstance(value, list):
        return [str(item).strip() for item in value if str(item).strip()]
    text = str(value or "").strip()
    if not text:
        return []
    return [item.strip() for item in text.replace("，", ",").split(",") if item.strip()]


def _consumer_feedback_configs(configs: dict[str, dict[str, Any]]) -> dict[str, dict[str, Any]]:
    scoped = {key: dict(value) for key, value in configs.items()}
    scoped["dingtalk"] = dict(configs["dingtalk"])
    scoped["dingtalk"]["tableId"] = CONSUMER_FEEDBACK_TABLE_ID
    scoped["field_mapping"] = {
        "fields": {
            "brand": {"fieldId": "01ZM8y7"},
            "productName": {"fieldId": "mHe1U1b"},
            "allData": {"fieldId": "iycnz5g"},
            "dianping": {"fieldId": "Hkl2xDv"},
            "weibo": {"fieldId": "GGXHqyG"},
            "xiaohongshu": {"fieldId": "NvvlACX"},
            "douyin": {"fieldId": "CtZivbf"},
            "bilibili": {"fieldId": "KfgxCcf"},
            "status": {"fieldId": "j8IgB7P"},
        }
    }
    return scoped


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


def run_prepare_product_menu(record_id: str) -> dict[str, Any]:
    configs = load_configs()
    progress_callback = build_progress_callback(configs, record_id)
    try:
        dingtalk_table.mark_status(configs, record_id, _status(configs, "productMenuRunning", "产品清单生成中"))
        progress_callback("1/4 正在读取外卖数据")
        record = dingtalk_table.fetch_record(configs, record_id)
        input_dir = _task_input_dir(configs, record)
        output = prepare_product_menu(record_id, input_dir, output_root(configs) / record_id, progress_callback=progress_callback)
        progress_callback("4/4 正在上传到钉钉文档")
        url = dingtalk_docs.upload_file(configs.get("dingtalk_docs", {}), output, _task_docs_folder_id(configs, record))
        dingtalk_table.mark_links(
            configs,
            record_id,
            {"productMenu": (output.name, url)},
            _status(configs, "waitingAnnotation", "待标注"),
        )
        progress_callback(f"已生成：{output.name}")
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
    try:
        dingtalk_table.mark_status(configs, record_id, _status(configs, "dataTablesRunning", "外卖数据统计中"))
        record = dingtalk_table.fetch_record(configs, record_id)
        folder_id = _task_docs_folder_id(configs, record)
        if not folder_id:
            raise RuntimeError("无法确定钉钉目标文件夹，请确认“外卖数据”或“产品清单”字段链接指向钉钉文件夹内的文件。")
        input_dir = _task_input_dir(configs, record)
        annotation_path = _task_annotation_path(configs, record)
        if not annotation_path or not annotation_path.exists():
            raise RuntimeError("找不到产品清单及上新日期，请先点击“提取产品清单”并完成标注。")
        outputs = generate_delivery_tables(record_id, input_dir, annotation_path)
        links: dict[str, tuple[str, str]] = {}
        for key, path in outputs.items():
            links[key] = (path.name, dingtalk_docs.upload_file(configs.get("dingtalk_docs", {}), path, folder_id))
        writable_links = _filter_links_for_existing_fields(configs, links)
        if writable_links:
            dingtalk_table.mark_links(configs, record_id, writable_links, _status(configs, "waitingReport", "外卖数据已生成"))
        else:
            dingtalk_table.update_feedback(configs, record_id, "已生成：" + "、".join(path.name for path in outputs.values()))
        return {"ok": True, "stage": "generate-delivery-tables", "recordId": record_id, "outputs": {k: str(v) for k, v in outputs.items()}}
    except Exception as exc:
        message = f"{type(exc).__name__}: {exc}"
        try:
            dingtalk_table.mark_failed(configs, record_id, message)
        except Exception:
            message += "\n\n回写失败状态也失败：\n" + traceback.format_exc()
        raise RuntimeError(message) from exc


def run_generate_consumer_feedback_tables(record_id: str) -> dict[str, Any]:
    configs = _consumer_feedback_configs(load_configs())
    try:
        dingtalk_table.mark_status(configs, record_id, "消费者反馈统计中")
        record = dingtalk_table.fetch_record(configs, record_id)
        input_dir = _task_input_dir(configs, record)
        outputs = generate_consumer_feedback_tables(record_id, input_dir)
        links: dict[str, tuple[str, str]] = {}
        folder_id = _task_docs_folder_id(configs, record)
        for key, path in outputs.items():
            links[key] = (path.name, dingtalk_docs.upload_file(configs.get("dingtalk_docs", {}), path, folder_id))
        dingtalk_table.mark_links(configs, record_id, links, "消费者反馈已生成")
        return {"ok": True, "stage": "generate-consumer-feedback-tables", "recordId": record_id, "outputs": {k: str(v) for k, v in outputs.items()}}
    except Exception as exc:
        message = f"{type(exc).__name__}: {exc}"
        try:
            dingtalk_table.mark_failed(configs, record_id, message)
        except Exception:
            message += "\n\n回写失败状态也失败：\n" + traceback.format_exc()
        raise RuntimeError(message) from exc


def run_generate_report(record_id: str) -> dict[str, Any]:
    configs = load_configs()
    try:
        dingtalk_table.mark_status(configs, record_id, _status(configs, "reportRunning", "报告生成中"))
        record = dingtalk_table.fetch_record(configs, record_id)
        brand = str(record.cells.get("brand") or "").strip() or "品牌"
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
