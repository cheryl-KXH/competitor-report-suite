#!/usr/bin/env python3
from __future__ import annotations

import argparse
import re
import sys
from datetime import date, datetime
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[2]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from service import dingtalk_table
from service.config import load_configs, output_root
from scripts.reporting.html import (
    DeliveryReport,
    JdSale,
    ProductInfo,
    ReportBuildResult,
    SocialReport,
    build_report_html,
    normalize_product_key,
)


def _safe_filename(value: str) -> str:
    return re.sub(r'[\\/*?"<>|\r\n]+', "_", value).strip() or "竞品跟踪反馈报告"


def _select_name(value: Any) -> str:
    if isinstance(value, dict):
        return str(value.get("name") or value.get("value") or value.get("text") or "").strip()
    if isinstance(value, list):
        return _select_name(value[0]) if value else ""
    return str(value or "").strip()


def _plain_text(value: Any) -> str:
    if isinstance(value, dict):
        value = value.get("markdown") or value.get("text") or value.get("value") or value.get("name") or ""
    if isinstance(value, list):
        return "、".join(part for item in value if (part := _plain_text(item)))
    text = str(value or "").replace("\r\n", "\n").replace("\r", "\n")
    return "\n".join(line.strip() for line in text.split("\n") if line.strip())


def _attachment_image_urls(value: Any) -> tuple[str, ...]:
    items = value if isinstance(value, list) else [value]
    urls: list[str] = []
    for item in items:
        if not isinstance(item, dict):
            continue
        nested = item.get("files") or item.get("attachments")
        if isinstance(nested, list):
            urls.extend(_attachment_image_urls(nested))
            continue
        filename = str(item.get("filename") or item.get("fileName") or item.get("name") or "").lower()
        file_type = str(item.get("type") or item.get("mimeType") or item.get("contentType") or "").lower()
        is_image = file_type == "image" or file_type.startswith("image/") or filename.endswith(
            (".png", ".jpg", ".jpeg", ".webp", ".gif", ".bmp")
        )
        if not is_image:
            continue
        for key in ("url", "resourceUrl", "downloadUrl", "previewUrl", "thumbnailUrl"):
            if item.get(key):
                urls.append(str(item[key]))
                break
    return tuple(dict.fromkeys(urls))


def _date_value(value: Any) -> date | None:
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
        return None
    try:
        return datetime.fromisoformat(text.replace("Z", "+00:00")).date()
    except ValueError:
        try:
            return date.fromisoformat(text[:10])
        except ValueError:
            return None


def _query_product_source(configs: dict[str, dict[str, Any]]) -> list[dict[str, Any]]:
    source = configs.get("report_rules", {}).get("productInfoSource", {})
    base_id = str(source.get("baseId") or "").strip()
    table_id = str(source.get("tableId") or "").strip()
    if not base_id or not table_id:
        raise RuntimeError("report_rules.json 未配置 productInfoSource")
    fields = source.get("fields", {})
    field_ids = [
        str(source.get("brandFieldId") or "01ZM8y7"),
        str(source.get("launchDateFieldId") or "mKUEya0"),
        *[str(meta.get("fieldId")) for meta in fields.values() if meta.get("fieldId")],
    ]
    records: list[dict[str, Any]] = []
    cursor = ""
    seen: set[str] = set()
    while True:
        payload: dict[str, Any] = {
            "baseId": base_id,
            "tableId": table_id,
            "limit": 100,
            "fieldIds": list(dict.fromkeys(field_ids)),
        }
        if cursor:
            payload["cursor"] = cursor
        response = dingtalk_table.call_table_tool(configs["dingtalk"], "query_records", payload)
        data = response.get("data", {}) if isinstance(response, dict) else {}
        batch = data.get("records", []) if isinstance(data, dict) else []
        records.extend(batch or [])
        next_cursor = str(data.get("nextCursor") or "") if isinstance(data, dict) else ""
        if not next_cursor or next_cursor in seen or not batch:
            break
        seen.add(next_cursor)
        cursor = next_cursor
    return records


def load_product_infos(
    configs: dict[str, dict[str, Any]],
    brand: str,
    products: list[str],
    launch_dates: dict[str, date | None],
) -> tuple[dict[str, ProductInfo], list[str]]:
    source = configs.get("report_rules", {}).get("productInfoSource", {})
    fields = source.get("fields", {})
    brand_field = str(source.get("brandFieldId") or "01ZM8y7")
    launch_field = str(source.get("launchDateFieldId") or "mKUEya0")
    product_field = str(fields.get("productName", {}).get("fieldId") or "mHe1U1b")
    records = _query_product_source(configs)
    infos: dict[str, ProductInfo] = {}
    warnings: list[str] = []
    for product in products:
        target = normalize_product_key(product)
        candidates: list[tuple[dict[str, Any], date | None]] = []
        for record in records:
            cells = record.get("cells", {}) if isinstance(record, dict) else {}
            if _select_name(cells.get(brand_field)) != brand:
                continue
            if normalize_product_key(cells.get(product_field)) != target:
                continue
            candidates.append((cells, _date_value(cells.get(launch_field))))
        expected_date = launch_dates.get(product)
        dated = [item for item in candidates if expected_date and item[1] == expected_date]
        selected = dated or candidates
        if not selected:
            warnings.append(f"{product}：周报源表未匹配到产品信息")
            continue
        if len(selected) > 1:
            warnings.append(f"{product}：周报源表匹配到多条产品信息，已使用第一条")
        cells = selected[0][0]
        infos[product] = ProductInfo(
            product_name=_plain_text(cells.get(product_field)) or product,
            series=_plain_text(cells.get(fields.get("series", {}).get("fieldId"))),
            selling_point=_plain_text(cells.get(fields.get("sellingPoint", {}).get("fieldId"))),
            price=_plain_text(cells.get(fields.get("price", {}).get("fieldId"))),
            ingredients=_plain_text(cells.get(fields.get("ingredients", {}).get("fieldId"))),
            image_urls=_attachment_image_urls(cells.get(fields.get("appearanceImages", {}).get("fieldId"))),
        )
    return infos, warnings


def generate_report(
    record_id: str,
    brand: str,
    products: list[str],
    *,
    report_date: date,
    meituan_path: Path | None = None,
    eleme_path: Path | None = None,
    jd_path: Path | None = None,
    social_paths: dict[str, Path] | None = None,
    launch_dates: dict[str, date | None] | None = None,
    output_dir: Path | None = None,
    configs: dict[str, dict[str, Any]] | None = None,
    delivery_report: DeliveryReport | None = None,
    jd_report: tuple[tuple[JdSale, ...], float] | None = None,
    social_reports: dict[str, SocialReport] | None = None,
) -> ReportBuildResult:
    configs = configs or load_configs()
    launch_dates = launch_dates or {}
    social_paths = social_paths or {}
    product_warnings: list[str] = []
    try:
        product_infos, product_warnings = load_product_infos(
            configs, brand, products, launch_dates
        )
    except Exception as exc:
        product_infos = {}
        product_warnings = [f"周报产品信息读取失败（{exc}）"]
    out_dir = output_dir or output_root(configs) / record_id
    filename = _safe_filename(f"{brand}：{'、'.join(products)} {report_date:%Y%m%d}.html")
    result = build_report_html(
        title=configs.get("report_rules", {}).get("reportTitle", "竞品跟踪反馈报告"),
        brand=brand,
        products=products,
        report_date=report_date,
        meituan_path=meituan_path,
        eleme_path=eleme_path,
        jd_path=jd_path,
        social_paths=social_paths,
        product_infos=product_infos,
        launch_dates=launch_dates,
        output_path=out_dir / filename,
        configs=configs,
        delivery_report=delivery_report,
        jd_report=jd_report,
        social_report_models=social_reports,
    )
    return ReportBuildResult(result.path, tuple(dict.fromkeys([*product_warnings, *result.warnings])))


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="生成 HTML 竞品跟踪反馈报告")
    parser.add_argument("--record-id", required=True)
    parser.add_argument("--brand", required=True)
    parser.add_argument("--products", required=True, help="多个产品用英文或中文逗号分隔")
    parser.add_argument("--report-date", required=True, help="YYYY-MM-DD")
    parser.add_argument("--meituan", required=True)
    parser.add_argument("--eleme", required=True)
    parser.add_argument("--jd")
    parser.add_argument("--output-dir")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    products = [item.strip() for item in re.split(r"[,，、]+", args.products) if item.strip()]
    result = generate_report(
        args.record_id,
        args.brand,
        products,
        report_date=date.fromisoformat(args.report_date),
        meituan_path=Path(args.meituan),
        eleme_path=Path(args.eleme),
        jd_path=Path(args.jd) if args.jd else None,
        output_dir=Path(args.output_dir) if args.output_dir else None,
    )
    print(f"已生成：{result.path}")
    for warning in result.warnings:
        print(f"WARNING: {warning}", file=sys.stderr)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
