#!/usr/bin/env python3
from __future__ import annotations

import argparse
import sys
from collections import Counter
from datetime import date, timedelta
from pathlib import Path
from typing import Any, Callable

ROOT = Path(__file__).resolve().parents[2]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from service import dingtalk_table
from service.config import load_configs, output_root
from scripts.delivery.processing import RECENT_LAUNCH_DAYS, RawSaleRow, find_delivery_rows, normalize_text, parse_date, write_product_menu


def _extract_select_name(value: Any) -> str:
    if isinstance(value, dict):
        return normalize_text(value.get("name"))
    return normalize_text(value)


def _safe_filename(value: str) -> str:
    return "".join("_" if char in '\\/:*?"<>|' else char for char in value).strip() or "产品清单"


def _primary_brand(rows: list[RawSaleRow]) -> str:
    counts = Counter(row.brand for row in rows if row.brand)
    return counts.most_common(1)[0][0] if counts else "品牌"


def _crawl_date(rows: list[RawSaleRow]) -> date:
    return max(row.crawl_date for row in rows)


def _minimum_crawl_date(rows: list[RawSaleRow]) -> date:
    return min(row.crawl_date for row in rows)


def _product_source_fields(
    configs: dict[str, dict[str, Any]],
) -> tuple[str, str, str, str, str]:
    source = configs.get("report_rules", {}).get("productInfoSource", {})
    base_id = str(source.get("baseId") or "").strip()
    table_id = str(source.get("tableId") or "").strip()
    fields = source.get("fields", {})
    brand_field = str(source.get("brandFieldId") or "01ZM8y7")
    product_field = str(fields.get("productName", {}).get("fieldId") or "mHe1U1b")
    launch_field = str(source.get("launchDateFieldId") or "mKUEya0")
    return base_id, table_id, brand_field, product_field, launch_field


def _brand_option_ids(configs: dict[str, dict[str, Any]]) -> dict[str, str]:
    base_id, table_id, brand_field, _, _ = _product_source_fields(configs)
    if not base_id or not table_id:
        return {}
    payload = {"baseId": base_id, "tableId": table_id, "fieldIds": [brand_field]}
    response = dingtalk_table.call_table_tool(configs["dingtalk"], "get_fields", payload)
    fields = response.get("data", {}).get("fields", []) or []
    for field in fields:
        if str(field.get("fieldId") or "") != brand_field:
            continue
        options = (field.get("config") or {}).get("options", []) or []
        return {
            normalize_text(option.get("name")): str(option.get("id") or "").strip()
            for option in options
            if normalize_text(option.get("name")) and str(option.get("id") or "").strip()
        }
    return {}


def _product_source_records(
    configs: dict[str, dict[str, Any]],
    brands: set[str],
    start_date: date,
    report_date: date,
) -> list[dict[str, Any]]:
    base_id, table_id, brand_field, product_field, launch_field = (
        _product_source_fields(configs)
    )
    if not base_id or not table_id or not brands:
        return []

    brand_options = _brand_option_ids(configs)
    records: list[dict[str, Any]] = []
    for brand in sorted(brands):
        brand_option_id = brand_options.get(brand)
        if not brand_option_id:
            continue
        cursor = ""
        seen_cursors: set[str] = set()
        while True:
            payload: dict[str, Any] = {
                "baseId": base_id,
                "tableId": table_id,
                "limit": 100,
                "fieldIds": [brand_field, product_field, launch_field],
                "filters": {
                    "operator": "and",
                    "operands": [
                        {
                            "operator": "eq",
                            "operands": [brand_field, brand_option_id],
                        },
                        {
                            "operator": "not_before",
                            "operands": [launch_field, start_date.isoformat()],
                        },
                        {
                            "operator": "not_after",
                            "operands": [launch_field, report_date.isoformat()],
                        },
                    ],
                },
            }
            if cursor:
                payload["cursor"] = cursor
            response = dingtalk_table.call_table_tool(
                configs["dingtalk"], "query_records", payload
            )
            data = response.get("data", {}) if isinstance(response, dict) else {}
            page_records = data.get("records", []) if isinstance(data, dict) else []
            records.extend(page_records or [])
            next_cursor = (
                str(data.get("nextCursor") or "").strip()
                if isinstance(data, dict)
                else ""
            )
            if not next_cursor or not page_records:
                break
            if next_cursor in seen_cursors:
                raise RuntimeError("钉钉产品资料源分页返回了重复 cursor。")
            seen_cursors.add(next_cursor)
            cursor = next_cursor
    return records


def recent_launch_dates(
    configs: dict[str, dict[str, Any]], rows: list[RawSaleRow], report_date: date
) -> dict[tuple[str, str], date]:
    _, _, brand_field, product_field, launch_field = _product_source_fields(configs)
    target_products = {(row.brand, row.product) for row in rows}
    brands = {row.brand for row in rows if row.brand}
    start_date = _minimum_crawl_date(rows) - timedelta(days=RECENT_LAUNCH_DAYS - 1)
    if report_date < start_date:
        raise RuntimeError(
            f"报告日期 {report_date:%Y-%m-%d} 早于新品查询起始日期 {start_date:%Y-%m-%d}。"
        )
    matched: dict[tuple[str, str], date] = {}
    for record in _product_source_records(configs, brands, start_date, report_date):
        cells = record.get("cells", {})
        brand = _extract_select_name(cells.get(brand_field))
        product = normalize_text(cells.get(product_field))
        launch_date = parse_date(cells.get(launch_field))
        key = (brand, product)
        if key not in target_products or not launch_date:
            continue
        if start_date <= launch_date <= report_date:
            matched[key] = launch_date
    return matched


def prepare_product_menu(
    record_id: str,
    input_dir: Path,
    output_dir: Path | None = None,
    progress_callback: Callable[[str], None] | None = None,
    *,
    report_date: date,
) -> Path:
    configs = load_configs()
    out_dir = output_dir or output_root(configs) / record_id
    if progress_callback:
        progress_callback("2/4 正在提取产品清单")
    rows = find_delivery_rows(input_dir)
    if not rows:
        raise RuntimeError(f"未在目录中找到外卖原始数据：{input_dir}")
    if progress_callback:
        progress_callback("3/4 正在标记上新不满30天的产品")
    launch_dates = recent_launch_dates(configs, rows, report_date)
    filename = _safe_filename(f"{_primary_brand(rows)}-{_crawl_date(rows):%Y%m%d}-产品清单.xlsx")
    return write_product_menu(rows, out_dir / filename, launch_dates)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="生成产品清单及上新日期标注表")
    parser.add_argument("--record-id", required=True)
    parser.add_argument("--input-dir", required=True, help="本地原始数据文件夹")
    parser.add_argument("--output-dir")
    parser.add_argument(
        "--report-date",
        required=True,
        type=date.fromisoformat,
        help="报告日期，格式 YYYY-MM-DD",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    output = prepare_product_menu(
        args.record_id,
        Path(args.input_dir),
        Path(args.output_dir) if args.output_dir else None,
        report_date=args.report_date,
    )
    print(f"已生成：{output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
