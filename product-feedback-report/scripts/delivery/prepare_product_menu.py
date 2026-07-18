#!/usr/bin/env python3
from __future__ import annotations

import argparse
import sys
from collections import Counter
from datetime import date
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


def _product_source_records(configs: dict[str, dict[str, Any]]) -> list[dict[str, Any]]:
    source = configs.get("report_rules", {}).get("productInfoSource", {})
    base_id = str(source.get("baseId") or "").strip()
    table_id = str(source.get("tableId") or "").strip()
    fields = source.get("fields", {})
    brand_field = "01ZM8y7"
    product_field = fields.get("productName", {}).get("fieldId") or "mHe1U1b"
    launch_field = "mKUEya0"
    if not base_id or not table_id:
        return []

    records: list[dict[str, Any]] = []
    cursor = None
    while True:
        payload: dict[str, Any] = {
            "baseId": base_id,
            "tableId": table_id,
            "limit": 100,
            "fieldIds": [brand_field, product_field, launch_field],
        }
        if cursor:
            payload["cursor"] = cursor
        data = dingtalk_table.call_table_tool(configs["dingtalk"], "query_records", payload)
        page_records = data.get("data", {}).get("records", []) or []
        records.extend(page_records)
        cursor = data.get("data", {}).get("nextCursor")
        if not cursor or not page_records:
            break
    return records


def recent_launch_dates(configs: dict[str, dict[str, Any]], rows: list[RawSaleRow]) -> dict[tuple[str, str], date]:
    source = configs.get("report_rules", {}).get("productInfoSource", {})
    fields = source.get("fields", {})
    product_field = fields.get("productName", {}).get("fieldId") or "mHe1U1b"
    launch_field = "mKUEya0"
    target_products = {(row.brand, row.product) for row in rows}
    crawl_date = _crawl_date(rows)
    matched: dict[tuple[str, str], date] = {}
    for record in _product_source_records(configs):
        cells = record.get("cells", {})
        brand = _extract_select_name(cells.get("01ZM8y7"))
        product = normalize_text(cells.get(product_field))
        launch_date = parse_date(cells.get(launch_field))
        key = (brand, product)
        if key not in target_products or not launch_date:
            continue
        days = (crawl_date - launch_date).days + 1
        if 1 <= days <= RECENT_LAUNCH_DAYS:
            matched[key] = launch_date
    return matched


def prepare_product_menu(
    record_id: str,
    input_dir: Path,
    output_dir: Path | None = None,
    progress_callback: Callable[[str], None] | None = None,
) -> Path:
    configs = load_configs()
    out_dir = output_dir or output_root(configs) / record_id
    if progress_callback:
        progress_callback("2/4 正在提取产品清单")
    rows = find_delivery_rows(input_dir)
    if not rows:
        raise RuntimeError(f"未在目录中找到外卖原始数据：{input_dir}")
    if progress_callback:
        progress_callback("3/4 正在标记上新不满32天的产品")
    launch_dates = recent_launch_dates(configs, rows)
    filename = _safe_filename(f"{_primary_brand(rows)}-{_crawl_date(rows):%Y%m%d}-产品清单.xlsx")
    return write_product_menu(rows, out_dir / filename, launch_dates)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="生成产品清单及上新日期标注表")
    parser.add_argument("--record-id", required=True)
    parser.add_argument("--input-dir", required=True, help="本地原始数据文件夹")
    parser.add_argument("--output-dir")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    output = prepare_product_menu(args.record_id, Path(args.input_dir), Path(args.output_dir) if args.output_dir else None)
    print(f"已生成：{output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
