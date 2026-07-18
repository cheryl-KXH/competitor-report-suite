#!/usr/bin/env python3
from __future__ import annotations

import argparse
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from service.config import load_configs, output_root
from scripts.delivery.processing import find_delivery_rows, generate_jd_summary, generate_platform_delivery_summary, read_annotation
from scripts.delivery.prepare_product_menu import _crawl_date, _primary_brand, _safe_filename


def generate_delivery_tables(record_id: str, input_dir: Path, annotation_path: Path, output_dir: Path | None = None) -> dict[str, Path]:
    configs = load_configs()
    out_dir = output_dir or output_root(configs) / record_id / "delivery_tables"
    out_dir.mkdir(parents=True, exist_ok=True)
    raw_rows = find_delivery_rows(input_dir)
    if not raw_rows:
        raise RuntimeError(f"未在目录中找到外卖原始数据：{input_dir}")
    annotations = read_annotation(annotation_path)
    filename_prefix = _safe_filename(
        f"{_primary_brand(raw_rows)}-{_crawl_date(raw_rows):%Y%m%d}"
    )
    return {
        "meituanData": generate_platform_delivery_summary(raw_rows, annotations, "美团", out_dir / f"{filename_prefix}-美团.xlsx"),
        "elemeData": generate_platform_delivery_summary(raw_rows, annotations, "饿了么", out_dir / f"{filename_prefix}-饿了么.xlsx"),
        "jdData": generate_jd_summary(raw_rows, annotations, out_dir / f"{filename_prefix}-京东.xlsx"),
    }


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="统计外卖销售数据数表")
    parser.add_argument("--record-id", required=True)
    parser.add_argument("--input-dir", required=True, help="本地外卖数据文件夹")
    parser.add_argument("--annotation", required=True, help="产品清单及上新日期标注表")
    parser.add_argument("--output-dir")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    outputs = generate_delivery_tables(
        args.record_id,
        Path(args.input_dir),
        Path(args.annotation),
        Path(args.output_dir) if args.output_dir else None,
    )
    for key, path in outputs.items():
        print(f"{key}: {path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
