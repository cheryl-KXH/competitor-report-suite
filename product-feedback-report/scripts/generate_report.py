#!/usr/bin/env python3
from __future__ import annotations

import argparse
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from service.config import load_configs, output_root
from scripts.report_html import build_report_html


SOCIAL_LABELS = {
    "dianping": "大众点评",
    "weibo": "微博",
    "xiaohongshu": "小红书",
    "douyin": "抖音",
    "bilibili": "B站",
}


def generate_report(
    record_id: str,
    brand: str,
    products: list[str],
    data_dir: Path | None = None,
    output_dir: Path | None = None,
) -> Path:
    configs = load_configs()
    base = output_root(configs) / record_id
    data_root = data_dir or base / "data_tables"
    out_dir = output_dir or base
    social_paths = {
        label: data_root / f"{label}.xlsx"
        for label in SOCIAL_LABELS.values()
        if (data_root / f"{label}.xlsx").exists()
    }
    return build_report_html(
        title=configs.get("report_rules", {}).get("reportTitle", "竞品跟踪反馈报告"),
        brand=brand,
        products=products,
        delivery_path=data_root / "美团&饿了么外卖数据.xlsx",
        jd_path=data_root / "京东外卖数据.xlsx",
        social_paths=social_paths,
        output_path=out_dir / "竞品跟踪反馈报告.html",
    )


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="生成 HTML 竞品跟踪反馈报告")
    parser.add_argument("--record-id", required=True)
    parser.add_argument("--brand", default="")
    parser.add_argument("--products", default="", help="多个产品用英文逗号分隔")
    parser.add_argument("--data-dir")
    parser.add_argument("--output-dir")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    products = [item.strip() for item in args.products.split(",") if item.strip()]
    output = generate_report(
        args.record_id,
        args.brand or "品牌",
        products,
        Path(args.data_dir) if args.data_dir else None,
        Path(args.output_dir) if args.output_dir else None,
    )
    print(f"已生成：{output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
