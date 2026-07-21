#!/usr/bin/env python3
from __future__ import annotations

import argparse
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from service.config import load_configs, output_root
from scripts.social.processing import (
    build_social_feedback_workbook,
    summarize_social_cleaned_workbook,
)


def generate_consumer_feedback_tables(
    record_id: str,
    cleaned_raw_data: Path,
    *,
    brand: str,
    product: str,
    start_date: str,
    end_date: str,
    output_dir: Path | None = None,
) -> Path:
    configs = load_configs()
    out_dir = output_dir or output_root(configs) / record_id / "consumer_feedback_tables"
    out_dir.mkdir(parents=True, exist_ok=True)
    summaries = summarize_social_cleaned_workbook(cleaned_raw_data)
    return build_social_feedback_workbook(
        brand=brand,
        product=product,
        start_date=start_date,
        end_date=end_date,
        summaries=summaries,
        output_dir=out_dir,
    )


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="统计消费者反馈标签数表")
    parser.add_argument("--record-id", required=True)
    parser.add_argument("--cleaned-raw-data", required=True)
    parser.add_argument("--brand", required=True)
    parser.add_argument("--product", required=True)
    parser.add_argument("--start-date", required=True)
    parser.add_argument("--end-date", required=True)
    parser.add_argument("--output-dir")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    output = generate_consumer_feedback_tables(
        args.record_id,
        Path(args.cleaned_raw_data),
        brand=args.brand,
        product=args.product,
        start_date=args.start_date,
        end_date=args.end_date,
        output_dir=Path(args.output_dir) if args.output_dir else None,
    )
    print(output)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
