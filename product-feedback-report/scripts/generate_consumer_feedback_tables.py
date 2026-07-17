#!/usr/bin/env python3
from __future__ import annotations

import argparse
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from service.config import load_configs, output_root
from scripts.data_processing import generate_social_summaries


def generate_consumer_feedback_tables(record_id: str, input_dir: Path, output_dir: Path | None = None) -> dict[str, Path]:
    configs = load_configs()
    out_dir = output_dir or output_root(configs) / record_id / "consumer_feedback_tables"
    out_dir.mkdir(parents=True, exist_ok=True)
    social_rules = configs.get("platform_rules", {}).get("socialPlatforms", {})
    return generate_social_summaries(input_dir, out_dir, social_rules)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="统计消费者反馈标签数表")
    parser.add_argument("--record-id", required=True)
    parser.add_argument("--input-dir", required=True, help="本地所有数据文件夹")
    parser.add_argument("--output-dir")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    outputs = generate_consumer_feedback_tables(
        args.record_id,
        Path(args.input_dir),
        Path(args.output_dir) if args.output_dir else None,
    )
    for key, path in outputs.items():
        print(f"{key}: {path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
