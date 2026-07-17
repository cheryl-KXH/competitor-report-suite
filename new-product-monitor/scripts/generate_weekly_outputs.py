#!/usr/bin/env python3
from __future__ import annotations

import argparse
import re
import shutil
import sys
import time
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Callable

import generate_weekly_report as excel_report
from render_report_image import load_image_layout, render_report_outputs


ROOT = Path(__file__).resolve().parents[1]
VALID_OUTPUT_MODES = {"all", "excel", "html", "image"}


def log_duration(label: str, started_at: float) -> None:
    print(f"耗时：{label} {time.perf_counter() - started_at:.1f}s", file=sys.stderr)


@dataclass
class WeeklyOutputResult:
    report_dir: Path
    output_paths: list[Path]
    record_count: int
    missing_field_count: int
    image_issue_count: int
    data_quality_warnings: list[str]


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="从钉钉 AI 表生成竞品新品周报 Excel、HTML 和 PNG 长图")
    parser.add_argument("--startDate", help="开始日期，格式 YYYYMMDD，例如 20260502")
    parser.add_argument("--endDate", help="结束日期，格式 YYYYMMDD，例如 20260508")
    parser.add_argument("--brands", help="品牌范围，使用英文逗号分隔，例如：霸王茶姬,古茗")
    parser.add_argument("--business", help="业务范围：喜茶、野萃山、茶坊；--brands 优先于本参数")
    parser.add_argument(
        "--output-mode",
        choices=sorted(VALID_OUTPUT_MODES),
        default="all",
        help="输出模式：all 默认输出 Excel+HTML+PNG；excel 只输出 Excel；html 只输出 HTML；image 只输出 PNG",
    )
    parser.add_argument("--output-dir", help="输出目录，默认读取 config/image_layout.json 的 outputDirectory")
    return parser.parse_args()


def report_stem(start, end) -> str:
    return f"竞品新品周报{start.isoformat()}_{end.isoformat()}"


def format_schedule_report_stem(business: str, year: int | str, week: str, end: date) -> str:
    year_text = str(year).strip()
    yy = year_text[-2:] if len(year_text) >= 2 else f"{end.year % 100:02d}"
    week_text = str(week or "").strip().upper()
    match = re.search(r"W\s*(\d{1,2})", week_text)
    if not match:
        raise ValueError(f"周次格式无法识别：{week}")
    normalized_week = f"W{int(match.group(1)):02d}"
    return f"{business}{yy}{normalized_week}_{end:%Y%m%d}"


def resolve_output_root(arg_output_dir: str | None, image_layout: dict) -> Path:
    output_dir = Path(arg_output_dir or image_layout.get("outputDirectory") or image_layout.get("outputDirectory", "outputs"))
    if not output_dir.is_absolute():
        output_dir = ROOT / output_dir
    return output_dir


def resolve_report_dir(output_root: Path, stem: str) -> Path:
    output_root.mkdir(parents=True, exist_ok=True)
    base = output_root / stem
    if not base.exists():
        return base
    escaped = re.escape(stem)
    pattern = re.compile(rf"^{escaped}_(\d+)$")
    max_suffix = 1
    for existing in output_root.iterdir():
        if not existing.is_dir():
            continue
        if existing.name == stem:
            max_suffix = max(max_suffix, 1)
            continue
        match = pattern.match(existing.name)
        if match:
            max_suffix = max(max_suffix, int(match.group(1)))
    return output_root / f"{stem}_{max_suffix + 1}"


def resolve_overwriting_report_dir(output_root: Path, stem: str) -> Path:
    output_root.mkdir(parents=True, exist_ok=True)
    report_dir = output_root / stem
    if report_dir.exists():
        if not report_dir.is_dir():
            raise SystemExit(f"输出路径已存在但不是文件夹：{report_dir}")
        shutil.rmtree(report_dir)
    return report_dir


def dedupe_quality_warnings(data_quality_report: excel_report.DataQualityReport) -> None:
    data_quality_report.missing_fields = list(dict.fromkeys(data_quality_report.missing_fields))
    data_quality_report.missing_images = list(dict.fromkeys(data_quality_report.missing_images))
    data_quality_report.image_download_failures = list(dict.fromkeys(data_quality_report.image_download_failures))


def generate_weekly_outputs(
    *,
    start: date,
    end: date,
    business: str | None,
    brands: list[str] | None = None,
    output_mode: str = "all",
    output_dir: str | Path | None = None,
    stem: str | None = None,
    overwrite: bool = False,
    configs: dict | None = None,
    image_layout: dict | None = None,
    progress_callback: Callable[[str], None] | None = None,
) -> WeeklyOutputResult:
    if output_mode not in VALID_OUTPUT_MODES:
        raise ValueError(f"输出模式不支持：{output_mode}")

    total_started_at = time.perf_counter()
    configs = configs or excel_report.load_configs()
    image_layout = image_layout or load_image_layout()
    brands = brands or []

    stage_started_at = time.perf_counter()
    if progress_callback:
        progress_callback("1/5 正在读取并整理钉钉表数据")
    table_fields = excel_report.fetch_table_fields(configs)
    field_ids = excel_report.resolve_field_ids(configs["field_mapping"], table_fields)
    default_brand_order = excel_report.default_brand_option_order(configs, field_ids) if not brands else []
    if not brands and not default_brand_order:
        print("WARNING: 未从钉钉品牌字段读取到标签列表顺序，默认按本次记录首次出现顺序输出品牌。")
    raw_records = excel_report.query_records(configs, field_ids, start, end)
    records = excel_report.normalize_records(raw_records, configs, field_ids, brands, business, default_brand_order)
    if not brands:
        brands = excel_report.effective_output_brands(records, default_brand_order)
    tracked_brands = excel_report.fetch_tracked_brands(configs, business) if business else brands
    data_quality_report = excel_report.collect_data_quality_report(records, configs["report_rules"])
    log_duration("读取并整理钉钉表数据", stage_started_at)

    output_root = resolve_output_root(str(output_dir) if output_dir else None, image_layout)
    report_stem_value = stem or report_stem(start, end)
    report_dir = (
        resolve_overwriting_report_dir(output_root, report_stem_value)
        if overwrite
        else resolve_report_dir(output_root, report_stem_value)
    )
    report_dir.mkdir(parents=True, exist_ok=True)
    file_stem = report_dir.name

    output_paths: list[Path] = []
    image_cache = excel_report.ImageCache()
    try:
        stage_started_at = time.perf_counter()
        if progress_callback:
            progress_callback("2/5 正在预下载产品图片")
        configured_workers = int(image_layout.get("render", {}).get("imageDownloadWorkers", 1))
        image_cache.prefetch(records, max_workers=configured_workers)
        log_duration("预下载产品图片", stage_started_at)

        if output_mode in {"all", "excel"}:
            stage_started_at = time.perf_counter()
            if progress_callback:
                progress_callback("3/5 正在生成 Excel")
            xlsx_path = report_dir / f"{file_stem}.xlsx"
            excel_report.build_workbook(records, configs, start, end, brands, tracked_brands, xlsx_path, data_quality_report, image_cache)
            excel_report.cleanup_legacy_image_cache(report_dir)
            output_paths.append(xlsx_path)
            log_duration("生成 Excel", stage_started_at)

        html_path = report_dir / f"{file_stem}.html" if output_mode in {"all", "html"} else None
        png_path = report_dir / f"{file_stem}.png" if output_mode in {"all", "image"} else None
        if html_path or png_path:
            stage_started_at = time.perf_counter()
            if progress_callback:
                progress_callback("4/5 正在生成 HTML/PNG")
            render_report_outputs(
                records,
                configs,
                image_layout,
                start,
                end,
                brands,
                tracked_brands,
                html_path,
                png_path,
                data_quality_report,
                image_cache,
            )
            if html_path:
                output_paths.append(html_path)
            if png_path:
                output_paths.append(png_path)
            log_duration("生成 HTML/PNG", stage_started_at)
    finally:
        image_cache.cleanup()

    dedupe_quality_warnings(data_quality_report)
    excel_report.print_data_quality_warnings(data_quality_report)
    data_quality_warnings = excel_report.data_quality_warning_lines(data_quality_report)
    log_duration("总生成流程", total_started_at)
    return WeeklyOutputResult(
        report_dir=report_dir,
        output_paths=output_paths,
        record_count=len(records),
        missing_field_count=data_quality_report.missing_field_count,
        image_issue_count=data_quality_report.image_issue_count,
        data_quality_warnings=data_quality_warnings,
    )


def main() -> int:
    args = parse_args()
    configs = excel_report.load_configs()
    image_layout = load_image_layout()

    warnings = excel_report.validate_config(configs)
    blocking = [w for w in warnings if "缺少" in w or "必须" in w or "需要" in w]
    if blocking:
        print("配置校验失败：")
        for warning in warnings:
            print(f"- {warning}")
        return 1
    for warning in warnings:
        print(f"WARNING: {warning}")

    start, end = excel_report.resolve_date_window(args)
    business = excel_report.parse_business(args.business, configs["report_rules"])
    result = generate_weekly_outputs(
        start=start,
        end=end,
        business=business,
        brands=excel_report.parse_brands(args.brands),
        output_mode=args.output_mode,
        output_dir=args.output_dir,
        configs=configs,
        image_layout=image_layout,
    )
    print(f"已生成目录：{result.report_dir}")
    print(f"记录数：{result.record_count}")
    print(f"数据提醒：缺失字段 {result.missing_field_count} 处，缺图 {result.image_issue_count} 条。")
    for path in result.output_paths:
        print(f"本次生成文件：{path.name}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
