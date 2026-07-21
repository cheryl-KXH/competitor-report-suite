from __future__ import annotations

import asyncio
import base64
import html
import re
import sys
import tempfile
import urllib.request
from collections import defaultdict
from datetime import date
from pathlib import Path
from typing import Any

SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

from generate_weekly_report import DataQualityReport, ImageCache, ReportRecord, clean_price_text, format_title_date, quality_record_label


ROOT = Path(__file__).resolve().parents[1]


def load_image_layout(path: Path | None = None) -> dict[str, Any]:
    import json

    layout_path = path or ROOT / "config" / "image_layout.json"
    with layout_path.open("r", encoding="utf-8") as f:
        return json.load(f)


def image_to_data_uri(path: Path) -> str:
    if not path.exists():
        return ""
    mime = "image/png"
    if path.suffix.lower() in {".jpg", ".jpeg"}:
        mime = "image/jpeg"
    elif path.suffix.lower() == ".webp":
        mime = "image/webp"
    data = base64.b64encode(path.read_bytes()).decode("ascii")
    return f"data:{mime};base64,{data}"


def configured_font_path(configs: dict[str, dict[str, Any]], font_key: str) -> Path | None:
    font_files = configs.get("font_files", {})
    font_cfg = configs.get("font_files", {}).get(font_key, {})
    file_name = str(font_cfg.get("fileName") or "").strip()
    if not file_name or Path(file_name).name != file_name:
        return None
    font_dir_value = Path(str(font_files.get("fontDirectory") or "assets/fonts"))
    if font_dir_value.is_absolute():
        return None
    path = ROOT / font_dir_value / file_name
    return path if path.exists() else None


def html_font_family(configs: dict[str, dict[str, Any]], image_layout: dict[str, Any]) -> str:
    fonts = image_layout.get("fonts", {})
    font_files = configs.get("font_files", {})
    default_family = (
        font_files.get("defaultFonts", {}).get("cssFamily")
        or fonts.get("fallbackFamily")
        or "Microsoft YaHei, Arial, sans-serif"
    )
    chinese_key = fonts.get("chineseFontFileKey", "chineseFont")
    latin_key = fonts.get("latinFontFileKey", "latinFont")
    if configured_font_path(configs, chinese_key) or configured_font_path(configs, latin_key):
        return f"ReportFont, {default_family}"
    return str(default_family)


def font_data_uri(path: Path) -> str:
    suffix = path.suffix.lower()
    mime = "font/ttf" if suffix in {".ttf", ".ttc"} else "font/otf" if suffix == ".otf" else "application/octet-stream"
    data = base64.b64encode(path.read_bytes()).decode("ascii")
    return f"data:{mime};base64,{data}"


def font_format(path: Path) -> str:
    suffix = path.suffix.lower()
    if suffix in {".ttf", ".ttc"}:
        return "truetype"
    if suffix == ".otf":
        return "opentype"
    if suffix == ".woff":
        return "woff"
    if suffix == ".woff2":
        return "woff2"
    return "truetype"


def font_face_css(configs: dict[str, dict[str, Any]], image_layout: dict[str, Any]) -> str:
    fonts = image_layout.get("fonts", {})
    chinese_key = fonts.get("chineseFontFileKey", "chineseFont")
    latin_key = fonts.get("latinFontFileKey", "latinFont")
    chinese_path = configured_font_path(configs, chinese_key)
    latin_path = configured_font_path(configs, latin_key)
    css: list[str] = []
    if chinese_path:
        css.append(
            "@font-face { font-family: 'ReportFont'; "
            f"src: url('{font_data_uri(chinese_path)}') format('{font_format(chinese_path)}'); "
            "font-display: block; "
            "unicode-range: U+2E80-2EFF, U+3000-303F, U+3400-4DBF, U+4E00-9FFF, U+F900-FAFF, U+FF00-FFEF; }"
        )
    else:
        print(f"WARNING: HTML 中文字体文件不可访问：{chinese_key}", file=sys.stderr)
    if latin_path:
        css.append(
            "@font-face { font-family: 'ReportFont'; "
            f"src: url('{font_data_uri(latin_path)}') format('{font_format(latin_path)}'); "
            "font-display: block; "
            "unicode-range: U+0000-00FF, U+2000-206F, U+20A0-20CF; }"
        )
    else:
        print(f"WARNING: HTML 英文字体文件不可访问：{latin_key}", file=sys.stderr)
    return "\n".join(css)


def fetch_image_data_uri(url: str, record: ReportRecord, data_quality_report: DataQualityReport, image_cache: ImageCache | None = None) -> str:
    if not url:
        return ""
    if image_cache:
        data_uri = image_cache.data_uri(url, record.record_id)
        if not data_uri:
            data_quality_report.image_download_failures.append(f"{quality_record_label(record)}：产品外观图片下载失败")
        return data_uri
    try:
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=30) as resp:
            content_type = resp.headers.get("Content-Type", "image/png").split(";")[0].strip() or "image/png"
            data = base64.b64encode(resp.read()).decode("ascii")
        return f"data:{content_type};base64,{data}"
    except Exception:
        data_quality_report.image_download_failures.append(f"{quality_record_label(record)}：产品外观图片下载失败")
        return ""


def escape_text(value: Any) -> str:
    return html.escape(str(value or ""), quote=True)


def html_text(value: Any) -> str:
    return escape_text(value).replace("\n", "<br>")


def insert_brand_breakpoints(value: str, space_break: str = "<wbr>") -> str:
    text = str(value or "")
    text = re.sub(r"([\u4e00-\u9fff])([A-Za-z0-9]+)", r"\1<wbr>\2", text)
    text = re.sub(r"([A-Za-z0-9]+)([\u4e00-\u9fff])", r"\1<wbr>\2", text)
    text = text.replace(" ", space_break)
    return text


def render_brand_html(value: str, force_space_break: bool = False) -> str:
    text = insert_brand_breakpoints(value, "<br>" if force_space_break else " <wbr>")
    return html_text(text).replace("&lt;wbr&gt;", "<wbr>").replace("&lt;br&gt;", "<br>")


def render_remark_html(value: str) -> str:
    text = str(value or "")
    text = re.sub(r"([\u4e00-\u9fff])([A-Za-z0-9]+)", r"\1<wbr>\2", text)
    text = re.sub(r"([A-Za-z0-9]+)([\u4e00-\u9fff])", r"\1<wbr>\2", text)
    break_after_chars = set(" 、/-，；：")
    parts: list[str] = []
    index = 0
    while index < len(text):
        if text.startswith("<wbr>", index):
            parts.append("<wbr>")
            index += len("<wbr>")
            continue
        char = text[index]
        if char == "\r":
            index += 1
            continue
        if char == "\n":
            parts.append("<br>")
            index += 1
            continue
        parts.append(html.escape(char, quote=True))
        if char in break_after_chars:
            parts.append("<wbr>")
        index += 1
    return "".join(parts)


PRICE_SOFT_BREAK = "\ue000"


def insert_price_soft_breaks(text: str) -> str:
    pieces = [piece for piece in re.split(r"(?=[(（])|(?<=[)）])", text) if piece]
    return PRICE_SOFT_BREAK.join(pieces)


def split_outside_parentheses(text: str, separator: str = "/") -> list[str]:
    parts: list[str] = []
    current: list[str] = []
    depth = 0
    for char in text:
        if char in "(（":
            depth += 1
        elif char in ")）" and depth > 0:
            depth -= 1
        if char == separator and depth == 0:
            part = "".join(current).strip()
            if part:
                parts.append(part)
            current = []
            continue
        current.append(char)
    tail = "".join(current).strip()
    if tail:
        parts.append(tail)
    return parts


def render_price_html(value: str, wrap_after_slash: bool) -> str:
    text = clean_price_text(value)
    if wrap_after_slash and "/" in text:
        text = "/\n".join(split_outside_parentheses(text))
    else:
        text = insert_price_soft_breaks(text)

    escaped = html.escape(text, quote=True)
    pattern = re.compile(r"(?<=[(（])(\d+(?:\.\d+)?\s*元)(?=[)）])")
    escaped = pattern.sub(r'<span class="strike">\1</span>', escaped)
    return escaped.replace(PRICE_SOFT_BREAK, "<wbr>").replace("\n", "<br>")


def visual_len(text: str) -> float:
    total = 0.0
    for char in str(text or ""):
        if char == "\n":
            continue
        if "\u4e00" <= char <= "\u9fff":
            total += 1.0
        elif char in "（）()[]【】":
            total += 0.62
        elif char in "，。；：！？、,.!?;:/":
            total += 0.42
        elif char.isspace():
            total += 0.34
        else:
            total += 0.56
    return total


def text_width_px(text: str, font_size_px: int, padding_px: int) -> int:
    return int(round(visual_len(text) * font_size_px + padding_px))


def clamp(value: int, minimum: int, maximum: int) -> int:
    return max(minimum, min(maximum, value))


def price_blocks(text: str) -> list[str]:
    value = clean_price_text(text)
    if not value:
        return [""]
    return split_outside_parentheses(value) or [value]


def safe_price_width_segments(text: str) -> list[str]:
    value = clean_price_text(text)
    if "/" in value:
        return price_blocks(value)
    segments: list[str] = []
    for block in price_blocks(value):
        pieces = re.split(r"(?=[(（])|(?<=[)）])", block)
        segments.extend(piece.strip() for piece in pieces if piece.strip())
    return segments or [value]


def longest_ascii_token(text: str) -> str:
    tokens = re.findall(r"[A-Za-z0-9]+", str(text or ""))
    return max(tokens, key=len, default="")


def estimate_wrapped_lines(text: str, width_px: int, font_size_px: int, explicit_breaks: bool = False) -> int:
    if not text:
        return 1
    parts = str(text).split("\n") if explicit_breaks else [str(text)]
    capacity = max(1.0, (width_px - 8) / max(1, font_size_px))
    lines = 0
    for part in parts:
        lines += max(1, int((visual_len(part) + capacity - 0.01) // capacity))
    return max(1, lines)


def estimate_price_lines(text: str, width_px: int, font_size_px: int) -> int:
    if "/" in clean_price_text(text):
        return max(
            len(price_blocks(text)),
            estimate_wrapped_lines(max(safe_price_width_segments(text), key=visual_len, default=""), width_px, font_size_px),
        )

    capacity = max(1.0, (width_px - 8) / max(1, font_size_px))
    lines = 1
    current = ""
    for piece in safe_price_width_segments(text):
        candidate = f"{current}{piece}" if current else piece
        if current and visual_len(candidate) > capacity:
            lines += 1
            current = piece
        else:
            current = candidate
    return max(1, lines)


SUMMARY_COLUMN_KEYS = ("brand", "count", "category", "productName", "launchDate", "price", "remark")


def summary_table_width(columns: dict[str, int]) -> int:
    return sum(int(columns[key]) for key in SUMMARY_COLUMN_KEYS)


def compute_summary_column_widths(records: list[ReportRecord], layout: dict[str, Any]) -> dict[str, int]:
    summary = layout["summary"]
    fixed = dict(summary.get("fixedColumnWidthsPx", {}))
    table_width = int(summary["tableWidthPx"])
    min_max = summary.get("flexColumnMinMaxPx", {})
    padding = summary.get("textWidthPaddingPx", {})
    font_size = int(layout.get("fonts", {}).get("defaultSizePx", 12))

    fixed_width = sum(int(fixed[key]) for key in ("brand", "count", "category", "launchDate"))
    remaining = table_width - fixed_width

    max_name = max((record.product_name for record in records), key=visual_len, default="")
    name_min, name_max = min_max.get("productName", [118, 190])
    name_need = text_width_px(max_name, font_size, int(padding.get("productName", 18)))

    price_min, price_max = min_max.get("price", [72, 104])
    remark_min, remark_max = min_max.get("remark", [72, 140])

    longest_price_segment = max((block for record in records for block in safe_price_width_segments(record.price)), key=visual_len, default="")
    price_need = text_width_px(longest_price_segment, font_size, int(padding.get("price", 18)))
    longest_remark = max((record.remark for record in records), key=visual_len, default="")
    remark_padding = int(padding.get("remark", 16))
    remark_need = text_width_px(longest_remark, font_size, remark_padding)
    longest_remark_token = max((longest_ascii_token(record.remark) for record in records), key=len, default="")
    token_need = text_width_px(longest_remark_token, font_size, remark_padding)

    product_width = max(int(name_min), name_need)
    price_width = clamp(price_need, int(price_min), int(price_max))
    remark_width = max(int(remark_min), min(max(token_need, min(remark_need, int(remark_max))), int(remark_max)))

    hard_width = fixed_width + product_width + price_width + remark_width
    target_width = table_width
    if hard_width > target_width:
        shortage = hard_width - target_width
        target_width += shortage
        extra = 0
    else:
        extra = target_width - hard_width

    product_target = max(product_width, int(name_max))
    give_to_product = min(extra, max(0, product_target - product_width))
    product_width += give_to_product
    extra -= give_to_product

    remark_width += extra

    return {
        "brand": int(fixed["brand"]),
        "count": int(fixed["count"]),
        "category": int(fixed["category"]),
        "productName": int(product_width),
        "launchDate": int(fixed["launchDate"]),
        "price": int(price_width),
        "remark": int(remark_width),
    }


def summary_line_count(record: ReportRecord, columns: dict[str, int], layout: dict[str, Any]) -> int:
    font_size = int(layout.get("fonts", {}).get("defaultSizePx", 12))
    return max(
        1,
        estimate_wrapped_lines(record.category, columns["category"], font_size),
        estimate_price_lines(record.price, columns["price"], font_size),
        estimate_wrapped_lines(record.remark, columns["remark"], font_size),
    )


def title_text(start: date, end: date, configs: dict[str, dict[str, Any]]) -> str:
    template = configs["excel_layout"].get("titleTemplate", "竞品新品周报{start_m}.{start_d}-{end_m}.{end_d}")
    return format_title_date(start, end, template)


def grouped_records(records: list[ReportRecord], brands: list[str]) -> dict[str, list[ReportRecord]]:
    grouped: dict[str, list[ReportRecord]] = defaultdict(list)
    for record in records:
        grouped[record.brand].append(record)
    return {brand: grouped[brand] for brand in brands if grouped.get(brand)}


def build_summary_html(records: list[ReportRecord], brands: list[str], layout: dict[str, Any]) -> str:
    grouped = grouped_records(records, brands)
    columns = compute_summary_column_widths(records, layout)
    colgroup = "".join(f'<col style="width:{columns[key]}px">' for key in ("brand", "count", "category", "productName", "launchDate", "price", "remark"))
    rows: list[str] = [
        '<table class="summary-table">',
        f"<colgroup>{colgroup}</colgroup>",
        '<thead><tr class="summary-header"><th>品牌</th><th>本周新品<br>数量</th><th>品类</th><th>新品名称</th><th>上市时间</th><th>价格</th><th>备注</th></tr></thead>',
        "<tbody>",
    ]
    for brand, brand_records in grouped.items():
        row_count = len(brand_records)
        for index, record in enumerate(brand_records):
            launch = f"{record.launch_date.month}月{record.launch_date.day}日" if record.launch_date else ""
            line_count = min(5, summary_line_count(record, columns, layout))
            row = [f'<tr class="line-count-{line_count}">']
            if index == 0:
                row.append(f'<td class="brand-cell" rowspan="{row_count}">{render_brand_html(brand, force_space_break=True)}</td>')
                row.append(f'<td class="count-cell" rowspan="{row_count}">{row_count}</td>')
            row.extend(
                [
                    f'<td class="category-cell">{html_text(record.category)}</td>',
                    f'<td class="product-name-cell">{html_text(record.product_name)}</td>',
                    f'<td class="launch-date-cell">{html_text(launch)}</td>',
                    f'<td class="price-cell">{render_price_html(record.price, True)}</td>',
                    f'<td class="remark-cell">{render_remark_html(record.remark)}</td>',
                    "</tr>",
                ]
            )
            rows.append("".join(row))
    rows.extend(["</tbody>", "</table>"])
    return "\n".join(rows)


def build_tracked_html(brands: list[str], prefix: str) -> str:
    parts = [escape_text(prefix)]
    for index, brand in enumerate(brands):
        if index:
            parts.append("、")
        parts.append(render_brand_html(brand))
    return "".join(parts)


def render_tracked_html(brands: list[str], prefix: str) -> str:
    return build_tracked_html(brands, prefix)


def build_detail_html(
    records: list[ReportRecord],
    brands: list[str],
    layout: dict[str, Any],
    data_quality_report: DataQualityReport,
    table_width: int | None = None,
    image_cache: ImageCache | None = None,
) -> str:
    grouped = grouped_records(records, brands)
    summary_columns = compute_summary_column_widths(records, layout)
    label_width = summary_columns["brand"]
    detail_table_width = table_width if table_width is not None else int(layout["details"]["tableWidthPx"])
    value_width = detail_table_width - label_width
    colgroup = f'<colgroup><col style="width:{label_width}px"><col style="width:{value_width}px"></colgroup>'
    rows: list[str] = ['<table class="detail-table">', colgroup]
    for brand, brand_records in grouped.items():
        rows.append(
            f'<tr><th class="brand-header" colspan="2">{escape_text(brand)}本周新品数量： {len(brand_records)}个</th></tr>'
        )
        for record in brand_records:
            image_uri = fetch_image_data_uri(record.image_urls[0], record, data_quality_report, image_cache) if record.image_urls else ""
            image_html = (
                f'<img class="product-image" src="{image_uri}" alt="{escape_text(record.product_name)}">'
                if image_uri
                else ""
            )
            rows.extend(
                [
                    f'<tr class="detail-row detail-row-compact"><th class="detail-label strong">新品名称</th><td class="detail-value name-row strong">{html_text(record.product_name)}</td></tr>',
                    f'<tr class="detail-row detail-row-compact"><th class="detail-label">产品系列归属</th><td class="detail-value series-row">{html_text(record.series or "/")}</td></tr>',
                    f'<tr class="detail-row detail-row-text"><th class="detail-label plain">产品卖点介绍</th><td class="detail-value long-text">{html_text(record.selling_point)}</td></tr>',
                    f'<tr class="detail-row detail-row-compact"><th class="detail-label plain">产品价格</th><td class="detail-value">{render_price_html(record.price, False)}</td></tr>',
                    f'<tr class="detail-row detail-row-compact"><th class="detail-label plain">原料构成</th><td class="detail-value">{html_text(record.ingredients)}</td></tr>',
                    f'<tr class="detail-row detail-row-image"><th class="detail-label image-label plain">产品外观</th><td class="detail-value image-cell">{image_html}</td></tr>',
                ]
            )
    rows.append("</table>")
    return "\n".join(rows)


def build_html_document(
    records: list[ReportRecord],
    configs: dict[str, dict[str, Any]],
    image_layout: dict[str, Any],
    start: date,
    end: date,
    brands: list[str],
    tracked_brands: list[str],
    data_quality_report: DataQualityReport,
    image_cache: ImageCache | None = None,
) -> str:
    page = image_layout["page"]
    fonts = image_layout["fonts"]
    colors = image_layout["colors"]
    details = image_layout["details"]
    border_width = int(image_layout.get("border", {}).get("widthPx", 1))
    summary_columns = compute_summary_column_widths(records, image_layout)
    summary_width = summary_table_width(summary_columns)
    configured_summary_width = int(image_layout["summary"]["tableWidthPx"])
    page_margin = max(40, int(page["widthPx"]) - configured_summary_width)
    page_width = max(int(page["widthPx"]), summary_width + page_margin)
    detail_table_width = max(int(details["tableWidthPx"]), summary_width)
    label_width = summary_columns["brand"]
    row_heights = image_layout.get("summary", {}).get("rowHeightByLineCountPx", {})
    row_height_css = "\n".join(
        f".summary-table tr.line-count-{count} > td {{ height: {height}px; }}"
        for count, height in row_heights.items()
    )
    font_css = font_face_css(configs, image_layout)
    font_family = html_font_family(configs, image_layout)
    logo_cfg = image_layout.get("logo", {})
    logo_uri = image_to_data_uri(ROOT / logo_cfg.get("path", ""))
    intro = configs["excel_layout"].get("introText", "")
    tracked_prefix = configs["excel_layout"].get("trackedBrandsPrefix", "*关注品牌包括：")
    tracked_html = render_tracked_html(tracked_brands, tracked_prefix)
    summary_html = build_summary_html(records, brands, image_layout)
    detail_html = build_detail_html(records, brands, image_layout, data_quality_report, detail_table_width, image_cache)
    css = f"""
{font_css}
* {{ box-sizing: border-box; }}
html, body {{
  margin: 0;
  padding: 0;
  background: #{colors['white']};
  -webkit-text-size-adjust: 100%;
  text-size-adjust: 100%;
}}
body {{
  width: 100%;
  min-width: {page_width}px;
  padding: 0;
  color: #{colors['black']};
  font-family: {font_family};
  font-size: {fonts['defaultSizePx']}px;
  line-height: 1.35;
}}
.report-viewport {{ width: {page_width}px; margin: 0 auto; overflow: visible; }}
.report-scale {{ width: {page_width}px; transform-origin: top left; }}
.report {{ width: {page_width}px; margin: 0 auto; padding: {page['paddingTopPx']}px 0 {page['paddingBottomPx']}px; }}
.logo {{ display: block; height: {logo_cfg.get('heightPx', 46)}px; margin: 0 auto {logo_cfg.get('marginBottomPx', 8)}px; }}
.title {{ text-align: center; font-size: {fonts['titleSizePx']}px; font-weight: 700; margin: 0 0 18px; }}
.intro {{ width: {summary_width}px; margin: 0 auto 2px; font-size: {fonts['introSizePx']}px; }}
table {{ border-collapse: collapse; table-layout: fixed; margin-left: auto; margin-right: auto; }}
th, td {{ border: {border_width}px solid #{colors['black']}; text-align: center; vertical-align: middle; padding: 2px 4px; word-break: break-word; }}
.summary-table {{ width: {summary_width}px; margin-bottom: 2px; }}
.summary-table th, .summary-table td {{ padding: 1px 3px; line-height: 1.35; }}
.summary-table th {{ background: #{colors['headerFill']}; font-weight: 700; }}
.summary-header > th {{ height: {image_layout['summary'].get('headerHeightPx', 58)}px; }}
.summary-header > th:nth-child(2) {{ white-space: nowrap; word-break: keep-all; }}
.brand-cell {{ white-space: normal; word-break: normal; overflow-wrap: normal; }}
.product-name-cell {{ white-space: nowrap; word-break: keep-all; overflow-wrap: normal; }}
.category-cell, .launch-date-cell {{ word-break: keep-all; }}
.price-cell {{ word-break: keep-all; overflow-wrap: normal; }}
.remark-cell {{ word-break: normal; overflow-wrap: normal; white-space: normal; }}
.brand-cell {{ background: #{colors['brandFill']}; font-weight: 700; }}
.count-cell {{ font-weight: 700; font-size: 14px; }}
.tracked {{ width: {summary_width}px; margin: 0 auto {image_layout.get('trackedBrands', {}).get('marginBottomPx', 12)}px; font-size: {fonts['smallSizePx']}px; text-align: justify; text-align-last: left; white-space: normal; word-break: normal; overflow-wrap: normal; line-height: 1.35; }}
.detail-table {{ width: {detail_table_width}px; margin-top: 0; margin-bottom: 0; }}
.brand-header {{ height: {details.get('brandHeaderHeightPx', 20)}px; background: #{colors['detailHeaderFill']}; font-weight: 700; }}
.detail-label {{ width: {label_width}px; background: #{colors['detailLabelFill']}; font-weight: 400; white-space: nowrap; word-break: keep-all; }}
.detail-label.plain {{ background: #{colors['white']}; }}
.detail-label.strong {{ font-weight: 700; }}
.detail-table th, .detail-table td {{ line-height: 1.35; padding: 1px 3px; }}
.detail-row-compact > th, .detail-row-compact > td {{ height: 18px; }}
.detail-row-text > th, .detail-row-text > td {{ min-height: 38px; }}
.detail-value {{ padding: {details['cellPaddingPx']}px 4px; }}
.name-row, .series-row {{ background: #{colors['detailLabelFill']}; }}
.strong {{ font-weight: 700; }}
.long-text {{ text-align: {image_layout.get('textAlignment', {}).get('sellingPoint', 'left')}; }}
.image-label {{ height: {details['imageHeightPx'] + 16}px; }}
.image-cell {{ height: {details['imageHeightPx'] + 16}px; padding: 6px; }}
.product-image {{ display: block; max-height: {details['imageHeightPx']}px; max-width: {details['imageMaxWidthPx']}px; margin: 0 auto; object-fit: contain; }}
.strike {{ text-decoration: line-through; }}
@media (min-width: 768px) {{
  html, body {{ background: #f2f2f0; }}
  .report-viewport {{ margin: 24px auto; background: #{colors['white']}; box-shadow: 0 2px 18px rgba(0, 0, 0, .08); }}
}}
@media (max-width: 767px) {{
  html, body {{ width: 100%; min-width: 0; overflow-x: hidden; background: #{colors['white']}; }}
  body {{ min-height: 100%; -webkit-overflow-scrolling: touch; }}
  .report-viewport {{ width: 100vw; margin: 0; overflow: hidden; touch-action: pan-y pinch-zoom; }}
  .report-scale {{ margin: 0; will-change: transform; }}
  .report {{ margin: 0; }}
  th, td {{ border-width: 1.5px; }}
}}
{row_height_css}
"""
    mobile_zoom_script = f"""
<script>
(function() {{
  var reportWidth = {page_width};
  var viewport = null;
  var scaleWrap = null;
  var report = null;
  var baseScale = 1;
  var userScale = 1;
  var pinchStartDistance = 0;
  var pinchStartScale = 1;
  var currentScale = 1;
  var panX = 0;
  var dragStartX = 0;
  var dragStartY = 0;
  var dragStartPanX = 0;
  var dragDirection = '';

  function isMobilePreview() {{
    return window.matchMedia('(max-width: 767px)').matches;
  }}

  function clamp(value, min, max) {{
    return Math.min(max, Math.max(min, value));
  }}

  function touchDistance(touches) {{
    var dx = touches[0].clientX - touches[1].clientX;
    var dy = touches[0].clientY - touches[1].clientY;
    return Math.sqrt(dx * dx + dy * dy);
  }}

  function applyScale() {{
    if (!viewport || !scaleWrap || !report) return;
    if (!isMobilePreview()) {{
      baseScale = 1;
      userScale = 1;
      currentScale = 1;
      panX = 0;
      viewport.style.width = reportWidth + 'px';
      viewport.style.height = '';
      scaleWrap.style.transform = '';
      return;
    }}
    var viewportWidth = Math.max(1, window.innerWidth || document.documentElement.clientWidth || reportWidth);
    baseScale = Math.min(1, viewportWidth / reportWidth);
    var scale = baseScale * userScale;
    currentScale = scale;
    var maxPan = Math.max(0, reportWidth - viewportWidth / scale);
    panX = clamp(panX, -maxPan, 0);
    scaleWrap.style.transform = 'scale(' + scale + ') translateX(' + panX + 'px)';
    viewport.style.width = '100vw';
    viewport.style.height = (report.offsetHeight * scale) + 'px';
  }}

  function setup() {{
    viewport = document.querySelector('.report-viewport');
    scaleWrap = document.querySelector('.report-scale');
    report = document.querySelector('.report');
    applyScale();
    window.addEventListener('resize', applyScale, {{ passive: true }});
    window.addEventListener('orientationchange', function() {{ setTimeout(applyScale, 200); }}, {{ passive: true }});
    if (!viewport) return;
    viewport.addEventListener('touchstart', function(event) {{
      if (event.touches.length === 2) {{
        pinchStartDistance = touchDistance(event.touches);
        pinchStartScale = userScale;
        dragDirection = '';
        event.preventDefault();
      }} else if (event.touches.length === 1 && userScale > 1) {{
        dragStartX = event.touches[0].clientX;
        dragStartY = event.touches[0].clientY;
        dragStartPanX = panX;
        dragDirection = '';
      }}
    }}, {{ passive: false }});
    viewport.addEventListener('touchmove', function(event) {{
      if (event.touches.length === 2 && pinchStartDistance) {{
        userScale = clamp(pinchStartScale * (touchDistance(event.touches) / pinchStartDistance), 1, 4);
        applyScale();
        event.preventDefault();
        return;
      }}
      if (event.touches.length !== 1 || userScale <= 1) return;
      var dx = event.touches[0].clientX - dragStartX;
      var dy = event.touches[0].clientY - dragStartY;
      if (!dragDirection && Math.max(Math.abs(dx), Math.abs(dy)) > 6) {{
        dragDirection = Math.abs(dx) > Math.abs(dy) ? 'horizontal' : 'vertical';
      }}
      if (dragDirection !== 'horizontal') return;
      panX = dragStartPanX + dx / currentScale;
      applyScale();
      event.preventDefault();
    }}, {{ passive: false }});
    viewport.addEventListener('touchend', function(event) {{
      if (event.touches.length < 2) pinchStartDistance = 0;
      if (event.touches.length === 0) dragDirection = '';
    }}, {{ passive: true }});
  }}

  if (document.readyState === 'loading') {{
    document.addEventListener('DOMContentLoaded', setup);
  }} else {{
    setup();
  }}
}})();
</script>
"""
    logo_html = f'<img class="logo" src="{logo_uri}" alt="logo">' if logo_uri else ""
    return f"""<!doctype html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, user-scalable=yes">
<title>{escape_text(title_text(start, end, configs))}</title>
<style>{css}</style>
</head>
<body>
<div class="report-viewport">
<div class="report-scale">
<main class="report">
{logo_html}
<h1 class="title">{escape_text(title_text(start, end, configs))}</h1>
<div class="intro">{escape_text(intro)}</div>
{summary_html}
<div class="tracked">{tracked_html}</div>
{detail_html}
</main>
</div>
</div>
{mobile_zoom_script}
</body>
</html>
"""


async def screenshot_html_with_playwright(html_path: Path, png_path: Path, image_layout: dict[str, Any]) -> None:
    from PIL import Image as PILImage
    from playwright.async_api import async_playwright

    render_cfg = image_layout.get("render", {})
    width = int(image_layout.get("page", {}).get("widthPx", 640))
    scale = int(render_cfg.get("deviceScaleFactor", 2))
    max_chunk_height = int(render_cfg.get("maxScreenshotChunkHeightPx", 4000))
    executable = render_cfg.get("chromeExecutable")
    async with async_playwright() as p:
        browser = await p.chromium.launch(
            executable_path=executable if executable and Path(executable).exists() else None,
            headless=True,
        )
        page = await browser.new_page(viewport={"width": width, "height": 900}, device_scale_factor=scale)
        await page.goto(html_path.resolve().as_uri(), wait_until="load")
        await page.evaluate("document.fonts && document.fonts.ready")
        await page.wait_for_function(
            "() => Array.from(document.images).every(img => img.complete)",
            timeout=30000,
        )
        report_box = await page.locator(".report").evaluate(
            """element => {
                const rect = element.getBoundingClientRect();
                return {
                    x: rect.x,
                    y: rect.y,
                    width: Math.ceil(Math.max(element.scrollWidth, rect.width)),
                    height: Math.ceil(Math.max(element.scrollHeight, rect.height))
                };
            }"""
        )
        await page.set_viewport_size({"width": int(report_box["width"]), "height": int(report_box["height"])})
        report_height = int(report_box["height"])
        if report_height <= max_chunk_height:
            try:
                await page.locator(".report").screenshot(path=str(png_path))
            except Exception:
                await page.screenshot(
                    path=str(png_path),
                    clip={
                        "x": float(report_box["x"]),
                        "y": float(report_box["y"]),
                        "width": float(report_box["width"]),
                        "height": float(report_box["height"]),
                    },
                )
        else:
            chunks: list[PILImage.Image] = []
            y = 0
            with tempfile.TemporaryDirectory() as temp_dir:
                temp_root = Path(temp_dir)
                while y < report_height:
                    chunk_height = min(max_chunk_height, report_height - y)
                    await page.set_viewport_size({"width": int(report_box["width"]), "height": chunk_height})
                    await page.evaluate("(scrollY) => window.scrollTo(0, scrollY)", y)
                    chunk_path = temp_root / f"chunk_{len(chunks)}.png"
                    await page.screenshot(
                        path=str(chunk_path),
                        clip={
                            "x": 0,
                            "y": 0,
                            "width": float(report_box["width"]),
                            "height": float(chunk_height),
                        },
                    )
                    chunks.append(PILImage.open(chunk_path).copy())
                    y += chunk_height
            if chunks:
                stitched = PILImage.new("RGB", (chunks[0].width, sum(chunk.height for chunk in chunks)), "white")
                top = 0
                for chunk in chunks:
                    stitched.paste(chunk.convert("RGB"), (0, top))
                    top += chunk.height
                stitched.save(png_path)
        await browser.close()


def render_report_outputs(
    records: list[ReportRecord],
    configs: dict[str, dict[str, Any]],
    image_layout: dict[str, Any],
    start: date,
    end: date,
    brands: list[str],
    tracked_brands: list[str],
    html_path: Path | None,
    png_path: Path | None,
    data_quality_report: DataQualityReport,
    image_cache: ImageCache | None = None,
) -> None:
    html_content = build_html_document(records, configs, image_layout, start, end, brands, tracked_brands, data_quality_report, image_cache)
    if html_path:
        html_path.parent.mkdir(parents=True, exist_ok=True)
        html_path.write_text(html_content, encoding="utf-8")
        screenshot_source = html_path
    else:
        temp = tempfile.NamedTemporaryFile("w", suffix=".html", encoding="utf-8", delete=False)
        try:
            temp.write(html_content)
            temp.close()
            screenshot_source = Path(temp.name)
        finally:
            pass

    try:
        if png_path:
            png_path.parent.mkdir(parents=True, exist_ok=True)
            asyncio.run(screenshot_html_with_playwright(screenshot_source, png_path, image_layout))
    finally:
        if not html_path:
            screenshot_source.unlink(missing_ok=True)
