"""参考原 PDF 结构生成自包含 HTML 跟踪报告。"""

from __future__ import annotations

import base64
import html
import mimetypes
import re
import unicodedata
import urllib.request
from collections import Counter
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any

from openpyxl import load_workbook


ROOT = Path(__file__).resolve().parents[2]
SUITE_ROOT = ROOT.parent.resolve()
GENERIC_POSITIVE_TAGS = {"好喝喜欢推荐", "喜欢推荐", "好评"}
GENERIC_NEGATIVE_TAGS = {"难喝不喜欢不推荐", "不喜欢不推荐", "差评"}


@dataclass(frozen=True)
class ProductInfo:
    product_name: str
    series: str = ""
    selling_point: str = ""
    price: str = ""
    ingredients: str = ""
    image_urls: tuple[str, ...] = ()


@dataclass(frozen=True)
class PlatformSale:
    product: str
    daily_store_avg: float
    total_sales: float
    stores: int


@dataclass(frozen=True)
class CombinedSale:
    rank: int
    key: str
    product: str
    meituan_daily: float
    eleme_daily: float
    combined_daily: float
    total_sales: float
    share: float
    tracked: bool = False


@dataclass(frozen=True)
class DeliveryReport:
    all_rows: tuple[CombinedSale, ...]
    display_rows: tuple[CombinedSale, ...]
    tracked_rows: dict[str, CombinedSale]
    meituan_total_daily: float
    eleme_total_daily: float
    meituan_stores: int
    eleme_stores: int


@dataclass(frozen=True)
class JdSale:
    rank: int
    product: str
    total_sales: float
    share: float


@dataclass(frozen=True)
class SocialSection:
    label: str
    positive_tags: tuple[tuple[str, int], ...]
    negative_tags: tuple[tuple[str, int], ...]
    positive_users: int
    negative_users: int


@dataclass(frozen=True)
class SocialReport:
    title: str
    period: str
    sections: tuple[SocialSection, ...]
    positive_users: int
    negative_users: int
    positive_top: tuple[tuple[str, int], ...]
    negative_top: tuple[tuple[str, int], ...]

    @property
    def total_users(self) -> int:
        return self.positive_users + self.negative_users

    @property
    def positive_rate(self) -> float:
        return self.positive_users / self.total_users if self.total_users else 0.0


@dataclass(frozen=True)
class ReportBuildResult:
    path: Path
    warnings: tuple[str, ...]


def _text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def _number(value: Any) -> float:
    if value is None or value == "":
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    match = re.search(r"-?\d+(?:\.\d+)?", str(value).replace(",", ""))
    return float(match.group(0)) if match else 0.0


def _integer(value: Any) -> int:
    return int(round(_number(value)))


def normalize_product_key(value: Any) -> str:
    text = unicodedata.normalize("NFKC", _text(value)).casefold()
    return re.sub(r"[^0-9a-z\u3400-\u9fff]+", "", text)


def _header_map(values: tuple[Any, ...] | list[Any]) -> dict[str, int]:
    return {_text(value).replace(" ", ""): index for index, value in enumerate(values)}


def read_platform_sales(path: Path, platform: str) -> tuple[dict[str, PlatformSale], int]:
    workbook = load_workbook(path, read_only=True, data_only=True)
    worksheet = workbook.active
    rows = worksheet.iter_rows(values_only=True)
    try:
        header = next(rows)
    except StopIteration as exc:
        raise RuntimeError(f"{platform}外卖数表为空：{path.name}") from exc
    columns = _header_map(header)
    required = {"商品名称", "日店均销量", "总销量", "在售门店数"}
    missing = required - set(columns)
    if missing:
        raise RuntimeError(f"{platform}外卖数表缺少字段：{'、'.join(sorted(missing))}")
    sales: dict[str, PlatformSale] = {}
    max_stores = 0
    for row in rows:
        product = _text(row[columns["商品名称"]])
        if not product:
            continue
        key = normalize_product_key(product)
        if key in sales:
            raise RuntimeError(f"{platform}外卖数表存在规范化后重名商品：{sales[key].product}、{product}")
        stores = _integer(row[columns["在售门店数"]])
        sales[key] = PlatformSale(
            product=product,
            daily_store_avg=_number(row[columns["日店均销量"]]),
            total_sales=_number(row[columns["总销量"]]),
            stores=stores,
        )
        max_stores = max(max_stores, stores)
    if not sales:
        raise RuntimeError(f"{platform}外卖数表没有有效商品数据：{path.name}")
    return sales, max_stores


def _tracked_matches(products: list[str], names: dict[str, str]) -> dict[str, str]:
    matched: dict[str, str] = {}
    for product in products:
        target = normalize_product_key(product)
        exact = [key for key in names if key == target]
        candidates = exact or [key for key in names if target and target in key]
        if len(candidates) > 1:
            labels = "、".join(names[key] for key in candidates)
            raise RuntimeError(f"关注新品“{product}”匹配到多个外卖商品：{labels}")
        if candidates:
            matched[product] = candidates[0]
    return matched


def combine_delivery_sales(
    meituan_path: Path, eleme_path: Path, tracked_products: list[str]
) -> DeliveryReport:
    meituan, meituan_stores = read_platform_sales(meituan_path, "美团")
    eleme, eleme_stores = read_platform_sales(eleme_path, "饿了么")
    keys = set(meituan) | set(eleme)
    names = {
        key: (meituan.get(key) or eleme[key]).product
        for key in keys
    }
    tracked_matches = _tracked_matches(tracked_products, names)
    tracked_keys = set(tracked_matches.values())
    total_sales = sum(
        (meituan.get(key).total_sales if key in meituan else 0.0)
        + (eleme.get(key).total_sales if key in eleme else 0.0)
        for key in keys
    )
    sortable: list[tuple[str, str, float, float, float, float]] = []
    for key in keys:
        mt = meituan.get(key)
        elm = eleme.get(key)
        mt_daily = mt.daily_store_avg if mt else 0.0
        elm_daily = elm.daily_store_avg if elm else 0.0
        combined_sales = (mt.total_sales if mt else 0.0) + (elm.total_sales if elm else 0.0)
        sortable.append((key, names[key], mt_daily, elm_daily, mt_daily + elm_daily, combined_sales))
    sortable.sort(key=lambda item: (-item[5], item[1]))
    ranked: list[CombinedSale] = []
    previous_sales: float | None = None
    previous_rank = 0
    for index, (key, product, mt_daily, elm_daily, combined_daily, combined_sales) in enumerate(sortable, 1):
        rank = previous_rank if previous_sales is not None and combined_sales == previous_sales else index
        previous_sales = combined_sales
        previous_rank = rank
        ranked.append(
            CombinedSale(
                rank=rank,
                key=key,
                product=product,
                meituan_daily=mt_daily,
                eleme_daily=elm_daily,
                combined_daily=combined_daily,
                total_sales=combined_sales,
                share=combined_sales / total_sales if total_sales else 0.0,
                tracked=key in tracked_keys,
            )
        )
    display_rows = tuple(row for row in ranked if row.rank <= 20 or row.tracked)
    by_key = {row.key: row for row in ranked}
    tracked_rows = {
        product: by_key[key]
        for product, key in tracked_matches.items()
    }
    return DeliveryReport(
        all_rows=tuple(ranked),
        display_rows=display_rows,
        tracked_rows=tracked_rows,
        meituan_total_daily=sum(item.daily_store_avg for item in meituan.values()),
        eleme_total_daily=sum(item.daily_store_avg for item in eleme.values()),
        meituan_stores=meituan_stores,
        eleme_stores=eleme_stores,
    )


def read_jd_sales(path: Path | None) -> tuple[tuple[JdSale, ...], float]:
    if not path or not path.exists():
        return (), 0.0
    workbook = load_workbook(path, read_only=True, data_only=True)
    worksheet = workbook.active
    rows = worksheet.iter_rows(values_only=True)
    try:
        header = next(rows)
    except StopIteration:
        return (), 0.0
    columns = _header_map(header)
    required = {"排名", "商品名称", "总销量", "总销量占比"}
    if required - set(columns):
        raise RuntimeError(f"京东外卖数表缺少字段：{'、'.join(sorted(required - set(columns)))}")
    all_rows: list[JdSale] = []
    for row in rows:
        product = _text(row[columns["商品名称"]])
        if not product:
            continue
        all_rows.append(
            JdSale(
                rank=_integer(row[columns["排名"]]),
                product=product,
                total_sales=_number(row[columns["总销量"]]),
                share=_number(row[columns["总销量占比"]]),
            )
        )
    return tuple(row for row in all_rows if row.rank <= 20), sum(row.total_sales for row in all_rows)


def _counter_top(counter: Counter[str], generic: set[str], limit: int = 3) -> tuple[tuple[str, int], ...]:
    filtered = [
        (label, count)
        for label, count in counter.items()
        if normalize_product_key(label) not in generic and label and count > 0
    ]
    filtered.sort(key=lambda item: (-item[1], item[0]))
    return tuple(filtered[:limit])


def read_social_report(path: Path) -> SocialReport:
    workbook = load_workbook(path, read_only=True, data_only=True)
    worksheet = workbook.active
    values = list(worksheet.iter_rows(values_only=True))
    if not values:
        raise RuntimeError(f"社媒评论统计表为空：{path.name}")
    title = _text(values[0][0])
    period_match = re.search(r"(\d{1,2}\.\d{1,2}\s*-\s*\d{1,2}\.\d{1,2})", title)
    period = period_match.group(1).replace(" ", "") if period_match else ""
    sections: list[SocialSection] = []
    positive_counter: Counter[str] = Counter()
    negative_counter: Counter[str] = Counter()
    row_index = 1
    while row_index < len(values):
        row = list(values[row_index]) + [None] * 6
        positive_header = _text(row[0])
        negative_header = _text(row[2])
        if _text(row[1]) != "评论数" or _text(row[3]) != "评论数" or not positive_header.endswith("好评"):
            row_index += 1
            continue
        label = positive_header[:-2]
        positive_tags: list[tuple[str, int]] = []
        negative_tags: list[tuple[str, int]] = []
        positive_users = 0
        negative_users = 0
        row_index += 1
        while row_index < len(values):
            detail = list(values[row_index]) + [None] * 6
            if _text(detail[0]) == "好评用户数" and _text(detail[2]) == "差评用户数":
                positive_users = _integer(detail[1])
                negative_users = _integer(detail[3])
                row_index += 1
                break
            positive_label = _text(detail[0])
            negative_label = _text(detail[2])
            positive_count = _integer(detail[1])
            negative_count = _integer(detail[3])
            if positive_label and positive_label != "/" and positive_count > 0:
                positive_tags.append((positive_label, positive_count))
                positive_counter[positive_label] += positive_count
            if negative_label and negative_label != "/" and negative_count > 0:
                negative_tags.append((negative_label, negative_count))
                negative_counter[negative_label] += negative_count
            row_index += 1
        sections.append(
            SocialSection(
                label=label or negative_header[:-2],
                positive_tags=tuple(positive_tags),
                negative_tags=tuple(negative_tags),
                positive_users=positive_users,
                negative_users=negative_users,
            )
        )
    generic_positive = {normalize_product_key(value) for value in GENERIC_POSITIVE_TAGS}
    generic_negative = {normalize_product_key(value) for value in GENERIC_NEGATIVE_TAGS}
    return SocialReport(
        title=title,
        period=period,
        sections=tuple(sections),
        positive_users=sum(section.positive_users for section in sections),
        negative_users=sum(section.negative_users for section in sections),
        positive_top=_counter_top(positive_counter, generic_positive),
        negative_top=_counter_top(negative_counter, generic_negative),
    )


def _safe_config_path(value: str) -> Path | None:
    if not value:
        return None
    candidate = (ROOT / value).resolve()
    try:
        candidate.relative_to(SUITE_ROOT)
    except ValueError:
        return None
    return candidate if candidate.exists() and candidate.is_file() else None


def _file_data_uri(path: Path) -> str:
    mime = mimetypes.guess_type(path.name)[0] or "application/octet-stream"
    return f"data:{mime};base64,{base64.b64encode(path.read_bytes()).decode('ascii')}"


def _remote_image_data_uri(url: str) -> str:
    if url.startswith("data:image/"):
        return url
    request = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(request, timeout=30) as response:
        content = response.read(15 * 1024 * 1024 + 1)
        if len(content) > 15 * 1024 * 1024:
            raise RuntimeError("产品图片超过 15MB")
        content_type = response.headers.get("Content-Type", "image/png").split(";")[0].strip()
        if not content_type.startswith("image/"):
            raise RuntimeError(f"产品图片类型无效：{content_type}")
    return f"data:{content_type};base64,{base64.b64encode(content).decode('ascii')}"


def _font_css(configs: dict[str, dict[str, Any]], warnings: list[str]) -> tuple[str, str]:
    layout = configs.get("html_layout", {})
    font_config = configs.get("font_files", {})
    fonts = layout.get("fonts", {})
    default_family = str(
        font_config.get("defaultFonts", {}).get("cssFamily")
        or fonts.get("fallbackFamily")
        or "Microsoft YaHei, Arial, sans-serif"
    )
    font_dir = str(font_config.get("fontDirectory") or "")
    css: list[str] = []
    for key, unicode_range in (
        (fonts.get("chineseFontFileKey", "chineseFont"), "U+2E80-2EFF, U+3000-303F, U+3400-4DBF, U+4E00-9FFF, U+F900-FAFF, U+FF00-FFEF"),
        (fonts.get("latinFontFileKey", "latinFont"), "U+0000-00FF, U+2000-206F, U+20A0-20CF"),
    ):
        filename = _text(font_config.get(key, {}).get("fileName"))
        path = _safe_config_path(str(Path(font_dir) / filename)) if filename else None
        if not path:
            warnings.append(f"HTML 字体文件不可访问：{filename or key}，已使用系统字体")
            continue
        font_format = "opentype" if path.suffix.lower() == ".otf" else "truetype"
        css.append(
            "@font-face{font-family:'ReportFont';"
            f"src:url('{_file_data_uri(path)}') format('{font_format}');"
            f"font-display:block;unicode-range:{unicode_range};}}"
        )
    family = f"'ReportFont', {default_family}" if css else default_family
    return "\n".join(css), family


def _escape(value: Any) -> str:
    return html.escape(_text(value), quote=True)


def _multiline(value: Any) -> str:
    text = _escape(value)
    text = re.sub(r"\*\*([^*]+)\*\*", r"\1", text)
    return text.replace("\n", "<br>")


def _format_count(value: float) -> str:
    return f"{value:,.0f}"


def _format_daily(value: float) -> str:
    return f"{value:.1f}"


def _format_percent(value: float) -> str:
    return f"{value:.1%}"


def _top_labels(values: tuple[tuple[str, int], ...]) -> str:
    return "、".join(label for label, _ in values) or "暂无明确高频观点"


def _delivery_html(brand: str, delivery: DeliveryReport) -> str:
    rows = []
    for row in delivery.display_rows:
        class_name = ' class="tracked-row"' if row.tracked else ""
        rows.append(
            f"<tr{class_name}><td>{_escape(row.product)}</td>"
            f"<td>{_format_daily(row.meituan_daily)}</td><td>{_format_daily(row.eleme_daily)}</td>"
            f"<td>{_format_daily(row.combined_daily)}</td><td>{_format_percent(row.share)}</td><td>{row.rank}</td></tr>"
        )
    total_daily = delivery.meituan_total_daily + delivery.eleme_total_daily
    rows.append(
        "<tr class=\"total-row\"><td>总计</td>"
        f"<td>{_format_daily(delivery.meituan_total_daily)}</td>"
        f"<td>{_format_daily(delivery.eleme_total_daily)}</td>"
        f"<td>{_format_daily(total_daily)}</td><td>100.0%</td><td>-</td></tr>"
    )
    return f"""
<div class="table-scroll">
<table class="sales-table">
<caption>{_escape(brand)} 美团&amp;饿了么线上外卖日店均销售表</caption>
<thead><tr><th rowspan="2">商品名称</th><th>美团</th><th>饿了么</th><th colspan="3">线上合计</th></tr>
<tr><th>日店均销量</th><th>日店均销量</th><th>日店均销量</th><th>销量占比</th><th>排名</th></tr></thead>
<tbody>{''.join(rows)}</tbody>
</table></div>
<p class="source-note">数据来源：根据{_escape(brand)}美团外卖 {_format_count(delivery.meituan_stores)} 家店铺和饿了么外卖 {_format_count(delivery.eleme_stores)} 家店铺月销数据计算</p>
"""


def _jd_html(brand: str, rows: tuple[JdSale, ...], total_sales: float) -> str:
    if not rows:
        return f'<p class="missing-note">{_escape(brand)}京东外卖销售数据暂无法获取。</p>'
    body = "".join(
        f"<tr><td>{row.rank}</td><td>{_escape(row.product)}</td><td>{_format_count(row.total_sales)}</td><td>{_format_percent(row.share)}</td></tr>"
        for row in rows
    )
    body += f'<tr class="total-row"><td colspan="2">店铺合计</td><td>{_format_count(total_sales)}</td><td>100.0%</td></tr>'
    return f"""
<div class="table-scroll"><table class="jd-table"><caption>{_escape(brand)} 京东外卖销量表现</caption>
<thead><tr><th>排名</th><th>商品名称</th><th>京东外卖销量</th><th>销量占比</th></tr></thead>
<tbody>{body}</tbody></table></div>
"""


def _product_info_html(product: str, info: ProductInfo | None, image_uri: str) -> str:
    if not info:
        return '<p class="missing-note">周报产品信息暂无法获取。</p>'
    image_html = (
        f'<img class="product-image" src="{image_uri}" alt="{_escape(product)} 产品外观">'
        if image_uri
        else '<span class="missing-inline">产品外观暂无法获取</span>'
    )
    values = (
        ("新品名称", info.product_name or product, "compact"),
        ("产品系列归属", info.series, "compact"),
        ("产品卖点介绍", info.selling_point, "long"),
        ("产品价格", info.price, "compact"),
        ("原料构成", info.ingredients, "long"),
    )
    rows = "".join(
        f'<tr class="{kind}"><th>{label}</th><td>{_multiline(value) or "-"}</td></tr>'
        for label, value, kind in values
    )
    return f'<table class="product-info"><tbody>{rows}<tr class="image-row"><th>产品外观</th><td>{image_html}</td></tr></tbody></table>'


def _social_summary_html(social: SocialReport | None) -> str:
    if not social:
        return '<p class="missing-note">消费者反馈统计暂无法获取。</p>'
    period = f"（{_escape(social.period)}）" if social.period else ""
    if not social.total_users:
        return f'<p>上市30日{period}暂未获取到有效第三方评论。</p>'
    return (
        f'<p>上市30日{period}第三方评论共 {_format_count(social.total_users)} 条，好评率为 {_format_percent(social.positive_rate)}：</p>'
        f'<p>好评（{_format_count(social.positive_users)} 条）主要提及关键词：{_escape(_top_labels(social.positive_top))}；<br>'
        f'差评（{_format_count(social.negative_users)} 条）主要提及关键词：{_escape(_top_labels(social.negative_top))}。</p>'
    )


def _social_detail_html(social: SocialReport | None) -> str:
    if not social:
        return '<p class="missing-note">消费者反馈详情暂无法获取。</p>'
    section_rows: list[str] = []
    for section in social.sections:
        section_rows.append(
            f'<tr class="platform-header"><th>{_escape(section.label)}好评</th><th>评论数</th>'
            f'<th>{_escape(section.label)}差评</th><th>评论数</th></tr>'
        )
        detail_count = max(len(section.positive_tags), len(section.negative_tags))
        if not detail_count:
            section_rows.append('<tr><td colspan="4" class="empty-platform">暂无评论</td></tr>')
        for index in range(detail_count):
            positive = section.positive_tags[index] if index < len(section.positive_tags) else ("", 0)
            negative = section.negative_tags[index] if index < len(section.negative_tags) else ("", 0)
            section_rows.append(
                f'<tr><td>{_escape(positive[0])}</td><td>{positive[1] if positive[0] else ""}</td>'
                f'<td>{_escape(negative[0])}</td><td>{negative[1] if negative[0] else ""}</td></tr>'
            )
        if section.positive_users or section.negative_users:
            section_rows.append(
                f'<tr class="platform-total"><th>好评用户数</th><td>{section.positive_users or ""}</td>'
                f'<th>差评用户数</th><td>{section.negative_users or ""}</td></tr>'
            )
    summary = ""
    if social.total_users:
        summary = (
            f'<div class="feedback-kpis"><span>好评率</span><strong>{_format_percent(social.positive_rate)}</strong>'
            f'<span>总计</span><strong>{_format_count(social.total_users)}</strong>'
            f'<span>好评用户数</span><strong>{_format_count(social.positive_users)}</strong>'
            f'<span>差评用户数</span><strong>{_format_count(social.negative_users)}</strong></div>'
        )
    return f"""
<div class="feedback-title">{_escape(social.title)}</div>{summary}
<div class="table-scroll"><table class="feedback-table"><tbody>{''.join(section_rows)}</tbody></table></div>
"""


def build_report_html(
    *,
    title: str,
    brand: str,
    products: list[str],
    report_date: date,
    meituan_path: Path,
    eleme_path: Path,
    jd_path: Path | None,
    social_paths: dict[str, Path],
    product_infos: dict[str, ProductInfo],
    launch_dates: dict[str, date | None],
    output_path: Path,
    configs: dict[str, dict[str, Any]],
) -> ReportBuildResult:
    warnings: list[str] = []
    delivery = combine_delivery_sales(meituan_path, eleme_path, products)
    jd_rows, jd_total = read_jd_sales(jd_path)
    social_reports: dict[str, SocialReport] = {}
    for product, path in social_paths.items():
        try:
            social_reports[product] = read_social_report(path)
        except Exception as exc:
            warnings.append(f"{product}：消费者反馈统计读取失败（{exc}）")

    layout = configs.get("html_layout", {})
    page = layout.get("page", {})
    fonts = layout.get("fonts", {})
    colors = layout.get("colors", {})
    tables = layout.get("tables", {})
    product_image = layout.get("productImage", {})
    font_css, font_family = _font_css(configs, warnings)
    logo_path = _safe_config_path(_text(layout.get("logo", {}).get("path")))
    logo_uri = _file_data_uri(logo_path) if logo_path else ""
    if not logo_uri:
        warnings.append("HTML 报告 Logo 不可访问，已使用纯文字标题")

    product_sections: list[str] = []
    for index, product in enumerate(products, 1):
        info = product_infos.get(product)
        if not info:
            warnings.append(f"{product}：周报产品信息暂无法获取")
        social = social_reports.get(product)
        if not social:
            warnings.append(f"{product}：消费者反馈统计暂无法获取")
        image_uri = ""
        if info and info.image_urls:
            try:
                image_uri = _remote_image_data_uri(info.image_urls[0])
            except Exception as exc:
                warnings.append(f"{product}：产品外观图片下载失败（{exc}）")
        price_suffix = f"（{_escape(info.price)}）" if info and info.price else ""
        product_sections.append(
            f'<section class="product-section"><h2>{index}. {_escape(product)}{price_suffix}</h2>'
            f'<h3>1. 产品信息</h3>{_product_info_html(product, info, image_uri)}'
            f'<h3>2. 新品消费者反馈汇总</h3>{_social_summary_html(social)}'
            f'<h3>3. 消费者反馈详情</h3>{_social_detail_html(social)}</section>'
        )

    launch_values = [value for value in launch_dates.values() if value]
    launch_text = ""
    if launch_values and len(set(launch_values)) == 1:
        launch_text = f"{launch_values[0].month}.{launch_values[0].day} "
    intro_rows: list[str] = []
    for product in products:
        sale = delivery.tracked_rows.get(product)
        if sale:
            intro_rows.append(
                f'<li><strong>{_escape(product)}</strong>：销量排名第 {sale.rank}，线上外卖日店均 '
                f'{_format_daily(sale.combined_daily)} 杯，销量占比 {_format_percent(sale.share)}；</li>'
            )
        else:
            warnings.append(f"{product}：未在美团和饿了么外卖数表中匹配到对应商品")
            intro_rows.append(f'<li><strong>{_escape(product)}</strong>：暂未匹配到外卖销售数据；</li>')

    logo_html = f'<img class="logo" src="{logo_uri}" alt="报告 Logo">' if logo_uri else ""
    title_text = f"{brand}：{'、'.join(products)}"
    document_title = title_text or title
    css = f"""
{font_css}
*{{box-sizing:border-box}}
html,body{{margin:0;padding:0;background:#f2f2f0;color:#{colors.get('black','000000')};font-family:{font_family};font-size:{fonts.get('defaultSizePx',12)}px;line-height:1.55}}
.report{{width:{page.get('widthPx',794)}px;max-width:100%;min-height:1123px;margin:24px auto;background:#{page.get('background','FFFFFF')};padding:{page.get('paddingTopPx',34)}px {page.get('paddingRightPx',44)}px {page.get('paddingBottomPx',52)}px {page.get('paddingLeftPx',44)}px;box-shadow:0 2px 18px rgba(0,0,0,.08)}}
.logo{{display:block;height:{layout.get('logo',{}).get('heightPx',72)}px;max-width:70%;object-fit:contain;margin:0 auto {layout.get('logo',{}).get('marginBottomPx',8)}px}}
h1{{font-size:{fonts.get('titleSizePx',21)}px;text-align:center;margin:0 0 24px;font-weight:700}}
h2{{font-size:{fonts.get('productTitleSizePx',17)}px;margin:30px 0 10px}}
h3{{font-size:{fonts.get('sectionSizePx',16)}px;margin:20px 0 8px}}
p{{margin:6px 0}} .lead{{margin-bottom:10px}} .sales-summary{{padding-left:20px;margin:4px 0 14px}}
.sales-summary li{{margin:2px 0}} .table-scroll{{width:100%;overflow-x:auto}}
table{{border-collapse:collapse;table-layout:fixed;width:100%;margin:8px auto 12px}}
caption,.feedback-title{{font-weight:700;font-size:14px;text-align:center;margin:8px 0 5px}}
th,td{{border:{tables.get('borderWidthPx',1)}px solid #{colors.get('black','000000')};padding:{tables.get('cellPaddingVerticalPx',5)}px {tables.get('cellPaddingHorizontalPx',6)}px;text-align:center;vertical-align:middle;word-break:break-word}}
thead th,.platform-header th{{background:#{colors.get('headerFill','D9D9D9')};font-weight:700}}
.tracked-row td{{background:#{colors.get('trackedFill','FCE4D6')}}}.total-row>*{{font-weight:700}}
.sales-table th:first-child,.sales-table td:first-child{{width:34%}} .jd-table th:nth-child(2){{width:54%}}
.source-note{{font-size:{fonts.get('smallSizePx',10)}px;color:#{colors.get('muted','666666')};text-align:center;margin-top:-5px}}
.missing-note{{border-left:3px solid #{colors.get('headerFill','D9D9D9')};padding:6px 10px;color:#{colors.get('muted','666666')}}}
.product-section{{margin-top:28px}} .product-info th{{width:22%;background:#{colors.get('labelFill','E7EFFA')}}}.product-info td{{text-align:center}}
.product-info .long td{{text-align:justify}} .image-row td{{height:{product_image.get('maxHeightPx',220) + 24}px}}
.product-image{{display:block;max-width:{product_image.get('maxWidthPx',300)}px;max-height:{product_image.get('maxHeightPx',220)}px;margin:0 auto;object-fit:contain}}
.missing-inline,.empty-platform{{color:#{colors.get('muted','666666')}}}.feedback-kpis{{display:grid;grid-template-columns:repeat(4,auto);justify-content:center;border:1px solid #{colors.get('black','000000')};border-bottom:0;margin-top:8px}}
.feedback-kpis span,.feedback-kpis strong{{padding:5px 9px;border-right:1px solid #{colors.get('black','000000')}}}.feedback-kpis strong:last-child{{border-right:0}}
.feedback-table td:nth-child(1),.feedback-table td:nth-child(3){{width:38%}} .feedback-table td:nth-child(2),.feedback-table td:nth-child(4){{width:12%}}
.platform-total th{{font-weight:700}} .product-section,.product-info,.feedback-kpis{{break-inside:avoid}}
@media(max-width:820px){{html,body{{background:#{page.get('background','FFFFFF')}}}.report{{width:100%;margin:0;min-height:0;padding:24px 18px;box-shadow:none}}.table-scroll table{{min-width:680px}}h1{{font-size:19px}}}}
@media print{{@page{{size:A4 portrait;margin:10mm}}html,body{{background:#fff}}.report{{width:auto;max-width:none;min-height:0;margin:0;padding:0;box-shadow:none}}.table-scroll{{overflow:visible}}.table-scroll table{{min-width:0}}thead{{display:table-header-group}}tr{{break-inside:avoid}}}}
"""
    document = f"""<!doctype html>
<html lang="zh-CN"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>{_escape(document_title)}</title><style>{css}</style></head>
<body><main class="report">{logo_html}<h1>{_escape(title_text)}</h1>
<p class="lead">以下是{_escape(brand)} {launch_text}新品销量表现及消费者评论情况：</p>
<h3>新品销量表现（30日）</h3><ul class="sales-summary">{''.join(intro_rows)}</ul>
{_delivery_html(brand, delivery)}{_jd_html(brand, jd_rows, jd_total)}{''.join(product_sections)}
</main></body></html>"""
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(document, encoding="utf-8")
    return ReportBuildResult(output_path, tuple(dict.fromkeys(warnings)))
