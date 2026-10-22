from __future__ import annotations

from collections import Counter
from copy import copy
from dataclasses import dataclass
from pathlib import Path
import re
import unicodedata
from typing import Any, Iterable

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.page import PageMargins


SOCIAL_RAW_PLATFORMS = (
    ("weibo", "微博"),
    ("xiaohongshu", "小红书"),
    ("douyin", "抖音"),
    ("bilibili", "B站"),
)

SOCIAL_RAW_SHEET_ORDER = (
    "大众点评-截图",
    "大众点评",
    *(label for _, label in SOCIAL_RAW_PLATFORMS),
)

SOCIAL_PLATFORM_ALIASES = {
    "微博": "微博",
    "微博app": "微博",
    "weibo": "微博",
    "小红书": "小红书",
    "小红书app": "小红书",
    "xiaohongshu": "小红书",
    "抖音": "抖音",
    "抖音app": "抖音",
    "douyin": "抖音",
    "douyinapp": "抖音",
    "b站": "B站",
    "b站app": "B站",
    "哔哩哔哩": "B站",
    "bilibili": "B站",
}


@dataclass(frozen=True)
class PlatformSummary:
    key: str
    label: str
    positive_tags: tuple[tuple[str, int], ...]
    negative_tags: tuple[tuple[str, int], ...]
    positive_users: int
    negative_users: int

    @property
    def positive_count(self) -> int:
        return sum(count for _, count in self.positive_tags)

    @property
    def negative_count(self) -> int:
        return sum(count for _, count in self.negative_tags)


@dataclass(frozen=True)
class SocialPlatformRows:
    worksheet: Any
    header_row: int
    row_numbers: tuple[int, ...]
    is_dedicated_sheet: bool = False


def _text(value: Any) -> str:
    return str(value or "").strip()


def _header_index(header: list[Any], name: str) -> int | None:
    normalize = lambda value: re.sub(
        r"\s+", "", unicodedata.normalize("NFKC", _text(value))
    )
    normalized = [normalize(value) for value in header]
    target = normalize(name)
    return next((index for index, value in enumerate(normalized) if value == target), None)


def _normalized_header(value: Any) -> str:
    return re.sub(r"\s+", "", unicodedata.normalize("NFKC", _text(value))).casefold()


def _platform_label(value: Any) -> str | None:
    normalized = re.sub(r"[\s_-]+", "", unicodedata.normalize("NFKC", _text(value))).casefold()
    return SOCIAL_PLATFORM_ALIASES.get(normalized)


def _worksheet_values(worksheet) -> list[list[Any]]:
    return [list(row) for row in worksheet.iter_rows(values_only=True)]


def _find_social_header_row(values: list[list[Any]], *, require_site: bool = False) -> int | None:
    for row_index, row in enumerate(values, start=1):
        normalized = [_normalized_header(value) for value in row]
        if "产品名称" not in normalized or "情感识别" not in normalized:
            continue
        if require_site and not ({"站点", "平台"} & set(normalized)):
            continue
        if any(re.fullmatch(r"评价\d+-对应标签", name) for name in normalized):
            return row_index
    return None


def _nonempty_row_numbers(values: list[list[Any]], *, after_row: int) -> tuple[int, ...]:
    return tuple(
        row_number
        for row_number, row in enumerate(values, start=1)
        if row_number > after_row and any(_text(value) for value in row)
    )


def _resolve_social_platform_rows(workbook, *, filename: str) -> dict[str, tuple[SocialPlatformRows, ...]]:
    """Resolve the four social platforms from dedicated or mixed-platform sheets."""
    labels = [label for _, label in SOCIAL_RAW_PLATFORMS]
    dedicated: dict[str, list[SocialPlatformRows]] = {label: [] for label in labels}
    mixed: dict[str, list[SocialPlatformRows]] = {label: [] for label in labels}
    empty_dedicated: dict[str, SocialPlatformRows] = {}

    for worksheet in workbook.worksheets:
        if worksheet.title in ("大众点评-截图", "大众点评"):
            continue
        values = _worksheet_values(worksheet)
        header_row = _find_social_header_row(values)
        mixed_header_row = _find_social_header_row(values, require_site=True)
        nonempty_rows = [row for row in values if any(_text(value) for value in row)]

        if mixed_header_row is not None:
            header = values[mixed_header_row - 1]
            site_index = next(
                index
                for index, value in enumerate(header)
                if _normalized_header(value) in {"站点", "平台"}
            )
            rows_by_label: dict[str, list[int]] = {label: [] for label in labels}
            unknown_sites: list[str] = []
            for row_number in _nonempty_row_numbers(values, after_row=mixed_header_row):
                row = values[row_number - 1]
                site = row[site_index] if site_index < len(row) else None
                label = _platform_label(site)
                if label is None:
                    if _text(site):
                        unknown_sites.append(_text(site))
                    continue
                rows_by_label[label].append(row_number)
            if unknown_sites:
                raise RuntimeError(
                    f"社媒清洗数据的“{worksheet.title}”工作表包含无法识别的站点："
                    f"{'、'.join(dict.fromkeys(unknown_sites))}（文件：{filename}）"
                )
            for label, row_numbers in rows_by_label.items():
                if row_numbers:
                    mixed[label].append(
                        SocialPlatformRows(worksheet, mixed_header_row, tuple(row_numbers))
                    )
            continue

        if worksheet.title not in labels:
            continue
        if header_row is not None:
            row_numbers = _nonempty_row_numbers(values, after_row=header_row)
            source = SocialPlatformRows(
                worksheet, header_row, row_numbers, is_dedicated_sheet=True
            )
            if row_numbers:
                dedicated[worksheet.title].append(source)
            else:
                empty_dedicated[worksheet.title] = source
            continue
        if len(nonempty_rows) <= 1:
            empty_dedicated[worksheet.title] = SocialPlatformRows(
                worksheet, 1, (), is_dedicated_sheet=True
            )
            continue
        raise RuntimeError(
            f"社媒清洗数据的“{worksheet.title}”工作表缺少“产品名称/情感识别/"
            f"评价n-对应标签”列：{filename}"
        )

    resolved: dict[str, tuple[SocialPlatformRows, ...]] = {}
    for label in labels:
        if dedicated[label]:
            resolved[label] = tuple(dedicated[label])
        elif mixed[label]:
            resolved[label] = tuple(mixed[label])
        elif label in empty_dedicated:
            resolved[label] = (empty_dedicated[label],)
        else:
            resolved[label] = ()
    return resolved


def _source_rows(source: SocialPlatformRows) -> list[list[Any]]:
    header = [cell.value for cell in source.worksheet[source.header_row]]
    rows = [header]
    rows.extend(
        [cell.value for cell in source.worksheet[row_number]]
        for row_number in source.row_numbers
    )
    return rows


def _reset_read_only_dimensions(workbook) -> None:
    """Ignore unreliable worksheet dimension metadata from third-party exports."""
    for worksheet in workbook.worksheets:
        reset_dimensions = getattr(worksheet, "reset_dimensions", None)
        if reset_dimensions is not None:
            reset_dimensions()


def _is_no_data_rows(rows: Iterable[Iterable[Any]]) -> bool:
    values = [list(row) for row in rows]
    return not any(
        _text(value)
        for row in values[1:]
        for value in row
    )


def normalize_social_product_key(value: Any) -> str:
    """Normalize product labels without allowing fuzzy/substring matching."""
    text = unicodedata.normalize("NFKC", _text(value)).casefold()
    return re.sub(r"[^0-9a-z\u3400-\u9fff]+", "", text)


def social_workbook_products(path: Path) -> dict[str, str]:
    """Return normalized product keys and their first source labels."""
    workbook = load_workbook(path, read_only=True, data_only=False)
    try:
        _reset_read_only_dimensions(workbook)
        worksheet = workbook["Sheet1"] if "Sheet1" in workbook.sheetnames else workbook.worksheets[0]
        values = [list(row) for row in worksheet.iter_rows(values_only=True)]
        header = values[0] if values else []
        product_index = _header_index(header, "产品名称")
        if product_index is None:
            raise RuntimeError(f"社媒文件缺少“产品名称”列：{path.name}")

        products: dict[str, str] = {}
        for row_number, row in enumerate(
            values[1:], start=2
        ):
            if not any(_text(value) for value in row):
                continue
            raw_product = row[product_index] if product_index < len(row) else None
            label = _text(raw_product)
            key = normalize_social_product_key(label)
            if not key:
                raise RuntimeError(
                    f"社媒文件第 {row_number} 行“产品名称”为空：{path.name}"
                )
            products.setdefault(key, label)
        if not products:
            raise RuntimeError(f"社媒文件没有可处理的数据行：{path.name}")
        return products
    finally:
        workbook.close()


def social_cleaned_workbook_products(path: Path) -> dict[str, str]:
    """Return products found across dedicated or mixed social-platform data."""
    workbook = load_workbook(path, read_only=False, data_only=False)
    try:
        products: dict[str, str] = {}
        sources = _resolve_social_platform_rows(workbook, filename=path.name)
        for _, label in SOCIAL_RAW_PLATFORMS:
            for source in sources[label]:
                values = _source_rows(source)
                if _is_no_data_rows(values):
                    continue
                header = values[0]
                product_index = _header_index(header, "产品名称")
                if product_index is None:
                    raise RuntimeError(
                        f"社媒清洗数据的“{source.worksheet.title}”工作表缺少"
                        f"“产品名称”列：{path.name}"
                    )
                for row_offset, row in enumerate(values[1:], start=1):
                    row_number = source.row_numbers[row_offset - 1]
                    if not any(_text(value) for value in row):
                        continue
                    raw_product = row[product_index] if product_index < len(row) else None
                    product = _text(raw_product)
                    product_key = normalize_social_product_key(product)
                    if not product_key:
                        raise RuntimeError(
                            f"社媒清洗数据的“{source.worksheet.title}”工作表第 {row_number} 行"
                            f"“产品名称”为空：{path.name}"
                        )
                    products.setdefault(product_key, product)
        return products
    finally:
        workbook.close()


def split_social_workbook(
    path: Path,
    *,
    product_names: dict[str, str],
    output_dir: Path,
    output_names: dict[str, str],
) -> dict[str, Path]:
    """Create one formatting-preserving workbook per product in the source file."""
    source_products = social_workbook_products(path)
    unknown = [source_products[key] for key in source_products if key not in product_names]
    if unknown:
        raise RuntimeError(
            f"社媒文件包含无法匹配到同品牌同一“30日”记录的产品：{'、'.join(unknown)}"
        )

    output_dir.mkdir(parents=True, exist_ok=True)
    outputs: dict[str, Path] = {}
    for product_key in source_products:
        workbook = load_workbook(path, data_only=False)
        try:
            worksheet = workbook["Sheet1"] if "Sheet1" in workbook.sheetnames else workbook.worksheets[0]
            header = [cell.value for cell in worksheet[1]]
            product_index = _header_index(header, "产品名称")
            if product_index is None:
                raise RuntimeError(f"社媒文件缺少“产品名称”列：{path.name}")
            for row_number in range(worksheet.max_row, 1, -1):
                row_values = [cell.value for cell in worksheet[row_number]]
                if not any(_text(value) for value in row_values):
                    worksheet.delete_rows(row_number)
                    continue
                raw_product = (
                    row_values[product_index]
                    if product_index < len(row_values)
                    else None
                )
                if normalize_social_product_key(raw_product) != product_key:
                    worksheet.delete_rows(row_number)
            output_path = output_dir / output_names[product_key]
            workbook.save(output_path)
            outputs[product_key] = output_path
        finally:
            workbook.close()
    return outputs


def summarize_social_rows(key: str, label: str, rows: Iterable[Iterable[Any]]) -> PlatformSummary:
    values = [list(row) for row in rows]
    if _is_no_data_rows(values):
        return PlatformSummary(key, label, (), (), 0, 0)
    header = values[0]
    sentiment_index = _header_index(header, "情感识别")
    if sentiment_index is None:
        raise RuntimeError(f"{label}数据缺少“情感识别”列。")

    label_columns: list[int] = []
    normalized = [_text(value).replace(" ", "") for value in header]
    for index, column_name in enumerate(normalized):
        if re.fullmatch(r"评价\d+-对应标签", column_name):
            label_columns.append(index)
    if not label_columns:
        raise RuntimeError(f"{label}数据缺少“评价n-对应标签”列。")

    positive_tags: Counter[str] = Counter()
    negative_tags: Counter[str] = Counter()
    positive_users = 0
    negative_users = 0
    for row in values[1:]:
        sentiment = _text(row[sentiment_index] if sentiment_index < len(row) else "")
        if sentiment == "正向":
            positive_users += 1
        elif sentiment == "负向":
            negative_users += 1
        for tag_index in label_columns:
            tag = _text(row[tag_index] if tag_index < len(row) else "")
            if not tag:
                continue
            if sentiment == "正向":
                positive_tags[tag] += 1
            elif sentiment == "负向":
                negative_tags[tag] += 1

    sort_tags = lambda counter: tuple(sorted(counter.items(), key=lambda item: (-item[1], item[0])))
    return PlatformSummary(
        key,
        label,
        sort_tags(positive_tags),
        sort_tags(negative_tags),
        positive_users,
        negative_users,
    )


def summarize_social_file(key: str, label: str, path: Path) -> PlatformSummary:
    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        _reset_read_only_dimensions(workbook)
        worksheet = workbook["Sheet1"] if "Sheet1" in workbook.sheetnames else workbook.worksheets[0]
        return summarize_social_rows(key, label, worksheet.iter_rows(values_only=True))
    finally:
        workbook.close()


def _count_value(value: Any, *, label: str) -> int:
    if isinstance(value, bool):
        raise RuntimeError(f"{label}不是有效整数：{value}")
    if isinstance(value, (int, float)) and float(value).is_integer():
        return int(value)
    text = _text(value).replace(",", "")
    try:
        parsed = float(text)
    except ValueError as exc:
        raise RuntimeError(f"{label}不是有效整数：{value}") from exc
    if not parsed.is_integer():
        raise RuntimeError(f"{label}不是有效整数：{value}")
    return int(parsed)


def _find_dianping_header_row(values: list[list[Any]]) -> int | None:
    for row_number, row in enumerate(values):
        normalized = [_text(value).replace(" ", "") for value in row]
        try:
            positive_index = normalized.index("大众点评好评")
            negative_index = normalized.index("大众点评差评")
        except ValueError:
            continue
        has_positive_count = any(
            normalized[index] == "评论数"
            for index in range(positive_index + 1, len(normalized))
        )
        has_negative_count = any(
            normalized[index] == "评论数"
            for index in range(negative_index + 1, len(normalized))
        )
        if has_positive_count and has_negative_count:
            return row_number
    return None


def summarize_dianping_rows(rows: Iterable[Iterable[Any]]) -> PlatformSummary:
    values = [list(row) for row in rows]
    if _is_no_data_rows(values):
        return PlatformSummary("dianping", "大众点评", (), (), 0, 0)
    header_row = _find_dianping_header_row(values)
    indexes: tuple[int, int, int, int] | None = None
    for row_number, row in enumerate(values):
        if row_number != header_row:
            continue
        normalized = [_text(value).replace(" ", "") for value in row]
        positive_index = normalized.index("大众点评好评")
        negative_index = normalized.index("大众点评差评")
        positive_count_index = next(
            (index for index in range(positive_index + 1, len(normalized)) if normalized[index] == "评论数"),
            None,
        )
        negative_count_index = next(
            (index for index in range(negative_index + 1, len(normalized)) if normalized[index] == "评论数"),
            None,
        )
        if positive_count_index is not None and negative_count_index is not None:
            header_row = row_number
            indexes = (
                positive_index,
                positive_count_index,
                negative_index,
                negative_count_index,
            )
        break
    if header_row is None or indexes is None:
        raise RuntimeError("大众点评工作表缺少统计表头。")

    positive_tags: Counter[str] = Counter()
    negative_tags: Counter[str] = Counter()
    positive_users: int | None = None
    negative_users: int | None = None
    positive_index, positive_count_index, negative_index, negative_count_index = indexes
    for row in values[header_row + 1 :]:
        padded = row + [None] * max(0, negative_count_index + 1 - len(row))
        positive_label = _text(padded[positive_index])
        negative_label = _text(padded[negative_index])
        if positive_label == "好评用户数" and negative_label == "差评用户数":
            positive_users = _count_value(
                padded[positive_count_index], label="大众点评好评用户数"
            )
            negative_users = _count_value(
                padded[negative_count_index], label="大众点评差评用户数"
            )
            break
        if positive_label and positive_label != "/":
            positive_tags[positive_label] += _count_value(
                padded[positive_count_index], label=f"大众点评好评标签“{positive_label}”评论数"
            )
        if negative_label and negative_label != "/":
            negative_tags[negative_label] += _count_value(
                padded[negative_count_index], label=f"大众点评差评标签“{negative_label}”评论数"
            )
    if positive_users is None or negative_users is None:
        raise RuntimeError("大众点评工作表缺少“好评用户数/差评用户数”汇总行。")

    sort_tags = lambda counter: tuple(sorted(counter.items(), key=lambda item: (-item[1], item[0])))
    return PlatformSummary(
        "dianping",
        "大众点评",
        sort_tags(positive_tags),
        sort_tags(negative_tags),
        positive_users,
        negative_users,
    )


def summarize_social_cleaned_workbook(path: Path) -> list[PlatformSummary]:
    """Read Dianping plus dedicated or mixed rows for the four social platforms."""
    workbook = load_workbook(path, read_only=False, data_only=True)
    try:
        if "大众点评" not in workbook.sheetnames:
            raise RuntimeError(
                f"社媒清洗数据缺少工作表：大众点评（文件：{path.name}）"
            )
        summaries = [
            summarize_dianping_rows(workbook["大众点评"].iter_rows(values_only=True))
        ]
        sources = _resolve_social_platform_rows(workbook, filename=path.name)
        for key, label in SOCIAL_RAW_PLATFORMS:
            platform_sources = sources[label]
            if not platform_sources or not any(source.row_numbers for source in platform_sources):
                summaries.append(empty_platform_summary(key, label))
                continue
            partials = [
                summarize_social_rows(key, label, _source_rows(source))
                for source in platform_sources
                if source.row_numbers
            ]
            positive_tags: Counter[str] = Counter()
            negative_tags: Counter[str] = Counter()
            for partial in partials:
                positive_tags.update(dict(partial.positive_tags))
                negative_tags.update(dict(partial.negative_tags))
            sort_tags = lambda counter: tuple(
                sorted(counter.items(), key=lambda item: (-item[1], item[0]))
            )
            summaries.append(
                PlatformSummary(
                    key,
                    label,
                    sort_tags(positive_tags),
                    sort_tags(negative_tags),
                    sum(partial.positive_users for partial in partials),
                    sum(partial.negative_users for partial in partials),
                )
            )
        return summaries
    finally:
        workbook.close()


def empty_platform_summary(key: str, label: str) -> PlatformSummary:
    return PlatformSummary(key, label, (), (), 0, 0)


def safe_filename(value: str) -> str:
    cleaned = re.sub(r'[\\/:*?"<>|\r\n]+', "_", value).strip().strip(".")
    return cleaned or "社媒评论统计"


def safe_sheet_name(value: str) -> str:
    cleaned = re.sub(r"[\\/*?:\[\]]+", "_", value).strip().strip("'")
    return (cleaned or "社媒评论统计")[:31]


def _apply_table_style(worksheet, cell_range: str, *, fill: PatternFill | None = None, bold: bool = False) -> None:
    thin = Side(style="thin", color="404040")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for row in worksheet[cell_range]:
        for cell in row:
            cell.border = border
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.font = Font(name="微软雅黑", size=11, bold=bold)
            if fill:
                cell.fill = fill


def _copy_cell(source_cell, target_cell) -> None:
    target_cell.value = source_cell.value
    if source_cell.has_style:
        target_cell.font = copy(source_cell.font)
        target_cell.fill = copy(source_cell.fill)
        target_cell.border = copy(source_cell.border)
        target_cell.alignment = copy(source_cell.alignment)
        target_cell.protection = copy(source_cell.protection)
    if source_cell.number_format:
        target_cell.number_format = source_cell.number_format
    if source_cell.hyperlink:
        target_cell._hyperlink = copy(source_cell.hyperlink)
    if source_cell.comment:
        target_cell.comment = copy(source_cell.comment)


def _copy_row_layout(source, target, source_row: int, target_row: int) -> None:
    if source_row not in source.row_dimensions:
        return
    source_dimension = source.row_dimensions[source_row]
    target_dimension = target.row_dimensions[target_row]
    target_dimension.height = source_dimension.height
    target_dimension.hidden = source_dimension.hidden
    target_dimension.outlineLevel = source_dimension.outlineLevel
    target_dimension.collapsed = source_dimension.collapsed
    target_dimension.thickTop = source_dimension.thickTop
    target_dimension.thickBot = source_dimension.thickBot


def _copy_column_layout(source_dimension, target_dimension) -> None:
    target_dimension.width = source_dimension.width
    target_dimension.hidden = source_dimension.hidden
    target_dimension.bestFit = source_dimension.bestFit
    target_dimension.outlineLevel = source_dimension.outlineLevel
    target_dimension.collapsed = source_dimension.collapsed


def _copy_worksheet_rows(
    source,
    target,
    *,
    row_numbers: tuple[int, ...] | None = None,
    max_column: int | None = None,
    copy_drawings: bool = False,
) -> None:
    selected_rows = (
        tuple(range(1, source.max_row + 1))
        if row_numbers is None
        else row_numbers
    )
    selected_max_column = max_column or source.max_column
    for target_row, source_row in enumerate(selected_rows, start=1):
        for source_cell in source[source_row][:selected_max_column]:
            _copy_cell(source_cell, target.cell(target_row, source_cell.column))
        _copy_row_layout(source, target, source_row, target_row)
    for key, dimension in source.column_dimensions.items():
        if dimension.min > selected_max_column:
            continue
        _copy_column_layout(dimension, target.column_dimensions[key])
    target.sheet_properties = copy(source.sheet_properties)
    target.sheet_view.showGridLines = source.sheet_view.showGridLines
    target.freeze_panes = source.freeze_panes
    target.sheet_state = source.sheet_state
    if row_numbers is None:
        for merged_range in source.merged_cells.ranges:
            target.merge_cells(str(merged_range))
        target.auto_filter = copy(source.auto_filter)
        target.print_options = copy(source.print_options)
        target.page_margins = copy(source.page_margins)
        target.page_setup = copy(source.page_setup)
        target.sheet_properties.pageSetUpPr = copy(source.sheet_properties.pageSetUpPr)
        target.print_area = source.print_area
        target.print_title_rows = source.print_title_rows
        target.print_title_cols = source.print_title_cols
        if copy_drawings:
            for image in source._images:
                target.add_image(copy(image))
            for chart in source._charts:
                target.add_chart(copy(chart))
    else:
        target.freeze_panes = "A2" if len(selected_rows) > 1 else None
        target.auto_filter.ref = f"A1:{target.cell(target.max_row, target.max_column).coordinate}"


def _normalize_dianping_raw_table(
    worksheet,
    *,
    min_column: int,
    max_column: int,
) -> None:
    """Render the copied Dianping table as an unmerged, fully bordered grid."""
    thin = Side(style="thin", color="404040")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for row in worksheet.iter_rows(
        min_row=1,
        max_row=worksheet.max_row,
        min_col=min_column,
        max_col=max_column,
    ):
        for cell in row:
            if _text(cell.value) == "/":
                cell.value = None
            cell.border = border


def _source_max_content_column(source: SocialPlatformRows) -> int:
    source_rows = (source.header_row, *source.row_numbers)
    populated_columns = [
        cell.column
        for row_number in source_rows
        for cell in source.worksheet[row_number]
        if cell.value is not None and _text(cell.value)
    ]
    return max(populated_columns, default=1)


def _append_source_rows(
    source: SocialPlatformRows,
    target,
    *,
    include_header: bool,
    max_column: int,
) -> None:
    source_rows = ((source.header_row,) if include_header else ()) + source.row_numbers
    for source_row in source_rows:
        target_row = target.max_row + 1 if target.max_row else 1
        if target.max_row == 1 and target["A1"].value is None:
            target_row = 1
        for source_cell in source.worksheet[source_row][:max_column]:
            _copy_cell(source_cell, target.cell(target_row, source_cell.column))
        _copy_row_layout(source.worksheet, target, source_row, target_row)


def _add_raw_social_sheets(output_workbook, raw_data_path: Path) -> None:
    source_workbook = load_workbook(raw_data_path, data_only=False)
    try:
        sources = _resolve_social_platform_rows(
            source_workbook, filename=raw_data_path.name
        )
        screenshot_target = output_workbook.create_sheet("大众点评-截图")
        if "大众点评-截图" in source_workbook.sheetnames:
            _copy_worksheet_rows(
                source_workbook["大众点评-截图"],
                screenshot_target,
                copy_drawings=True,
            )

        dianping_target = output_workbook.create_sheet("大众点评")
        if "大众点评" in source_workbook.sheetnames:
            dianping_source = source_workbook["大众点评"]
            dianping_values = _worksheet_values(dianping_source)
            header_index = _find_dianping_header_row(dianping_values)
            populated_rows = tuple(
                row_number
                for row_number in range(1, dianping_source.max_row + 1)
                if any(_text(cell.value) for cell in dianping_source[row_number])
            )
            if header_index is not None and populated_rows:
                first_row, last_row = header_index + 1, populated_rows[-1]
                selected_rows = tuple(range(first_row, last_row + 1))
                populated_header_columns = [
                    cell.column
                    for cell in dianping_source[first_row]
                    if _text(cell.value)
                ]
                max_column = max(
                    cell.column
                    for row_number in selected_rows
                    for cell in dianping_source[row_number]
                    if _text(cell.value)
                )
                _copy_worksheet_rows(
                    dianping_source,
                    dianping_target,
                    row_numbers=selected_rows,
                    max_column=max_column,
                )
                _normalize_dianping_raw_table(
                    dianping_target,
                    min_column=min(populated_header_columns),
                    max_column=max_column,
                )

        for _, label in SOCIAL_RAW_PLATFORMS:
            platform_sources = sources[label]
            target = output_workbook.create_sheet(label)
            if not platform_sources:
                continue
            if (
                len(platform_sources) == 1
                and platform_sources[0].is_dedicated_sheet
                and platform_sources[0].worksheet.title == label
            ):
                source = platform_sources[0]
                _copy_worksheet_rows(
                    source.worksheet,
                    target,
                    row_numbers=(source.header_row, *source.row_numbers),
                    max_column=_source_max_content_column(source),
                )
                continue

            first = platform_sources[0]
            max_column = max(
                _source_max_content_column(source) for source in platform_sources
            )
            for key, dimension in first.worksheet.column_dimensions.items():
                if dimension.min > max_column:
                    continue
                _copy_column_layout(dimension, target.column_dimensions[key])
            for index, source in enumerate(platform_sources):
                _append_source_rows(
                    source,
                    target,
                    include_header=index == 0,
                    max_column=max_column,
                )
            target.freeze_panes = "A2"
            if target.max_column:
                target.auto_filter.ref = f"A1:{target.cell(target.max_row, target.max_column).coordinate}"
    finally:
        source_workbook.close()


def build_social_feedback_workbook(
    *,
    brand: str,
    product: str,
    start_date: str,
    end_date: str,
    summaries: list[PlatformSummary],
    output_dir: Path,
    raw_data_path: Path | None = None,
) -> Path:
    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = output_dir / f"{safe_filename(brand)}-{safe_filename(product)}-社媒评论统计.xlsx"
    workbook = Workbook()
    worksheet = workbook.active
    summary_name = safe_sheet_name(product)
    if summary_name in SOCIAL_RAW_SHEET_ORDER:
        summary_name = safe_sheet_name(f"{summary_name}-统计")
    worksheet.title = summary_name
    worksheet.sheet_view.showGridLines = False

    header_fill = PatternFill("solid", fgColor="D9D7D7")
    worksheet.merge_cells("A1:D1")
    worksheet["A1"] = f"{product} {start_date}-{end_date} 第三方平台评价反馈"
    worksheet["E1"] = "好评率"
    worksheet["E2"] = "总计"
    worksheet["E3"] = "好评用户数"
    worksheet["E4"] = "差评用户数"

    positive_users = sum(summary.positive_users for summary in summaries)
    negative_users = sum(summary.negative_users for summary in summaries)
    total_users = positive_users + negative_users
    worksheet["F1"] = positive_users / total_users if total_users else 0
    worksheet["F2"] = total_users
    worksheet["F3"] = positive_users
    worksheet["F4"] = negative_users
    worksheet["F1"].number_format = "0%"
    for row in range(2, 5):
        worksheet[f"F{row}"].number_format = "#,##0"

    _apply_table_style(worksheet, "A1:D1", fill=header_fill, bold=True)
    _apply_table_style(worksheet, "E1:F4")
    for cell in worksheet["A1:D1"][0]:
        cell.font = Font(name="微软雅黑", size=13, bold=True)
    for cell in worksheet["E"][:4]:
        cell.font = Font(name="微软雅黑", size=11, color="000000")

    current_row = 2
    for summary in summaries:
        header_row = current_row
        worksheet.cell(header_row, 1, f"{summary.label}好评")
        worksheet.cell(header_row, 2, "评论数")
        worksheet.cell(header_row, 3, f"{summary.label}差评")
        worksheet.cell(header_row, 4, "评论数")
        _apply_table_style(worksheet, f"A{header_row}:D{header_row}", fill=header_fill, bold=True)
        current_row += 1

        detail_count = max(len(summary.positive_tags), len(summary.negative_tags))
        for index in range(detail_count):
            positive = summary.positive_tags[index] if index < len(summary.positive_tags) else None
            negative = summary.negative_tags[index] if index < len(summary.negative_tags) else None
            worksheet.cell(current_row, 1, positive[0] if positive else "")
            worksheet.cell(current_row, 2, positive[1] if positive else "")
            worksheet.cell(current_row, 3, negative[0] if negative else "")
            worksheet.cell(current_row, 4, negative[1] if negative else "")
            current_row += 1
        if detail_count:
            _apply_table_style(worksheet, f"A{header_row + 1}:D{current_row - 1}")

        worksheet.cell(current_row, 1, "好评用户数")
        worksheet.cell(current_row, 2, summary.positive_users)
        worksheet.cell(current_row, 3, "差评用户数")
        worksheet.cell(current_row, 4, summary.negative_users)
        _apply_table_style(worksheet, f"A{current_row}:D{current_row}", bold=True)
        current_row += 1

    last_row = current_row - 1
    for column in ("B", "D"):
        for cell in worksheet[column][1:last_row]:
            if isinstance(cell.value, int):
                cell.number_format = "#,##0"
    widths = {"A": 27, "B": 11, "C": 27, "D": 11, "E": 15, "F": 11}
    for column, width in widths.items():
        worksheet.column_dimensions[column].width = width
    for row in range(1, last_row + 1):
        worksheet.row_dimensions[row].height = 26

    worksheet.freeze_panes = "A2"
    worksheet.print_area = f"A1:F{last_row}"
    worksheet.page_setup.paperSize = worksheet.PAPERSIZE_A4
    worksheet.page_setup.orientation = worksheet.ORIENTATION_PORTRAIT
    worksheet.page_setup.fitToWidth = 1
    worksheet.page_setup.fitToHeight = 0
    worksheet.sheet_properties.pageSetUpPr.fitToPage = True
    worksheet.page_margins = PageMargins(left=0.25, right=0.25, top=0.5, bottom=0.5, header=0.2, footer=0.2)
    if raw_data_path is not None:
        _add_raw_social_sheets(workbook, raw_data_path)
    workbook.save(output_path)
    return output_path
