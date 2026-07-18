from __future__ import annotations

from collections import Counter
from dataclasses import dataclass
from pathlib import Path
import re
from typing import Any, Iterable

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.page import PageMargins


SOCIAL_PLATFORMS = (
    ("weibo", "微博"),
    ("xiaohongshu", "小红书"),
    ("douyin", "抖音"),
)


@dataclass(frozen=True)
class PlatformSummary:
    key: str
    label: str
    positive_tags: tuple[tuple[str, int], ...]
    negative_tags: tuple[tuple[str, int], ...]
    positive_users: int
    negative_users: int


def _text(value: Any) -> str:
    return str(value or "").strip()


def _header_index(header: list[Any], name: str) -> int | None:
    normalized = [_text(value).replace(" ", "") for value in header]
    target = name.replace(" ", "")
    return next((index for index, value in enumerate(normalized) if value == target), None)


def summarize_social_rows(key: str, label: str, rows: Iterable[Iterable[Any]]) -> PlatformSummary:
    values = [list(row) for row in rows]
    if not values:
        return PlatformSummary(key, label, (), (), 0, 0)
    header = values[0]
    sentiment_index = _header_index(header, "情感识别")
    if sentiment_index is None:
        raise RuntimeError(f"{label}数据缺少“情感识别”列。")

    label_columns: list[tuple[int, int]] = []
    normalized = [_text(value).replace(" ", "") for value in header]
    for index, column_name in enumerate(normalized):
        match = re.fullmatch(r"评价(\d+)-好/差评", column_name)
        if not match:
            continue
        tag_name = f"评价{match.group(1)}-对应标签"
        tag_index = next((i for i, value in enumerate(normalized) if value == tag_name), None)
        if tag_index is not None:
            label_columns.append((index, tag_index))
    if not label_columns:
        raise RuntimeError(f"{label}数据缺少“评价n-好/差评”和对应标签列。")

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
        for polarity_index, tag_index in label_columns:
            polarity = _text(row[polarity_index] if polarity_index < len(row) else "")
            tag = _text(row[tag_index] if tag_index < len(row) else "")
            if not tag:
                continue
            if polarity in {"好评", "正向"}:
                positive_tags[tag] += 1
            elif polarity in {"差评", "负向"}:
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
        worksheet = workbook["Sheet1"] if "Sheet1" in workbook.sheetnames else workbook.worksheets[0]
        return summarize_social_rows(key, label, worksheet.iter_rows(values_only=True))
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


def build_social_feedback_workbook(
    *,
    brand: str,
    product: str,
    start_date: str,
    end_date: str,
    summaries: list[PlatformSummary],
    output_dir: Path,
) -> Path:
    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = output_dir / f"{safe_filename(brand)}-{safe_filename(product)}-社媒评论统计.xlsx"
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = safe_sheet_name(product)
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
        cell.font = Font(name="微软雅黑", size=11, color="1F4E79")

    current_row = 2
    for summary in summaries:
        header_row = current_row
        worksheet.cell(header_row, 1, f"{summary.label}好评")
        worksheet.cell(header_row, 2, "评论数")
        worksheet.cell(header_row, 3, f"{summary.label}差评")
        worksheet.cell(header_row, 4, "评论数")
        _apply_table_style(worksheet, f"A{header_row}:D{header_row}", fill=header_fill, bold=True)
        current_row += 1

        detail_count = max(len(summary.positive_tags), len(summary.negative_tags), 1)
        for index in range(detail_count):
            positive = summary.positive_tags[index] if index < len(summary.positive_tags) else None
            negative = summary.negative_tags[index] if index < len(summary.negative_tags) else None
            worksheet.cell(current_row, 1, positive[0] if positive else "/" if index == 0 and not summary.positive_tags else "")
            worksheet.cell(current_row, 2, positive[1] if positive else 0 if index == 0 and not summary.positive_tags else "")
            worksheet.cell(current_row, 3, negative[0] if negative else "/" if index == 0 and not summary.negative_tags else "")
            worksheet.cell(current_row, 4, negative[1] if negative else 0 if index == 0 and not summary.negative_tags else "")
            current_row += 1
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
    workbook.save(output_path)
    return output_path
