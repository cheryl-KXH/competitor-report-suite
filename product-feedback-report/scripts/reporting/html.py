"""HTML 跟踪报告排版。"""

from __future__ import annotations

import html
from pathlib import Path
from typing import Any

from openpyxl import load_workbook


def read_table(path: Path, max_rows: int | None = None) -> list[list[Any]]:
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    rows: list[list[Any]] = []
    for idx, row in enumerate(ws.iter_rows(values_only=True), 1):
        if max_rows and idx > max_rows:
            break
        rows.append([cell for cell in row])
    return rows


def table_html(rows: list[list[Any]], highlight_products: set[str] | None = None) -> str:
    if not rows:
        return "<p class=\"empty\">暂无有效数据</p>"
    highlight_products = highlight_products or set()
    parts = ["<table>"]
    for row_idx, row in enumerate(rows):
        tag = "th" if row_idx == 0 else "td"
        row_text = {str(cell) for cell in row if cell is not None}
        class_name = " class=\"highlight\"" if highlight_products & row_text else ""
        parts.append(f"<tr{class_name}>")
        for cell in row:
            value = "" if cell is None else cell
            if isinstance(value, float):
                value = f"{value:.1%}" if 0 <= value <= 1 else f"{value:.1f}"
            parts.append(f"<{tag}>{html.escape(str(value))}</{tag}>")
        parts.append("</tr>")
    parts.append("</table>")
    return "\n".join(parts)


def build_report_html(
    *,
    title: str,
    brand: str,
    products: list[str],
    delivery_path: Path | None,
    jd_path: Path | None,
    social_paths: dict[str, Path],
    output_path: Path,
) -> Path:
    highlight = set(products)
    sections: list[str] = []
    if delivery_path and delivery_path.exists():
        sections.append("<section><h2>美团&饿了么外卖数据</h2>" + table_html(read_table(delivery_path), highlight) + "</section>")
    if jd_path and jd_path.exists():
        sections.append("<section><h2>京东外卖数据</h2>" + table_html(read_table(jd_path), highlight) + "</section>")
    for label, path in social_paths.items():
        if path.exists():
            sections.append(f"<section><h2>{html.escape(label)}</h2>{table_html(read_table(path), highlight)}</section>")
    product_list = "、".join(products)
    document = f"""<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>{html.escape(title)}</title>
  <style>
    body {{ margin: 0; background: #f4f4f2; color: #1f1f1f; font-family: -apple-system, BlinkMacSystemFont, "PingFang SC", "Microsoft YaHei", sans-serif; }}
    main {{ max-width: 980px; margin: 0 auto; padding: 48px 32px 80px; background: #fff; min-height: 100vh; }}
    h1 {{ text-align: center; font-size: 28px; margin: 0 0 28px; }}
    h2 {{ font-size: 20px; margin: 32px 0 12px; }}
    .summary {{ font-size: 16px; line-height: 1.8; margin-bottom: 24px; }}
    table {{ border-collapse: collapse; width: 100%; table-layout: fixed; margin: 12px 0 18px; }}
    th, td {{ border: 1px solid #333; padding: 7px 8px; font-size: 14px; text-align: center; vertical-align: middle; word-break: break-word; }}
    th {{ background: #d9d9d9; font-weight: 700; }}
    tr.highlight td {{ background: #fce4d6; }}
    .empty {{ color: #777; }}
  </style>
</head>
<body>
  <main>
    <h1>{html.escape(brand)}：{html.escape(product_list)}</h1>
    <p class="summary">以下是 {html.escape(brand)} 竞品新品销量表现及消费者评论情况。本报告由整理后的平台数表生成，目标新品在表格中高亮。</p>
    {''.join(sections)}
  </main>
</body>
</html>
"""
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(document, encoding="utf-8")
    return output_path
