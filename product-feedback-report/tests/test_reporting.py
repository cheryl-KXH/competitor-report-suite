from __future__ import annotations

import tempfile
import unittest
from datetime import date
from pathlib import Path

from openpyxl import Workbook

from scripts.reporting.html import (
    ProductInfo,
    build_report_html,
    combine_delivery_sales,
    delivery_report_from_metrics,
    jd_report_from_metrics,
    read_jd_sales,
    read_social_report,
    social_report_from_summaries,
)
from scripts.social.processing import PlatformSummary, build_social_feedback_workbook


class ReportingTests(unittest.TestCase):
    def _platform_workbook(self, path: Path, rows: list[tuple[str, float, float, int]]) -> Path:
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.append(
            ["排名", "商品名称", "日店均销量", "总销量", "总销量占比", "总在售天数", "在售门店数", "在售天数", "上新日期"]
        )
        total = sum(row[2] for row in rows)
        for index, (product, daily, sales, stores) in enumerate(rows, 1):
            worksheet.append([index, product, daily, sales, sales / total if total else 0, 30 * stores, stores, 30, None])
        workbook.save(path)
        return path

    def _jd_workbook(self, path: Path) -> Path:
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.append(["排名", "商品名称", "销量加总", "在售门店数", "总销量", "总销量占比", "在售天数", "上新日期"])
        rows = [(1, "A", 100), (1, "B", 100), (3, "C", 50), (21, "D", 10)]
        total = sum(sales for _, _, sales in rows)
        for rank, product, sales in rows:
            worksheet.append([rank, product, sales, 1, sales, sales / total, 30, None])
        workbook.save(path)
        return path

    def test_delivery_combines_total_sales_and_adds_tracked_product_outside_top20(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            meituan_rows = [(f"商品{i}", 30 - i, 3000 - i * 100, 10) for i in range(1, 22)]
            eleme_rows = [(f"商品{i}", 20 - i / 2, 2000 - i * 50, 12) for i in range(1, 22)]
            meituan_rows[-1] = ("轻因·新品", 0.5, 5, 4)
            eleme_rows[-1] = ("轻因·新品", 0.4, 4, 5)
            mt = self._platform_workbook(root / "美团.xlsx", meituan_rows)
            elm = self._platform_workbook(root / "饿了么.xlsx", eleme_rows)

            report = combine_delivery_sales(mt, elm, ["新品"])

        self.assertEqual(len([row for row in report.display_rows if row.rank <= 20]), 20)
        self.assertTrue(report.display_rows[-1].tracked)
        self.assertEqual(report.tracked_rows["新品"].product, "轻因·新品")
        self.assertAlmostEqual(sum(row.share for row in report.all_rows), 1.0)
        self.assertEqual(report.meituan_stores, 10)
        self.assertEqual(report.eleme_stores, 12)

    def test_ambiguous_tracked_product_match_fails(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            mt = self._platform_workbook(
                root / "美团.xlsx", [("轻因·新品", 1, 10, 1), ("限定·新品", 1, 9, 1)]
            )
            elm = self._platform_workbook(root / "饿了么.xlsx", [("其他", 1, 8, 1)])
            with self.assertRaisesRegex(RuntimeError, "匹配到多个外卖商品"):
                combine_delivery_sales(mt, elm, ["新品"])

    def test_jd_top20_keeps_ties_and_total_uses_all_rows(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = self._jd_workbook(Path(tmp) / "京东.xlsx")
            rows, total = read_jd_sales(path)
        self.assertEqual([row.rank for row in rows], [1, 1, 3])
        self.assertEqual(total, 260)

    def test_social_parser_aggregates_summary_and_keeps_empty_platforms_blank(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = build_social_feedback_workbook(
                brand="霸王茶姬",
                product="糯青山柠檬奶",
                start_date="6.18",
                end_date="7.18",
                summaries=[
                    PlatformSummary("weibo", "微博", (("好喝，喜欢，推荐", 9), ("清爽不腻", 5)), (("味道奇怪", 2),), 14, 2),
                    PlatformSummary("xiaohongshu", "小红书", (("清爽不腻", 3),), (("果味淡", 1),), 3, 1),
                    PlatformSummary("douyin", "抖音", (), (), 0, 0),
                    PlatformSummary("bilibili", "B站", (), (), 0, 0),
                ],
                output_dir=Path(tmp),
            )
            report = read_social_report(path)
        self.assertEqual(report.period, "6.18-7.18")
        self.assertEqual(report.total_users, 20)
        self.assertEqual(report.positive_top[0], ("清爽不腻", 8))
        self.assertNotIn("好喝，喜欢，推荐", [label for label, _ in report.positive_top])
        self.assertEqual(report.sections[-1].positive_tags, ())

    def test_html_is_self_contained_and_omits_zero_placeholders(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            mt = self._platform_workbook(root / "美团.xlsx", [("轻因·新品", 1.2, 120, 5)])
            elm = self._platform_workbook(root / "饿了么.xlsx", [("轻因·新品", 1.4, 140, 6)])
            social = build_social_feedback_workbook(
                brand="品牌",
                product="新品",
                start_date="6.18",
                end_date="7.18",
                summaries=[PlatformSummary(key, label, (), (), 0, 0) for key, label in (("weibo", "微博"), ("xiaohongshu", "小红书"), ("douyin", "抖音"), ("bilibili", "B站"))],
                output_dir=root,
            )
            output = root / "报告.html"
            build_report_html(
                title="报告",
                brand="品牌",
                products=["新品"],
                report_date=date(2026, 7, 18),
                meituan_path=mt,
                eleme_path=elm,
                jd_path=None,
                social_paths={"新品": social},
                product_infos={"新品": ProductInfo("新品", price="18元")},
                launch_dates={"新品": date(2026, 6, 18)},
                output_path=output,
                configs={"html_layout": {}, "font_files": {}},
            )
            document = output.read_text(encoding="utf-8")

        self.assertIn("品牌京东外卖销售数据暂无法获取", document)
        self.assertIn("暂无评论", document)
        self.assertNotIn(">/</td>", document)
        self.assertNotIn(">0</td>", document)
        self.assertNotIn("file://", document)
        self.assertIn("新品销量表现（30日）", document)

    def test_html_accepts_in_memory_statistics_without_intermediate_files(self) -> None:
        delivery = delivery_report_from_metrics(
            [{"product": "新品", "daily_store_avg": 1.2, "sales": 120, "stores": 5}],
            [{"product": "新品", "daily_store_avg": 1.4, "sales": 140, "stores": 6}],
            ["新品"],
        )
        jd = jd_report_from_metrics([])
        social = social_report_from_summaries(
            title="品牌-新品 消费者反馈",
            period="6.18-7.18",
            summaries=[
                PlatformSummary(key, label, (), (), 0, 0)
                for key, label in (
                    ("weibo", "微博"),
                    ("xiaohongshu", "小红书"),
                    ("douyin", "抖音"),
                    ("bilibili", "B站"),
                )
            ],
        )
        with tempfile.TemporaryDirectory() as tmp:
            output = Path(tmp) / "报告.html"
            build_report_html(
                title="报告",
                brand="品牌",
                products=["新品"],
                report_date=date(2026, 7, 18),
                meituan_path=None,
                eleme_path=None,
                jd_path=None,
                social_paths={},
                product_infos={"新品": ProductInfo("新品")},
                launch_dates={"新品": date(2026, 6, 18)},
                output_path=output,
                configs={"html_layout": {}, "font_files": {}},
                delivery_report=delivery,
                jd_report=jd,
                social_report_models={"新品": social},
            )
            document = output.read_text(encoding="utf-8")
        self.assertIn("新品销量表现（30日）", document)


if __name__ == "__main__":
    unittest.main()
