from __future__ import annotations

import tempfile
import unittest
from datetime import date
from pathlib import Path
from unittest.mock import patch

from openpyxl import load_workbook

from scripts.delivery.prepare_product_menu import prepare_product_menu, recent_launch_dates
from scripts.delivery.processing import RawSaleRow


def sale(platform: str, brand: str, product: str, crawl_date: date) -> RawSaleRow:
    return RawSaleRow(platform, brand, "门店", "", product, 1, crawl_date)


class PrepareProductMenuTests(unittest.TestCase):
    def test_recent_launch_query_uses_brand_and_minimum_crawl_date_window(self) -> None:
        rows = [
            sale("美团", "品牌A", "产品A", date(2026, 7, 1)),
            sale("饿了么", "品牌A", "产品A", date(2026, 7, 5)),
            sale("美团", "品牌A", "产品B", date(2026, 7, 5)),
        ]
        configs = {
            "dingtalk": {},
            "report_rules": {
                "productInfoSource": {
                    "baseId": "base",
                    "tableId": "table",
                    "brandFieldId": "brand-field",
                    "launchDateFieldId": "launch-field",
                    "fields": {"productName": {"fieldId": "product-field"}},
                }
            },
        }
        responses = [
            {
                "data": {
                    "fields": [
                        {
                            "fieldId": "brand-field",
                            "config": {"options": [{"name": "品牌A", "id": "brand-option"}]},
                        }
                    ]
                }
            },
            {
                "data": {
                    "records": [
                        {
                            "cells": {
                                "brand-field": {"name": "品牌A"},
                                "product-field": "产品A",
                                "launch-field": "2026-06-02",
                            }
                        }
                    ],
                    "nextCursor": "next-page",
                }
            },
            {
                "data": {
                    "records": [
                        {
                            "cells": {
                                "brand-field": {"name": "品牌A"},
                                "product-field": "产品B",
                                "launch-field": "2026-07-10",
                            }
                        },
                        {
                            "cells": {
                                "brand-field": {"name": "品牌A"},
                                "product-field": "不在外卖清单",
                                "launch-field": "2026-07-01",
                            }
                        },
                        {
                            "cells": {
                                "brand-field": {"name": "品牌A"},
                                "product-field": "产品A",
                                "launch-field": "2026-06-01",
                            }
                        },
                    ]
                }
            },
        ]

        with patch(
            "scripts.delivery.prepare_product_menu.dingtalk_table.call_table_tool",
            side_effect=responses,
        ) as call_tool:
            matched = recent_launch_dates(configs, rows, date(2026, 7, 10))

        self.assertEqual(
            matched,
            {
                ("品牌A", "产品A"): date(2026, 6, 2),
                ("品牌A", "产品B"): date(2026, 7, 10),
            },
        )
        expected_filters = {
            "operator": "and",
            "operands": [
                {"operator": "eq", "operands": ["brand-field", "brand-option"]},
                {"operator": "not_before", "operands": ["launch-field", "2026-06-02"]},
                {"operator": "not_after", "operands": ["launch-field", "2026-07-10"]},
            ],
        }
        self.assertEqual(call_tool.call_args_list[1].args[1], "query_records")
        self.assertEqual(call_tool.call_args_list[1].args[2]["filters"], expected_filters)
        self.assertEqual(
            call_tool.call_args_list[1].args[2]["fieldIds"],
            ["brand-field", "product-field", "launch-field"],
        )
        self.assertNotIn("cursor", call_tool.call_args_list[1].args[2])
        self.assertEqual(call_tool.call_args_list[2].args[2]["cursor"], "next-page")

    def test_prepare_menu_dedupes_by_platform_and_keeps_latest_date_in_filename(self) -> None:
        rows = [
            sale("美团", "品牌A", "产品A", date(2026, 7, 1)),
            sale("美团", "品牌A", "产品A", date(2026, 7, 5)),
            sale("饿了么", "品牌A", "产品A", date(2026, 7, 5)),
        ]
        configs = {"report_rules": {"outputDirectory": "outputs"}}
        progress: list[str] = []
        with tempfile.TemporaryDirectory() as tmp:
            output_dir = Path(tmp)
            with (
                patch("scripts.delivery.prepare_product_menu.load_configs", return_value=configs),
                patch("scripts.delivery.prepare_product_menu.find_delivery_rows", return_value=rows),
                patch(
                    "scripts.delivery.prepare_product_menu.recent_launch_dates",
                    return_value={("品牌A", "产品A"): date(2026, 6, 20)},
                ) as recent,
            ):
                output = prepare_product_menu(
                    "record",
                    output_dir,
                    output_dir,
                    progress_callback=progress.append,
                    report_date=date(2026, 7, 10),
                )
            values = list(load_workbook(output, data_only=True).active.values)

        self.assertEqual(output.name, "品牌A-20260705-产品清单.xlsx")
        self.assertEqual(values[0], ("平台", "品牌", "产品", "近30日上新日期"))
        self.assertEqual(
            values[1:],
            [
                ("美团", "品牌A", "产品A", "2026年06月20日"),
                ("饿了么", "品牌A", "产品A", "2026年06月20日"),
            ],
        )
        recent.assert_called_once_with(configs, rows, date(2026, 7, 10))
        self.assertEqual(
            progress,
            ["2/4 正在提取产品清单", "3/4 正在标记上新不满30天的产品"],
        )

    def test_unknown_brand_generates_no_automatic_annotations(self) -> None:
        rows = [sale("美团", "未配置品牌", "产品A", date(2026, 7, 1))]
        configs = {
            "dingtalk": {},
            "report_rules": {
                "productInfoSource": {
                    "baseId": "base",
                    "tableId": "table",
                }
            },
        }
        fields_response = {
            "data": {
                "fields": [
                    {
                        "fieldId": "01ZM8y7",
                        "config": {"options": [{"name": "其他品牌", "id": "other-option"}]},
                    }
                ]
            }
        }
        with patch(
            "scripts.delivery.prepare_product_menu.dingtalk_table.call_table_tool",
            return_value=fields_response,
        ) as call_tool:
            matched = recent_launch_dates(configs, rows, date(2026, 7, 10))

        self.assertEqual(matched, {})
        call_tool.assert_called_once()
        self.assertEqual(call_tool.call_args.args[1], "get_fields")

    def test_report_date_before_query_start_is_rejected(self) -> None:
        rows = [sale("美团", "品牌A", "产品A", date(2026, 7, 31))]
        with self.assertRaisesRegex(RuntimeError, "报告日期.*早于新品查询起始日期"):
            recent_launch_dates({}, rows, date(2026, 6, 30))


if __name__ == "__main__":
    unittest.main()
