from __future__ import annotations

import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import call, patch


ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from service import dingtalk_docs  # noqa: E402


class ReportFolderReplacementTests(unittest.TestCase):
    def test_regeneration_reuses_folder_and_replaces_only_report_files(self) -> None:
        events: list[str] = []
        with tempfile.TemporaryDirectory() as temp_dir:
            paths = []
            for suffix in (".xlsx", ".html", ".png"):
                path = Path(temp_dir) / f"喜茶26W30_20260724{suffix}"
                path.write_bytes(b"new report")
                paths.append(path)

            with patch.object(
                dingtalk_docs,
                "ensure_report_folder",
                return_value=("stable-folder-id", "https://example.invalid/stable-folder"),
            ) as ensure_mock:
                with patch.object(
                    dingtalk_docs,
                    "delete_report_file_variants",
                    side_effect=lambda *_args: events.append("delete-files"),
                ) as delete_mock:
                    with patch.object(
                        dingtalk_docs,
                        "upload_file",
                        side_effect=lambda _config, _folder_id, path: events.append(f"upload:{path.suffix}"),
                    ) as upload_mock:
                        report_url = dingtalk_docs.upload_report_directory(
                            {},
                            "喜茶",
                            2026,
                            "喜茶26W30_20260724",
                            paths,
                        )

        self.assertEqual(report_url, "https://example.invalid/stable-folder")
        ensure_mock.assert_called_once_with({}, "喜茶", 2026, "喜茶26W30_20260724")
        delete_mock.assert_called_once_with({}, "stable-folder-id", paths)
        self.assertEqual(
            upload_mock.call_args_list,
            [call({}, "stable-folder-id", path) for path in paths],
        )
        self.assertEqual(events, ["delete-files", "upload:.xlsx", "upload:.html", "upload:.png"])

    def test_file_cleanup_removes_exact_names_and_numbered_duplicates_only(self) -> None:
        paths = [
            Path("/tmp/喜茶26W30_20260724.xlsx"),
            Path("/tmp/喜茶26W30_20260724.html"),
        ]
        nodes = [
            {
                "nodeId": "xlsx-current",
                "name": "喜茶26W30_20260724",
                "extension": "xlsx",
            },
            {
                "nodeId": "xlsx-duplicate",
                "name": "喜茶26W30_20260724(2)",
                "extension": "xlsx",
            },
            {"nodeId": "html-current", "name": "喜茶26W30_20260724.html"},
            {"nodeId": "keep-notes", "name": "审核备注.txt"},
            {"nodeId": "keep-other-week", "name": "喜茶26W29_20260717.xlsx"},
        ]

        with patch.object(dingtalk_docs, "list_nodes", side_effect=[nodes, []]):
            with patch.object(dingtalk_docs, "call_docs_tool") as call_mock:
                dingtalk_docs.delete_report_file_variants({}, "stable-folder-id", paths)

        self.assertEqual(
            [
                (item.args[1], item.args[2]["nodeId"])
                for item in call_mock.call_args_list
            ],
            [
                ("rename_document", "xlsx-current"),
                ("delete_document", "xlsx-current"),
                ("rename_document", "xlsx-duplicate"),
                ("delete_document", "xlsx-duplicate"),
                ("rename_document", "html-current"),
                ("delete_document", "html-current"),
            ],
        )

    def test_file_cleanup_waits_until_old_names_disappear(self) -> None:
        paths = [Path("/tmp/喜茶26W30_20260724.xlsx")]
        stale_nodes = [
            {
                "nodeId": "xlsx-current",
                "name": "喜茶26W30_20260724",
                "extension": "xlsx",
            }
        ]

        with patch.object(
            dingtalk_docs,
            "list_nodes",
            side_effect=[stale_nodes, stale_nodes, []],
        ):
            with patch.object(dingtalk_docs, "call_docs_tool"):
                with patch.object(dingtalk_docs.time, "sleep") as sleep_mock:
                    dingtalk_docs.delete_report_file_variants({}, "stable-folder-id", paths)

        sleep_mock.assert_called_once_with(0.5)


if __name__ == "__main__":
    unittest.main()
