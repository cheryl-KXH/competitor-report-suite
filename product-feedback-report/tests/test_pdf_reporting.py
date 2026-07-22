from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from scripts.reporting.pdf import render_html_to_pdf


class PdfReportingTests(unittest.TestCase):
    def test_render_html_to_pdf_uses_headless_browser_and_validates_output(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            html = root / "报告.html"
            pdf = root / "报告.pdf"
            html.write_text("<html><body>报告</body></html>", encoding="utf-8")

            class BrowserProcess:
                def __init__(self, command, **_kwargs):
                    self.command = command
                    self.returncode = None
                    output_arg = next(
                        item for item in command if item.startswith("--print-to-pdf=")
                    )
                    Path(output_arg.split("=", 1)[1]).write_bytes(
                        b"%PDF-1.4\ncontent\n%%EOF\n"
                    )

                def poll(self):
                    return self.returncode

                def terminate(self):
                    self.returncode = -15

                def kill(self):
                    self.returncode = -9

                def wait(self, timeout=None):
                    return self.returncode

            with (
                patch("scripts.reporting.pdf._browser_executable", return_value=Path("/browser")),
                patch("scripts.reporting.pdf.subprocess.Popen", side_effect=BrowserProcess) as execute,
            ):
                actual = render_html_to_pdf(html, pdf, config={"timeoutSeconds": 60})

        self.assertEqual(actual, pdf.resolve())
        command = execute.call_args.args[0]
        self.assertIn("--headless=new", command)
        self.assertIn("--disable-background-networking", command)
        self.assertIn("--no-first-run", command)
        self.assertIn("--no-pdf-header-footer", command)
        self.assertIn("--window-size=1280,900", command)
        self.assertEqual(command[-1], html.resolve().as_uri())

    def test_render_html_to_pdf_rejects_missing_html(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            missing = Path(tmp) / "missing.html"
            with self.assertRaisesRegex(RuntimeError, "HTML 报告不存在"):
                render_html_to_pdf(missing)


if __name__ == "__main__":
    unittest.main()
