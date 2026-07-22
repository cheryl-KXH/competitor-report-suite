"""使用 Chromium 系浏览器把自包含 HTML 报告打印为 PDF。"""

from __future__ import annotations

import os
import shutil
import subprocess
import tempfile
import time
from pathlib import Path
from typing import Any, Sequence


_BROWSER_COMMANDS = (
    "google-chrome",
    "google-chrome-stable",
    "chromium",
    "chromium-browser",
    "chrome",
    "microsoft-edge",
)

_BROWSER_PATHS = (
    Path("/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"),
    Path("/Applications/Chromium.app/Contents/MacOS/Chromium"),
    Path("/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge"),
)


def _browser_executable(config: dict[str, Any]) -> Path:
    configured = str(config.get("browserExecutable") or os.environ.get("CHROME_BIN") or "").strip()
    if configured:
        path = Path(configured).expanduser()
        if path.is_file():
            return path
        resolved = shutil.which(configured)
        if resolved:
            return Path(resolved)
        raise RuntimeError(f"PDF 浏览器程序不存在：{configured}")

    for command in _BROWSER_COMMANDS:
        if resolved := shutil.which(command):
            return Path(resolved)
    for path in _BROWSER_PATHS:
        if path.is_file():
            return path
    raise RuntimeError(
        "生成 PDF 需要安装 Google Chrome、Chromium 或 Microsoft Edge；"
        "也可通过 report_rules.json 的 pdf.browserExecutable 指定程序路径。"
    )


def _extra_args(config: dict[str, Any]) -> list[str]:
    value = config.get("extraArgs") or []
    if not isinstance(value, Sequence) or isinstance(value, (str, bytes)):
        raise RuntimeError("pdf.extraArgs 必须是字符串数组。")
    return [str(item) for item in value if str(item).strip()]


def _is_complete_pdf(path: Path) -> bool:
    if not path.is_file() or path.stat().st_size < 8:
        return False
    with path.open("rb") as stream:
        if stream.read(5) != b"%PDF-":
            return False
        stream.seek(max(0, path.stat().st_size - 2048))
        return b"%%EOF" in stream.read()


def _stop_browser(process: subprocess.Popen[Any]) -> None:
    if process.poll() is not None:
        return
    process.terminate()
    try:
        process.wait(timeout=3)
    except subprocess.TimeoutExpired:
        process.kill()
        process.wait(timeout=3)


def render_html_to_pdf(
    html_path: Path,
    pdf_path: Path | None = None,
    *,
    config: dict[str, Any] | None = None,
) -> Path:
    config = config or {}
    html_path = html_path.resolve()
    if not html_path.is_file():
        raise RuntimeError(f"HTML 报告不存在：{html_path}")
    pdf_path = (pdf_path or html_path.with_suffix(".pdf")).resolve()
    pdf_path.parent.mkdir(parents=True, exist_ok=True)
    pdf_path.unlink(missing_ok=True)
    browser = _browser_executable(config)
    timeout = max(30, int(config.get("timeoutSeconds") or 180))

    with (
        tempfile.TemporaryDirectory(prefix="feedback-report-pdf-") as profile,
        tempfile.TemporaryFile(mode="w+", encoding="utf-8") as browser_log,
    ):
        command = [
            str(browser),
            "--headless=new",
            "--disable-gpu",
            "--disable-dev-shm-usage",
            "--disable-background-networking",
            "--disable-component-update",
            "--disable-default-apps",
            "--disable-extensions",
            "--disable-sync",
            "--metrics-recording-only",
            "--no-first-run",
            "--no-pdf-header-footer",
            "--run-all-compositor-stages-before-draw",
            "--window-size=1280,900",
            f"--user-data-dir={profile}",
            f"--print-to-pdf={pdf_path}",
            *_extra_args(config),
            html_path.as_uri(),
        ]
        process = subprocess.Popen(
            command,
            stdout=browser_log,
            stderr=subprocess.STDOUT,
            text=True,
        )
        deadline = time.monotonic() + timeout
        while process.poll() is None and time.monotonic() < deadline:
            if _is_complete_pdf(pdf_path):
                # Chrome occasionally finishes writing the PDF but keeps background
                # processes alive. The complete EOF marker means the file is safe.
                _stop_browser(process)
                break
            time.sleep(0.1)
        else:
            if process.poll() is None:
                _stop_browser(process)
                if not _is_complete_pdf(pdf_path):
                    browser_log.seek(0)
                    detail = browser_log.read().strip()
                    suffix = f"；浏览器日志：{detail[-1000:]}" if detail else ""
                    raise RuntimeError(
                        f"PDF 生成超时（{timeout} 秒）：{html_path.name}{suffix}"
                    )

        returncode = process.poll()
        if returncode not in (0, -15) and not _is_complete_pdf(pdf_path):
            browser_log.seek(0)
            detail = browser_log.read().strip()
            raise RuntimeError(f"PDF 生成失败（退出码 {returncode}）：{detail[-1000:]}")

    if not pdf_path.is_file() or pdf_path.stat().st_size < 8:
        raise RuntimeError(f"PDF 生成失败，未产生有效文件：{pdf_path}")
    if not _is_complete_pdf(pdf_path):
        raise RuntimeError(f"PDF 文件不完整：{pdf_path}")
    return pdf_path
