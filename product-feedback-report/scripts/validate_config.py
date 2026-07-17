#!/usr/bin/env python3
from __future__ import annotations

import os
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from service.config import load_configs


def main() -> int:
    configs = load_configs()
    warnings: list[str] = []
    dingtalk = configs.get("dingtalk", {})
    if not dingtalk.get("baseId"):
        warnings.append("config/dingtalk.json 缺少 baseId。")
    if not dingtalk.get("tableId"):
        warnings.append("config/dingtalk.json 缺少 tableId。")
    docs = configs.get("dingtalk_docs", {})
    if not docs.get("streamableHttpUrl") and not os.getenv("DINGTALK_DOCS_MCP_URL"):
        server_name = docs.get("serverName") or "dingtalk-docs"
        warnings.append(
            f"钉钉文档 MCP 未配置 URL；请设置 DINGTALK_DOCS_MCP_URL，"
            f"或确认 mcporter 已注册 {server_name}。"
        )
    fields = configs.get("field_mapping", {}).get("fields", {})
    for key, meta in fields.items():
        if meta.get("required") and not meta.get("fieldId"):
            warnings.append(f"必填字段未配置 fieldId：{key} / {meta.get('label')}")
    if warnings:
        print("配置提醒：")
        for warning in warnings:
            print(f"- {warning}")
    else:
        print("配置检查通过。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
