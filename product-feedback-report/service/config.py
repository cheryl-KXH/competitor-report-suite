from __future__ import annotations

import json
from pathlib import Path
from typing import Any


ROOT = Path(__file__).resolve().parents[1]
CONFIG_DIR = ROOT / "config"


def load_json(path: Path) -> dict[str, Any]:
    with path.open("r", encoding="utf-8") as handle:
        return json.load(handle)


def load_optional_json(name: str, example_name: str | None = None) -> dict[str, Any]:
    path = CONFIG_DIR / name
    if path.exists():
        return load_json(path)
    if example_name:
        example = CONFIG_DIR / example_name
        if example.exists():
            return load_json(example)
    return {}


def load_configs() -> dict[str, dict[str, Any]]:
    return {
        "dingtalk": load_optional_json("dingtalk.json", "dingtalk.example.json"),
        "dingtalk_docs": load_optional_json("dingtalk_docs.json", "dingtalk_docs.example.json"),
        "field_mapping": load_optional_json("field_mapping.json"),
        "platform_rules": load_optional_json("platform_rules.json"),
        "report_rules": load_optional_json("report_rules.json"),
    }


def output_root(configs: dict[str, dict[str, Any]]) -> Path:
    value = configs.get("report_rules", {}).get("outputDirectory") or "outputs"
    path = Path(str(value))
    if not path.is_absolute():
        path = ROOT / path
    return path

