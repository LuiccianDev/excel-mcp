"""Contract tests for the MCPB manifest and canonical tool registry."""

from __future__ import annotations

import json
import tomllib
from pathlib import Path
from typing import Any

from excel_mcp.tools.register_tools import register_all_tools


ROOT = Path(__file__).resolve().parents[1]


class _Collector:
    def __init__(self) -> None:
        self.tools: dict[str, Any] = {}

    def tool(self, *, name: str, annotations: Any) -> Any:
        def decorator(fn: Any) -> Any:
            self.tools[name] = fn
            return fn

        return decorator


def _manifest() -> dict[str, Any]:
    return json.loads((ROOT / "manifest.json").read_text(encoding="utf-8"))


def _package_version() -> str:
    with (ROOT / "pyproject.toml").open("rb") as handle:
        return tomllib.load(handle)["project"]["version"]


def _canonical_tools() -> set[str]:
    collector = _Collector()
    register_all_tools(collector)
    return set(collector.tools)


def test_manifest_version_matches_package_metadata() -> None:
    assert _manifest()["version"] == _package_version()


def test_manifest_tools_are_registered_canonical_tools() -> None:
    canonical = _canonical_tools()
    manifest_names = {item["name"] for item in _manifest()["tools"]}

    assert len(canonical) == 33
    assert manifest_names <= canonical
    assert "excel_configure_page_setup" in manifest_names


def test_manifest_directory_configuration_supports_multiple_roots() -> None:
    assert _manifest()["user_config"]["directory"]["multiple"] is True
