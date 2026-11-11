"""Repository-wide contracts for the public MCP tool inventory."""

from __future__ import annotations

import json
import tomllib
from pathlib import Path
from typing import Any

from excel_mcp.tools.register_tools import register_all_tools
from excel_mcp.tools.validation_tools import add_data_validation


ROOT = Path(__file__).resolve().parents[1]


class Collector:
    def __init__(self) -> None:
        self.names: set[str] = set()

    def tool(self, *, name: str, annotations: Any) -> Any:
        def decorator(fn: Any) -> Any:
            self.names.add(name)
            return fn

        return decorator


def registered_names() -> set[str]:
    collector = Collector()
    register_all_tools(collector)
    return collector.names


def test_registry_has_current_canonical_tool_count() -> None:
    assert len(registered_names()) == 33


def test_every_registered_tool_is_documented() -> None:
    docs = (ROOT / "docs" / "TOOLS.md").read_text(encoding="utf-8")
    missing = [name for name in registered_names() if f"`{name}`" not in docs]
    assert missing == []


def test_manifest_version_matches_package_metadata() -> None:
    manifest = json.loads((ROOT / "manifest.json").read_text(encoding="utf-8"))
    with (ROOT / "pyproject.toml").open("rb") as handle:
        project = tomllib.load(handle)["project"]
    assert manifest["version"] == project["version"]


def test_malformed_tool_input_returns_structured_error(tmp_path, monkeypatch) -> None:
    monkeypatch.setenv("MCP_ALLOWED_DIRECTORIES", str(tmp_path))
    result = add_data_validation(
        filename=str(tmp_path / "missing.xlsx"),
        cell_range=None,
        formula1='"Open,Closed"',
    )
    assert result["status"] == "error"


def test_filesystem_fixture_is_isolated(tmp_path: Path) -> None:
    assert tmp_path.exists()
    assert list(tmp_path.iterdir()) == []
