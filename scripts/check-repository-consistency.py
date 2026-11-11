#!/usr/bin/env python3
"""Fail when package, manifest, registered tools, or generated docs drift."""

from __future__ import annotations

import json
import sys
import tomllib
from pathlib import Path
from typing import Any


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from excel_mcp.tools.register_tools import register_all_tools  # noqa: E402


class Collector:
    def __init__(self) -> None:
        self.names: set[str] = set()

    def tool(self, *, name: str, annotations: Any) -> Any:
        def decorator(fn: Any) -> Any:
            self.names.add(name)
            return fn

        return decorator


def main() -> int:
    with (ROOT / "pyproject.toml").open("rb") as handle:
        project = tomllib.load(handle)["project"]
    manifest = json.loads((ROOT / "manifest.json").read_text(encoding="utf-8"))
    docs = (ROOT / "docs" / "TOOLS.md").read_text(encoding="utf-8")

    collector = Collector()
    register_all_tools(collector)
    registered = collector.names
    manifest_names = {item["name"] for item in manifest["tools"]}
    missing_docs = sorted(name for name in registered if f"`{name}`" not in docs)
    errors: list[str] = []
    if manifest["version"] != project["version"]:
        errors.append("manifest version does not match pyproject version")
    if not manifest_names <= registered:
        errors.append(
            f"manifest contains unregistered tools: {sorted(manifest_names - registered)}"
        )
    if missing_docs:
        errors.append(f"docs/TOOLS.md is missing tools: {missing_docs}")
    if errors:
        for error in errors:
            print(f"ERROR: {error}", file=sys.stderr)
        return 1
    print(
        f"Repository consistency passed for {len(registered)} tools at version {project['version']}."
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
