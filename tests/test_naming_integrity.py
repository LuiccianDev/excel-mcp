"""Tests verifying naming disambiguation and packaging integrity contracts."""

from __future__ import annotations

import tomllib
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]


def test_package_and_scripts_naming_integrity() -> None:
    with (ROOT / "pyproject.toml").open("rb") as f:
        data = tomllib.load(f)

    project = data["project"]
    assert project["name"] == "luiccian-excel-mcp"

    scripts = project.get("scripts", {})
    assert "excel-mcp" in scripts
    assert "luiccian-excel-mcp" in scripts
    assert scripts["excel-mcp"] == "excel_mcp.__main__:main"
    assert scripts["luiccian-excel-mcp"] == "excel_mcp.__main__:main"

    entry_points = project.get("entry-points", {}).get("mcp.servers", {})
    assert "excel-mcp" in entry_points
    assert "luiccian-excel-mcp" in entry_points

    wheel_packages = data["tool"]["hatch"]["build"]["targets"]["wheel"]["packages"]
    assert wheel_packages == ["src/excel_mcp"]
