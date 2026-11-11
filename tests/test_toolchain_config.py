"""Contracts for the repository's uv-first development toolchain."""

from __future__ import annotations

import tomllib
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]


def _pyproject() -> dict:
    with (ROOT / "pyproject.toml").open("rb") as handle:
        return tomllib.load(handle)


def test_project_supports_python_311_or_newer() -> None:
    assert _pyproject()["project"]["requires-python"] == ">=3.11"


def test_development_group_contains_required_uv_tools() -> None:
    dependencies = set(_pyproject()["dependency-groups"]["dev"])
    for package in ("pytest", "pytest-cov", "ruff", "mypy", "build"):
        assert any(item.startswith(f"{package}") for item in dependencies)


def test_test_and_quality_configuration_is_centralized() -> None:
    config = _pyproject()
    assert config["tool"]["pytest"]["ini_options"]["testpaths"] == ["tests"]
    assert config["tool"]["coverage"]["report"]["fail_under"] >= 80
    assert "ruff" in config["tool"]
    assert "mypy" in config["tool"]


def test_readme_documents_only_uv_run_commands() -> None:
    readme = (ROOT / "README.md").read_text(encoding="utf-8")
    for command in (
        "uv run pytest",
        "uv run ruff check .",
        "uv run ruff format --check .",
        "uv run mypy src",
        "uv run python -m build",
    ):
        assert command in readme
