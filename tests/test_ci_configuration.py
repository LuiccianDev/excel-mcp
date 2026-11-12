"""Contracts for local and GitHub quality gates."""

from __future__ import annotations

from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
WORKFLOWS = ROOT / ".github" / "workflows"


def _read(name: str) -> str:
    return (WORKFLOWS / name).read_text(encoding="utf-8")


def test_ci_uses_uv_quality_gates_and_supported_python_matrix() -> None:
    ci = _read("ci.yml")
    assert "Verify PR does not bump version directly" in ci
    assert "uv lock --check" in ci
    assert "uv sync --all-groups" in ci
    assert "uv run ruff check" in ci
    assert "uv run ruff format --check" in ci
    assert "uv run mypy src" in ci
    assert "uv run pytest --cov" in ci
    for version in ("3.11", "3.12", "3.13", "3.14"):
        assert version in ci
    assert "cancel-in-progress: true" in ci
    assert "paths-ignore:" in ci


def test_codeql_and_release_workflows_use_safe_triggers_and_uv() -> None:
    codeql = _read("codeql.yml")
    release = _read("bump-and-release.yml")
    assert "cancel-in-progress: true" in codeql
    assert "paths-ignore:" in codeql
    assert "uv sync --all-groups" in release
    assert "uv run python scripts/bump-version.py" in release
    assert "uv run python scripts/extract-release-notes.py" in release
    assert "cancel-in-progress: true" in release


def test_uv_is_the_only_local_quality_gate() -> None:
    assert not (ROOT / ".pre-commit-config.yaml").exists()
    pyproject = (ROOT / "pyproject.toml").read_text(encoding="utf-8")
    assert "pre-commit" not in pyproject
    assert "uv run ruff check ." in (ROOT / "README.md").read_text(encoding="utf-8")
