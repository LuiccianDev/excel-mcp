"""Regression tests for atomic release metadata synchronization."""

from __future__ import annotations

import importlib.util
import json
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]


def _load_bump_module():
    spec = importlib.util.spec_from_file_location(
        "bump_version", ROOT / "scripts" / "bump-version.py"
    )
    assert spec and spec.loader
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def test_bump_files_updates_manifest_version(tmp_path: Path) -> None:
    (tmp_path / "pyproject.toml").write_text('version = "1.0.0"\n', encoding="utf-8")
    (tmp_path / "CHANGELOG.md").write_text("## [Unreleased]\n", encoding="utf-8")
    (tmp_path / "manifest.json").write_text(
        json.dumps({"version": "1.0.0"}), encoding="utf-8"
    )

    _load_bump_module().bump_files("1.0.1", tmp_path)

    manifest = json.loads((tmp_path / "manifest.json").read_text(encoding="utf-8"))
    assert manifest["version"] == "1.0.1"
