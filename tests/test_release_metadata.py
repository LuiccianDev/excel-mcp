"""Regression tests for atomic release metadata synchronization."""

from __future__ import annotations

import importlib.util
import json
import re
import subprocess
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


def _load_script(filename: str):
    name = filename.replace("-", "_").removesuffix(".py")
    spec = importlib.util.spec_from_file_location(name, ROOT / "scripts" / filename)
    assert spec and spec.loader
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def _project_version() -> str:
    match = re.search(
        r'^version = "([^"]+)"',
        (ROOT / "pyproject.toml").read_text(encoding="utf-8"),
        flags=re.MULTILINE,
    )
    assert match
    return match.group(1)


def test_consistency_check_rejects_tag_that_differs_from_version(capsys) -> None:
    module = _load_script("check-repository-consistency.py")

    assert module.main(["--tag", "v0.0.1"]) == 1
    assert "Tag v0.0.1 does not match pyproject version" in capsys.readouterr().err


def test_consistency_check_accepts_matching_tag() -> None:
    module = _load_script("check-repository-consistency.py")

    assert module.main(["--tag", f"v{_project_version()}"]) == 0


def _git(root: Path, *args: str) -> str:
    result = subprocess.run(
        ["git", "-c", "user.name=t", "-c", "user.email=t@t", *args],
        cwd=root,
        check=True,
        capture_output=True,
        text=True,
    )
    return result.stdout.strip()


def _repo_with_version(root: Path, version: str) -> str:
    _git(root, "init", "-q")
    (root / "pyproject.toml").write_text(f'version = "{version}"\n', encoding="utf-8")
    _git(root, "add", ".")
    _git(root, "commit", "-q", "-m", "init")
    return _git(root, "rev-parse", "HEAD")


def _bump(root: Path, version: str) -> None:
    (root / "pyproject.toml").write_text(f'version = "{version}"\n', encoding="utf-8")
    _git(root, "commit", "-q", "-am", "bump")


def test_release_state_false_when_version_unchanged(tmp_path: Path) -> None:
    previous = _repo_with_version(tmp_path, "1.0.0")
    (tmp_path / "README.md").write_text("docs\n", encoding="utf-8")
    _git(tmp_path, "add", ".")
    _git(tmp_path, "commit", "-q", "-m", "docs")

    state = _load_script("release-state.py").release_state(tmp_path, previous)

    assert state == {"version": "1.0.0", "tag": "v1.0.0", "should_release": "false"}


def test_release_state_true_when_version_changed_and_tag_missing(
    tmp_path: Path,
) -> None:
    previous = _repo_with_version(tmp_path, "1.0.0")
    _bump(tmp_path, "1.0.1")

    state = _load_script("release-state.py").release_state(tmp_path, previous)

    assert state == {"version": "1.0.1", "tag": "v1.0.1", "should_release": "true"}


def test_release_state_false_when_tag_already_exists(tmp_path: Path) -> None:
    previous = _repo_with_version(tmp_path, "1.0.0")
    _bump(tmp_path, "1.0.1")
    _git(tmp_path, "tag", "v1.0.1")

    state = _load_script("release-state.py").release_state(tmp_path, previous)

    assert state["should_release"] == "false"


def test_release_state_without_previous_ref_only_checks_tag(tmp_path: Path) -> None:
    _repo_with_version(tmp_path, "1.0.0")

    module = _load_script("release-state.py")

    assert module.release_state(tmp_path, None)["should_release"] == "true"
    _git(tmp_path, "tag", "v1.0.0")
    assert module.release_state(tmp_path, None)["should_release"] == "false"
