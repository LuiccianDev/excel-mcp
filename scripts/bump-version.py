#!/usr/bin/env python3
"""
Automated Semantic Version Bumper for Excel MCP.

Supports automatic semantic bumping (patch, minor, major) or explicit version string.
Synchronizes version across pyproject.toml and CHANGELOG.md.

Usage:
    python3 scripts/bump-version.py --type patch
    python3 scripts/bump-version.py --type minor
    python3 scripts/bump-version.py --type major
    python3 scripts/bump-version.py --custom 2.1.0
    python3 scripts/bump-version.py 2.0.1
"""

from __future__ import annotations

import argparse
import json
import os
import re
from datetime import date
from pathlib import Path


_PROJECT_VERSION_RE = re.compile(
    r"(\[project\][^\[]*?^\s*version\s*=\s*\")([^\"]+)(\")",
    re.MULTILINE | re.DOTALL,
)


def get_current_version(pyproject_path: Path) -> str:
    if not pyproject_path.is_file():
        raise FileNotFoundError(f"pyproject.toml not found at {pyproject_path}")
    content = pyproject_path.read_text(encoding="utf-8")
    match = _PROJECT_VERSION_RE.search(content)
    if not match:
        match = re.search(r'^\s*version\s*=\s*"([^"]+)"', content, re.MULTILINE)
    if not match:
        raise ValueError("Could not find 'version = \"...\"' in pyproject.toml")
    return match.group(2 if match.re is _PROJECT_VERSION_RE else 1).strip()


def _parse_semver(version: str) -> tuple[int, int, int]:
    match = re.fullmatch(r"(\d+)\.(\d+)\.(\d+)", version)
    if not match:
        raise ValueError(
            f"Version '{version}' is not standard semver (major.minor.patch)"
        )
    return int(match.group(1)), int(match.group(2)), int(match.group(3))


def calculate_next_version(
    current: str, bump_type: str, custom_version: str | None = None
) -> str:
    current_tuple = _parse_semver(current)
    major, minor, patch = current_tuple

    if bump_type == "custom":
        if not custom_version:
            raise ValueError(
                "custom_version must be provided when bump_type is 'custom'"
            )
        normalized = custom_version.lstrip("v").strip()
        next_tuple = _parse_semver(normalized)
        if next_tuple <= current_tuple:
            raise ValueError(
                f"Custom version '{normalized}' must be strictly greater than current '{current}'"
            )
        return normalized

    if bump_type == "patch":
        patch += 1
    elif bump_type == "minor":
        minor += 1
        patch = 0
    elif bump_type == "major":
        major += 1
        minor = 0
        patch = 0
    else:
        raise ValueError(f"Unknown bump type: {bump_type}")

    return f"{major}.{minor}.{patch}"


def _ensure_unreleased_notes(changelog_path: Path) -> None:
    if not changelog_path.is_file():
        return
    lines = changelog_path.read_text(encoding="utf-8").splitlines()
    in_unreleased = False
    has_entries = False
    for line in lines:
        stripped = line.strip()
        if re.match(r"^##\s+\[Unreleased\]", stripped, re.IGNORECASE):
            in_unreleased = True
            continue
        if in_unreleased:
            if re.match(r"^##\s+\[", stripped):
                break
            if stripped and stripped != "---":
                has_entries = True
                break
    if not has_entries:
        raise ValueError(
            "CHANGELOG.md [Unreleased] section is empty; add release notes before bumping."
        )


def _update_pyproject(pyproject_path: Path, new_version: str, dry_run: bool) -> None:
    if not pyproject_path.is_file():
        return
    content = pyproject_path.read_text(encoding="utf-8")
    if _PROJECT_VERSION_RE.search(content):
        updated = _PROJECT_VERSION_RE.sub(rf"\g<1>{new_version}\3", content, count=1)
    else:
        updated = re.sub(
            r'version\s*=\s*"[^"]+"',
            f'version = "{new_version}"',
            content,
            count=1,
        )
    if not dry_run:
        pyproject_path.write_text(updated, encoding="utf-8")
    print(f"  [OK] {pyproject_path.name} -> {new_version}")


def _update_changelog(changelog_path: Path, new_version: str, dry_run: bool) -> None:
    if not changelog_path.is_file():
        return
    today_str = date.today().isoformat()
    content = changelog_path.read_text(encoding="utf-8")
    target_header = f"## [{new_version}] - {today_str}"
    if f"[{new_version}]" in content:
        return
    updated = content.replace(
        "## [Unreleased]",
        f"## [Unreleased]\n\n---\n\n{target_header}",
        1,
    )
    if not dry_run:
        changelog_path.write_text(updated, encoding="utf-8")
    print(f"  [OK] {changelog_path.name} promoted [Unreleased] -> {new_version}")


def _update_manifest(manifest_path: Path, new_version: str, dry_run: bool) -> None:
    if not manifest_path.is_file() or dry_run:
        return
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    manifest["version"] = new_version
    manifest_path.write_text(json.dumps(manifest, indent=2) + "\n", encoding="utf-8")
    print(f"  [OK] {manifest_path.name} -> {new_version}")


def _export_github_output(new_version: str) -> None:
    github_output = os.environ.get("GITHUB_OUTPUT")
    if not github_output or not Path(github_output).is_file():
        return
    with open(github_output, "a", encoding="utf-8") as f:
        f.write(f"version={new_version}\n")
        f.write(f"tag=v{new_version}\n")
    print(
        f"  [OK] Exported version={new_version} and tag=v{new_version} to $GITHUB_OUTPUT"
    )


def bump_files(new_version: str, repo_root: Path, dry_run: bool = False) -> None:
    print(f">>> Bumping to version: {new_version} (dry_run={dry_run})")
    _update_pyproject(repo_root / "pyproject.toml", new_version, dry_run)
    _update_changelog(repo_root / "CHANGELOG.md", new_version, dry_run)
    _update_manifest(repo_root / "manifest.json", new_version, dry_run)
    _export_github_output(new_version)


def main() -> None:
    parser = argparse.ArgumentParser(description="Bump project version")
    parser.add_argument(
        "version", nargs="?", default=None, help="Explicit new version (e.g. 2.0.1)"
    )
    parser.add_argument(
        "--type",
        choices=["patch", "minor", "major", "custom"],
        default=None,
        help="Semantic version bump type (patch, minor, major, custom)",
    )
    parser.add_argument(
        "--custom", default=None, help="Custom version string when --type is custom"
    )
    parser.add_argument(
        "--dry-run", action="store_true", help="Calculate without modifying files"
    )

    args = parser.parse_args()
    repo_root = Path(__file__).resolve().parent.parent
    pyproject_path = repo_root / "pyproject.toml"

    current_ver = get_current_version(pyproject_path)
    print(f"Current version in pyproject.toml: {current_ver}")

    _ensure_unreleased_notes(repo_root / "CHANGELOG.md")

    if args.version:
        new_version = calculate_next_version(current_ver, "custom", args.version)
    elif args.type:
        new_version = calculate_next_version(current_ver, args.type, args.custom)
    else:
        new_version = calculate_next_version(current_ver, "patch")

    bump_files(new_version, repo_root, dry_run=args.dry_run)
    print(f"\nSuccessfully resolved next version: {new_version} (Tag: v{new_version})")


if __name__ == "__main__":
    main()
