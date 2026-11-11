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


def get_current_version(pyproject_path: Path) -> str:
    if not pyproject_path.is_file():
        raise FileNotFoundError(f"pyproject.toml not found at {pyproject_path}")
    content = pyproject_path.read_text(encoding="utf-8")
    match = re.search(r'version\s*=\s*"([^"]+)"', content)
    if not match:
        raise ValueError("Could not find 'version = \"...\"' in pyproject.toml")
    return match.group(1).strip()


def calculate_next_version(
    current: str, bump_type: str, custom_version: str | None = None
) -> str:
    if bump_type == "custom":
        if not custom_version:
            raise ValueError(
                "custom_version must be provided when bump_type is 'custom'"
            )
        return custom_version.lstrip("v").strip()

    # Parse semver (e.g., 2.0.0)
    match = re.match(r"^(\d+)\.(\d+)\.(\d+)(.*)$", current)
    if not match:
        raise ValueError(
            f"Current version '{current}' is not standard semver (major.minor.patch)"
        )

    major = int(match.group(1))
    minor = int(match.group(2))
    patch = int(match.group(3))

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


def bump_files(new_version: str, repo_root: Path, dry_run: bool = False) -> None:
    pyproject_path = repo_root / "pyproject.toml"
    changelog_path = repo_root / "CHANGELOG.md"
    manifest_path = repo_root / "manifest.json"

    print(f">>> Bumping to version: {new_version} (dry_run={dry_run})")

    # 1. Update pyproject.toml
    if pyproject_path.is_file():
        content = pyproject_path.read_text(encoding="utf-8")
        updated = re.sub(
            r'version\s*=\s*"[^"]+"',
            f'version = "{new_version}"',
            content,
            count=1,
        )
        if not dry_run:
            pyproject_path.write_text(updated, encoding="utf-8")
        print(f"  [OK] {pyproject_path.name} -> {new_version}")

    # 2. Update CHANGELOG.md if [Unreleased] is present
    today_str = date.today().isoformat()
    if changelog_path.is_file():
        content = changelog_path.read_text(encoding="utf-8")
        target_header = f"## [{new_version}] - {today_str}"
        if f"[{new_version}]" not in content:
            updated = content.replace(
                "## [Unreleased]",
                f"## [Unreleased]\n\n---\n\n{target_header}",
                1,
            )
            if not dry_run:
                changelog_path.write_text(updated, encoding="utf-8")
            print(
                f"  [OK] {changelog_path.name} promoted [Unreleased] -> {new_version}"
            )

    # 3. Keep the distributable manifest version aligned with the package.
    if manifest_path.is_file() and not dry_run:
        manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
        manifest["version"] = new_version
        manifest_path.write_text(
            json.dumps(manifest, indent=2) + "\n", encoding="utf-8"
        )
        print(f"  [OK] {manifest_path.name} -> {new_version}")

    # 4. Export to GITHUB_OUTPUT if running in GitHub Actions
    github_output = os.environ.get("GITHUB_OUTPUT")
    if github_output and Path(github_output).is_file():
        with open(github_output, "a", encoding="utf-8") as f:
            f.write(f"version={new_version}\n")
            f.write(f"tag=v{new_version}\n")
        print(
            f"  [OK] Exported version={new_version} and tag=v{new_version} to $GITHUB_OUTPUT"
        )


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

    if args.version:
        new_version = args.version.lstrip("v").strip()
    elif args.type:
        new_version = calculate_next_version(current_ver, args.type, args.custom)
    else:
        new_version = calculate_next_version(current_ver, "patch")

    bump_files(new_version, repo_root, dry_run=args.dry_run)
    print(f"\nSuccessfully resolved next version: {new_version} (Tag: v{new_version})")


if __name__ == "__main__":
    main()
