#!/usr/bin/env python3
"""Decide whether the current commit must produce a release.

Prints ``version``, ``tag`` and ``should_release`` as ``key=value`` lines so the
output can be appended to ``$GITHUB_OUTPUT``. A release is due only when the
``pyproject.toml`` version changed since ``--previous-ref`` (when given) and the
matching ``vX.Y.Z`` tag does not exist yet.
"""

from __future__ import annotations

import argparse
import re
import subprocess
from pathlib import Path


ZERO_SHA = "0" * 40
VERSION_RE = re.compile(r'^version = "([^"]+)"', re.MULTILINE)


def _git(root: Path, *args: str) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        ["git", *args], cwd=root, capture_output=True, text=True, check=False
    )


def _version(text: str) -> str | None:
    match = VERSION_RE.search(text)
    return match.group(1) if match else None


def release_state(root: Path, previous_ref: str | None) -> dict[str, str]:
    """Return ``version``, ``tag`` and ``should_release`` for ``root``."""
    version = _version((root / "pyproject.toml").read_text(encoding="utf-8"))
    if version is None:
        raise SystemExit("version not found in pyproject.toml")
    tag = f"v{version}"

    changed = True
    if previous_ref and previous_ref != ZERO_SHA:
        shown = _git(root, "show", f"{previous_ref}:pyproject.toml")
        if shown.returncode == 0:
            changed = _version(shown.stdout) != version

    tag_exists = bool(_git(root, "tag", "-l", tag).stdout.strip())
    should_release = changed and not tag_exists
    return {
        "version": version,
        "tag": tag,
        "should_release": "true" if should_release else "false",
    }


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--previous-ref", default=None)
    args = parser.parse_args(argv)
    for key, value in release_state(Path.cwd(), args.previous_ref).items():
        print(f"{key}={value}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
