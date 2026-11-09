#!/usr/bin/env python3
"""
Extract release notes for a specified version/tag from CHANGELOG.md.

Usage:
    python3 scripts/extract-release-notes.py --version 2.0.0 --output release_body.md
    python3 scripts/extract-release-notes.py --tag v2.0.0 --output release_body.md
"""

from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path


def extract_notes(changelog_path: Path, target_version: str | None) -> str:
    """Extract release notes section corresponding to the target version."""
    if not changelog_path.is_file():
        raise FileNotFoundError(f"Changelog not found at {changelog_path}")

    content = changelog_path.read_text(encoding="utf-8")
    lines = content.splitlines()

    # Normalize version: remove leading 'v'
    ver = target_version.lstrip("v").strip() if target_version else None

    # Find section header like ## [2.0.0] or ## [v2.0.0] or ## [Unreleased]
    header_pattern = re.compile(r"^##\s+\[(.*?)\](?:\s+-\s+.*)?$")

    in_target_section = False
    extracted_lines: list[str] = []

    for line in lines:
        match = header_pattern.match(line.strip())
        if match:
            current_ver = match.group(1).lstrip("v").strip()
            if in_target_section:
                # Reached next section header, stop
                break
            if ver is None or current_ver.lower() == ver.lower():
                in_target_section = True
                continue
        elif in_target_section:
            extracted_lines.append(line)

    result = "\n".join(extracted_lines).strip()
    if not result:
        # Fallback: return a clean default body
        fallback_ver = ver or "Latest"
        return f"# Release {fallback_ver}\n\nAutomated release for Excel MCP Server."

    return result


def main() -> None:
    parser = argparse.ArgumentParser(description="Extract release notes from CHANGELOG.md")
    parser.add_argument("--version", help="Version to extract (e.g. 2.0.0)")
    parser.add_argument("--tag", help="Tag to extract (e.g. v2.0.0)")
    parser.add_argument(
        "--changelog",
        default="CHANGELOG.md",
        help="Path to CHANGELOG.md (default: CHANGELOG.md)",
    )
    parser.add_argument(
        "--output",
        default="release_body.md",
        help="Path to output markdown file (default: release_body.md)",
    )

    args = parser.parse_args()
    target = args.version or args.tag
    changelog = Path(args.changelog)
    out_path = Path(args.output)

    try:
        notes = extract_notes(changelog, target)
        out_path.write_text(notes + "\n", encoding="utf-8")
        print(f"Successfully extracted release notes to {out_path}")
    except Exception as exc:
        print(f"Error extracting release notes: {exc}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
