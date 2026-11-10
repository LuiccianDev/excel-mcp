"""Drift guard tests ensuring skills and documentation stay strictly synchronized with registered tools."""

from __future__ import annotations

import os
import re
import unittest
from pathlib import Path
from unittest.mock import patch

from excel_mcp.tools.register_tools import register_all_tools


class MockMCP:
    def __init__(self) -> None:
        self.tools: set[str] = set()

    def tool(self, name: str | None = None, annotations: object = None) -> object:
        def decorator(fn: object) -> object:
            tool_name = name or getattr(fn, "__name__", "unnamed")
            self.tools.add(tool_name)
            return fn

        return decorator


class TestSkillsAndDocsSynchronization(unittest.TestCase):
    """Test suite ensuring documentation and skills never drift from actual registered tools."""

    @classmethod
    def setUpClass(cls) -> None:
        cls.repo_root = Path(__file__).resolve().parent.parent

        # Capture canonical registered tools
        with patch.dict(os.environ, {"EXCEL_MCP_LEGACY_ALIASES": "0"}):
            mcp = MockMCP()
            register_all_tools(mcp)
            cls.registered_tools = mcp.tools

    def test_canonical_tool_count(self) -> None:
        """Verify exactly 31 canonical tools are registered."""
        self.assertEqual(len(self.registered_tools), 31)

    def test_all_tools_documented_in_tools_md(self) -> None:
        """Verify every registered tool is documented in docs/TOOLS.md."""
        tools_doc = self.repo_root / "docs" / "TOOLS.md"
        self.assertTrue(tools_doc.exists(), "docs/TOOLS.md must exist")
        content = tools_doc.read_text(encoding="utf-8")

        for tool_name in self.registered_tools:
            self.assertIn(
                f"`{tool_name}`",
                content,
                f"Canonical tool '{tool_name}' is missing from docs/TOOLS.md",
            )

    def test_all_skill_tool_references_are_valid(self) -> None:
        """Verify all excel_* tools mentioned in skills/**/SKILL.md exist in registered tools."""
        skills_dir = self.repo_root / "skills"
        skill_files = list(skills_dir.glob("*/SKILL.md"))
        self.assertGreaterEqual(len(skill_files), 4, "Must find at least 4 skills")

        excel_tool_pattern = re.compile(r"\b(excel_[a-z0-9_]+)\b")

        for skill_file in skill_files:
            text = skill_file.read_text(encoding="utf-8")
            matches = set(excel_tool_pattern.findall(text))
            for ref in matches:
                self.assertIn(
                    ref,
                    self.registered_tools,
                    f"Skill '{skill_file.parent.name}' references unregistered tool '{ref}'",
                )


if __name__ == "__main__":
    unittest.main()
