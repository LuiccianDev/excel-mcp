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

        # Capture canonical registered tools (full mode) and compose-mode tools
        with patch.dict(
            os.environ, {"EXCEL_MCP_LEGACY_ALIASES": "0", "EXCEL_MCP_MODE": "full"}
        ):
            mcp = MockMCP()
            register_all_tools(mcp)
            cls.registered_tools = mcp.tools

        with patch.dict(
            os.environ, {"EXCEL_MCP_LEGACY_ALIASES": "0", "EXCEL_MCP_MODE": "compose"}
        ):
            compose_mcp = MockMCP()
            register_all_tools(compose_mcp)
            cls.compose_tools = compose_mcp.tools

    def test_canonical_tool_count(self) -> None:
        """Verify exactly 33 canonical tools are registered in full mode."""
        self.assertEqual(len(self.registered_tools), 33)

    def test_compose_tool_count(self) -> None:
        """Verify exactly 8 tools are registered in default compose mode."""
        self.assertEqual(len(self.compose_tools), 8)

    def test_readme_documents_v220_tool_modes(self) -> None:
        """Verify README.md documents compose default mode and full opt-in."""
        readme = (self.repo_root / "README.md").read_text(encoding="utf-8")
        for token in (
            "compose",
            "EXCEL_MCP_MODE=full",
            "--tool-mode full",
            "excel_workbook",
            "excel_data",
            "excel_format",
            "excel_formula",
            "excel_analytics",
        ):
            self.assertIn(token, readme, f"README.md missing '{token}' for v2.2.0")

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

        valid_tools = self.registered_tools | self.compose_tools
        for skill_file in skill_files:
            text = skill_file.read_text(encoding="utf-8")
            matches = set(excel_tool_pattern.findall(text))
            for ref in matches:
                self.assertIn(
                    ref,
                    valid_tools,
                    f"Skill '{skill_file.parent.name}' references unregistered tool '{ref}'",
                )


if __name__ == "__main__":
    unittest.main()
