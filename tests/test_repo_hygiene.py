"""Repository hygiene contracts for contributors and documentation."""

from __future__ import annotations

from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]


def test_contributor_and_github_templates_exist() -> None:
    required = [
        ROOT / "CONTRIBUTING.md",
        ROOT / ".github" / "pull_request_template.md",
        ROOT / ".github" / "ISSUE_TEMPLATE" / "bug_report.yml",
        ROOT / ".github" / "ISSUE_TEMPLATE" / "feature_request.yml",
        ROOT / ".github" / "ISSUE_TEMPLATE" / "config.yml",
    ]
    assert all(path.exists() for path in required)


def test_documentation_uses_supported_uv_workflow() -> None:
    readme = (ROOT / "README.md").read_text(encoding="utf-8")
    security = (ROOT / "SECURITY.md").read_text(encoding="utf-8")
    assert "uvx luiccian-excel-mcp" in readme
    assert "uv sync --all-groups" in readme
    assert "MCP_ALLOWED_DIRECTORIES" in security
    assert "2.1.x" in security


def test_documentation_contains_no_retired_tool_names() -> None:
    documents = [
        ROOT / "README.md",
        ROOT / "SECURITY.md",
        ROOT / "docs" / "TOOLS.md",
        *ROOT.glob("skills/*/SKILL.md"),
    ]
    combined = "\n".join(path.read_text(encoding="utf-8") for path in documents)
    for retired in ("build_sheet", "search_text", "primary_filename"):
        assert retired not in combined


def test_security_policy_requests_private_reporting_and_reproduction() -> None:
    security = (ROOT / "SECURITY.md").read_text(encoding="utf-8")
    assert "Do not create public GitHub issues" in security
    assert "Steps or proof-of-concept" in security
    assert "environment details" in security
