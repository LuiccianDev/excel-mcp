"""Tests for mode-aware server instructions and single-source COMPOUND_REGISTRY metadata."""

from __future__ import annotations

from excel_mcp.server import create_server
from excel_mcp.tools.compound_tools import COMPOUND_REGISTRY


def test_create_server_compound_instructions_describe_dispatchers() -> None:
    mcp = create_server(mode="compound")
    instructions = mcp.instructions or ""
    assert "compound" in instructions.lower()
    for group, actions in COMPOUND_REGISTRY.items():
        assert f"excel_{group}" in instructions
        for action in actions:
            assert f"`{action}`" in instructions


def test_create_server_full_instructions_describe_full_mode() -> None:
    mcp = create_server(mode="full")
    instructions = mcp.instructions or ""
    assert "full" in instructions.lower()
    assert "33" in instructions
    for group in COMPOUND_REGISTRY.values():
        for spec in group.values():
            assert f"`{spec.full_name}`" in instructions


def test_compound_registry_entries_declare_unique_full_names() -> None:
    full_names = [
        spec.full_name
        for group in COMPOUND_REGISTRY.values()
        for spec in group.values()
    ]
    assert len(full_names) == 30
    assert all(name.startswith("excel_") for name in full_names)
    assert len(set(full_names)) == 30
