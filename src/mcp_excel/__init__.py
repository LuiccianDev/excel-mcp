"""Backward-compatibility module alias redirecting `mcp_excel` to `excel_mcp`."""

from __future__ import annotations

import importlib
import importlib.abc
import importlib.machinery
import sys
from typing import Any

import excel_mcp

for _attr in dir(excel_mcp):
    if not _attr.startswith("__"):
        globals()[_attr] = getattr(excel_mcp, _attr)


class _AliasLoader(importlib.abc.Loader):
    def __init__(self, target_mod: Any) -> None:
        self.target_mod = target_mod

    def create_module(self, spec: Any) -> Any:
        return self.target_mod

    def exec_module(self, module: Any) -> None:
        pass


class _AliasFinder(importlib.abc.MetaPathFinder):
    def find_spec(self, fullname: str, path: Any = None, target: Any = None) -> Any:
        if fullname.startswith("mcp_excel."):
            canonical_name = "excel_mcp." + fullname[len("mcp_excel.") :]
            try:
                canonical_mod = importlib.import_module(canonical_name)
                loader = _AliasLoader(canonical_mod)
                spec = importlib.machinery.ModuleSpec(fullname, loader)
                if hasattr(canonical_mod, "__path__"):
                    spec.submodule_search_locations = list(canonical_mod.__path__)
                return spec
            except ImportError:
                return None
        return None


if not any(isinstance(f, _AliasFinder) for f in sys.meta_path):
    sys.meta_path.insert(0, _AliasFinder())


def __getattr__(name: str) -> Any:
    try:
        mod = importlib.import_module(f"excel_mcp.{name}")
        sys.modules[f"mcp_excel.{name}"] = mod
        return mod
    except ModuleNotFoundError as e:
        if e.name != f"excel_mcp.{name}":
            raise
    if hasattr(excel_mcp, name):
        return getattr(excel_mcp, name)
    raise AttributeError(f"module 'mcp_excel' has no attribute '{name}'")
