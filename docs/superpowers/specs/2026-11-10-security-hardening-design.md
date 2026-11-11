# Security Hardening and Tool Expansion — Design

## Intent

The published package is correct in its happy path (246 tests, mypy clean) but three defects break the documented setup: `MCP_ALLOWED_DIRECTORIES` is ignored by half the tools, the stdio transport writes non-JSON text to stdout, and `manifest.json` advertises a stale version and a tool that no longer exists. The repository also lost quality gates when pre-commit was kept as the only formatter entry point, and the tool surface stops short of publication-grade output (no validation rules, no conditional formatting, no print setup).

## Goals

- One resolver owns filesystem sandboxing; every tool honours `MCP_ALLOWED_DIRECTORIES`.
- `uvx luiccian-excel-mcp` speaks pure JSON-RPC on stdout.
- Version is reported from package metadata in the server handshake and in `manifest.json`.
- CI enforces lint, format, types and a coverage floor on Python 3.11, 3.12 and 3.13.
- Three new tools: `excel_add_data_validation`, `excel_add_conditional_format`, `excel_configure_page_setup`.

## Non-goals

- Rewriting the error contract across all 30 existing tools.
- Adding binary dependencies (images, PDF export).
- Changing the canonical `excel_*` naming or enabling legacy aliases by default.

## Anti-goals

- A coverage floor lowered to match the current number.
- A security fix that makes one tool stricter while another keeps the old resolver.

## Approaches considered

1. Patch `file_utils.py` to read `MCP_ALLOWED_DIRECTORIES` too. Smallest diff, but leaves two parsers that can drift again.
2. **Chosen:** extract `utils/paths.py` as the single resolver and make both security modules delegate to it.
3. Delete `file_utils.py` path checks entirely and force every tool through `security.py`. Cleanest, but it changes the decorator contract of nine tools in one step and risks regressions in error payloads.

## Decisions

- Precedence: `MCP_ALLOWED_DIRECTORIES` > `DIRECTORY` > configured default. The documented variable wins.
- Startup narration moves to the logger (stderr); the `list` CLI command keeps stdout because it is human output, not protocol output.
- `manifest.json` version is asserted against `pyproject.toml` by a test rather than generated, to keep the release script unchanged.
- New tools are openpyxl-native and validate their enums and colors before touching the workbook.
