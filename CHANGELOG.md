# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

---

## [Unreleased]

---

## [2.1.4] - 2026-11-11

### Added
- **Unified Path Sandbox (`excel_mcp.utils.paths`)**: Single resolver shared by every tool. `MCP_ALLOWED_DIRECTORIES` now governs all read, write and discovery operations and accepts multiple comma-separated roots.
- **Coverage Gate**: `pytest-cov` enforces a 75% floor over `src/excel_mcp` in CI and locally.
- **Validation and Print Layout Tools**: Added `excel_add_data_validation`, `excel_protect_sheet`, and `excel_configure_page_setup` for publication-ready workbook templates.

### Changed
- **CI Matrix**: GitHub Actions now tests Python 3.11, 3.12 and 3.13, matching the declared classifiers.
- **Quality Gates**: Pre-commit hooks removed; CI runs `ruff check`, `ruff format --check`, `mypy` and `pytest --cov` through `uv` on the whole repository, including `scripts/` and `tests/`.

### Fixed
- **Documented Sandbox Variable**: Tools decorated with `validate_file_access` rejected paths that were valid under `MCP_ALLOWED_DIRECTORIES`, and `excel_list_documents` silently scanned `./documents`. Both now use the shared resolver.
- **stdio Transport Purity**: Startup messages no longer write to stdout, which corrupted the JSON-RPC channel for strict MCP clients.
- **Server Version**: `serverInfo.version` reports the package version instead of the FastMCP version.
- **Scatter Charts**: `excel_create_chart` with `chart_type="scatter"` failed because a raw tuple was appended to `chart.series`; it now builds a proper `Series`.
- **Lint Errors in Automation Scripts**: Removed an unused import and documented the intentional `sys.path` import order.
- **Manifest Drift**: Synchronized the MCPB manifest, generated tool reference, skills, and README with the 33 canonical `excel_*` tools.

---

## [2.1.3] - 2026-11-09

### Fixed
- **PyPI Provenance Publishing**: Integrated the official PyPA `pypa/gh-action-pypi-publish@release/v1` in `.github/workflows/publish-pypi.yml` to automatically generate Sigstore cryptographic digital attestations (PEP 740) and publish distributions to PyPI under Trusted Publishing (OIDC).

---

## [2.1.2] - 2026-11-09

### Changed
- **Project Licensing**: Formally relicensed project from MIT License to Apache License, Version 2.0 (`Apache-2.0`), providing explicit patent grants, trademark protection, and corporate supply-chain compatibility.
- **PyPI Provenance & PEP 740 Attestations**: Integrated `astral-sh/attest-action` into `.github/workflows/publish-pypi.yml` to generate cryptographic Sigstore digital attestations alongside distribution packages, providing verified supply-chain provenance on PyPI.

---

## [2.1.1] - 2026-11-09

### Removed
- **Legacy Module Directory (`src/mcp_excel`)**: Removed the legacy compatibility directory and associated shim files, establishing `src/excel_mcp` as the exclusive, canonical Python package source module and minimizing PyPI distribution payload.

### Changed
- **PyPI Distribution Payload Optimization**: Configured Hatchling wheel and sdist targets to bundle exclusively `src/excel_mcp` and required root metadata, explicitly excluding test suites, skills, automation scripts, and workflow configurations.
- **Documentation & Standards Alignment**: Synchronized `AGENTS.md` and module docstrings to adhere strictly to technical English and the canonical `excel_mcp` architecture.

---

## [2.1.0] - 2026-11-09

### Added
- **High-Value Sheet & Data Tools**:
  - `excel_append_rows`: Atomically append batch rows below existing data with automatic bottom detection and formula security checks.
  - `excel_insert_rows_columns`: Insert empty rows or columns at specified 1-based indices.
  - `excel_delete_rows_columns`: Delete rows or columns starting at specified 1-based indices.
  - `excel_format_sheet`: Single-call sheet display and layout configuration (freeze panes, autofilter, gridlines, custom column widths).
  - `excel_create_table`: Native Excel Table generator with banded row styling (`TableStyleInfo`) and filter drop-downs.
  - `excel_sort_range`: In-place 2D range sorting by key column with header preservation and type-safe comparisons.
  - `excel_find_cells`: High-efficiency cell search across sheets matching substring, exact value, or regular expressions without full grid context dump.
  - `excel_profile_sheet`: Lightweight column statistical profiler reporting data types, completeness (null counts), min/max, and distinct samples.
- **Skills Ecosystem & Drift Guard**:
  - `excel-tools-reference` skill: Authoritative encyclopedia and decision tree covering all 30 canonical tools, types, signatures, and safety rules.
  - `scripts/generate-tools-doc.py`: Introspects FastMCP tool registrations and automatically generates `docs/TOOLS.md`.
  - `tests/test_skills_in_sync.py`: Automated drift guard test ensuring skills and documentation never drift from registered tools.
- **Unified `@tool_guard` Decorator**: Centralized error-handling decorator in `utils/guards.py` providing consistent structured JSON error responses.

### Changed
- **Crash-Safe Atomic File I/O**: Refactored all workbook serialization sites to use `atomic_save(wb, target_path)` via isolated temporary files and atomic `os.replace`.
- **Dynamic Compatibility Shim (`mcp_excel`)**: Replaced 24 duplicated, stale files (~4,900 lines) with a lightweight dynamic PEP 562 alias shim and `_AliasFinder` sys.meta_path redirector.
- **Canonical Tool Registration**: FastMCP tools registered with canonical `excel_*` names by default, reducing tool schema tokens for LLMs by >50%. Legacy unprefixed aliases optionally enabled via `EXCEL_MCP_LEGACY_ALIASES=1`.
- **Enhanced Formula Security**: Added Windows DDE command execution patterns (`cmd|`, `powershell|`, `|'`) to `INJECTION_PATTERNS` to mitigate CSV/formula injection vulnerabilities.
- **CI/CD Quality Gate Hardening**: Removed `continue-on-error: true` from `ruff` and `mypy` in `.github/workflows/ci.yml`, making both blocking quality gates.
- **Container Modernization**: Multi-stage `Dockerfile` with Astral `uv`, non-root user `appuser`, standard `/data` volume, and English documentation.
- **Documentation Standardization**: All repository documentation, docstrings, and code comments standardized to Technical English.

---

## [2.0.0] - 2026-11-08

### Added
- **Single-Pass Batch Workbook Builder**: High-performance engine in `excel_mcp.core.batch` capable of generating complete multi-sheet Excel workbooks with title banners, table headers, alternating row colors (zebra striping), accounting borders, auto-calculated totals (`SUM`), currency/number formatting, and native charts in **1 single MCP call** (`excel_build_workbook`), achieving >90% token reduction.
- **Lightweight Outline Inspection**: `excel_get_workbook_outline` for instant structural and dimensional inspection (`A1:H50`, sheet names, headers, formula count) without loading full cell grids into the model context.
- **Batch Matrix Writes**: `excel_batch_write_data` to write 2D arrays and ranges atomically into existing sheets.
- **Visual Design Presets**: Native layout presets (`corporate`, `financial`, `modern`, `compact`) standardizing font families, header fills, zebra striping, currency formats (`$#,##0.00`), and auto-fit column width calculation.
- **Modular Skills Ecosystem**: 3 focused skills in `skills/`:
  - `excel-spreadsheet-creator`: Single-pass workbook creation, financial models, presets, and charts.
  - `excel-spreadsheet-editor`: Lightweight outline inspection, batch updates, formulas, formatting, and pivot tables.
  - `excel-spreadsheet-analyzer`: Structural audits, targeted range sampling, and formula verification without context exhaustion.
- **Astral `uv` Automated Workflows**:
  - `ci.yml`: Fast multi-version testing and linting (Python 3.11, 3.12, 3.13) with cache and `paths-ignore`.
  - `bump-and-release.yml`: 1-click semantic version bumping with PEP 740 artifact attestation and PyPI Trusted Publishing via OIDC.
  - `publish-pypi.yml`: Trusted Publishing OIDC deployment to PyPI upon GitHub release publication.
  - `codeql.yml`: Automated CodeQL security scanning.
  - `release.yml`: GitHub release packaging with automatic changelog note extraction.
- **Automation Scripts**: `scripts/bump-version.py` and `scripts/extract-release-notes.py`.
- **Governance & Security**: `AGENTS.md` and `SECURITY.md`.

### Changed
- **Package Renaming**: Standardized core package name on PyPI to `luiccian-excel-mcp` and project to `excel-mcp`.
- **Dynamic Version Resolution**: `excel_mcp.__version__` resolves dynamically via `importlib.metadata` with clean fallback.
- **Backward-Compatible Shims**: Full backward compatibility maintained for existing imports through `src/mcp_excel/`.
- **Tool Naming Standardization**: All tools expose canonical `excel_` prefixes with legacy unprefixed aliases for backward compatibility.
- **Unified Configuration**: Centralized all project metadata and tools (`ruff`, `pytest`, `mypy`) into `pyproject.toml`.

### Removed
- Removed redundant configuration files (`ruff.toml`, `pytest.ini`, `mypy.ini`).
- Reorganized loose documentation files into `docs/`.
