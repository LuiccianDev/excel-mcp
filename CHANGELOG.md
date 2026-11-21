# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

---

## [Unreleased]

---

## [2.4.1] - 2026-11-21

### Security
- **Unified Formula & Cell Injection Validation**: Consolidated pure formula, DDE, and cell-value validation into `utils/validation_utils.py`, enforcing `validate_formula_secure` across `apply_formula`, `validate_formula_in_cell_operation`, `data_validation` (`formula1`/`formula2` across all validation types), and `batch.py` title banners and headers while blocking `_xlfn.`/`_xlws.`/`_xlpm.`/`_xleta.` prefixes and unquoted DDE payloads without false positives on bulleted prose.
- **Archive, Range, Regex & Metrics Resource Bounds**: Added ZIP entry-count, uncompressed-size, and compression-ratio validation in `core/_io.py` (`workbook_session`) plus post-save `DEFAULT_MAX_FILE_SIZE` enforcement in `atomic_save`, shared `assert_range_cell_limit` bounds across `data.py`, `formatting.py`, and `sheet_utils.py`, `preview_only` scan clamping in `_get_used_range`, depth-aware nested-quantifier ReDoS protection and `scan_truncated` reporting in `core/search.py`, and log rotation in `utils/metrics.py`.
- **Path Placeholder Filtering & Destructive Tool Annotations**: Ignored unexpanded `${...}` directory placeholders and trailing-space Windows reserved names in `utils/paths.py`, centralized directory rejection in `validate_file_constraints`, and marked `excel_build_workbook` and `excel_create_workbook` with `destructive_hint=True`.
- **CI/CD Release & Supply-Chain Hardening**: Moved `workflow_dispatch` inputs and outputs into `env:` across release workflows, bound `build` and `github-release` in `release.yml` to the immutable `needs.verify-tag.outputs.sha` commit after verifying `refs/tags/${TAG}^{commit}` and `master` ancestry, qualified `refs/tags/` checkout and enforced `uv lock --check` in `publish-pypi.yml`, disabled shared runner caching (`enable-cache: false`) in release and PyPI publish jobs, restricted the `release/v*` version-bump exemption in `ci.yml` to same-repository PRs, added top-level `contents: read` permissions in `codeql.yml`, and anchored `[project].version` updates with monotonic semver and non-empty `[Unreleased]` validation in `scripts/bump-version.py`.

### Changed
- **Modular Sheet & Formatting Helpers**: Decomposed `_build_sheet_with_stats` (`core/batch.py`) and `format_range` (`core/formatting.py`) into focused helpers under 150 lines per function and removed the `ToolAnnotations` fallback stub in `tools/register_tools.py`.

---

## [2.4.0] - 2026-11-18

### Added
- **Single-Pass Builder Dynamism (`excel_build_workbook`)**: Extended `build_single_pass_sheet` to support configurable `auto_totals` aggregations (`SUM`, `AVERAGE`, `COUNT`, `COUNTA`, `MIN`, `MAX`, `SUBTOTAL` via dict/column-name mapping), inline `data_validations` rule lists, and multiple native `charts` per sheet.
- **Compound Parameter Normalization & Documentation**: Added transparent `range_string`/`cell_range` alias normalization in `dispatch_compound_action` and documented exact `params` schema keys across `excel_data`, `excel_format`, and `excel_analytics`.

### Changed
- **Complete `workbook_session` / `resolve_sheet` Migration**: Migrated `core/chart.py`, `core/pivot.py`, `core/formatting.py`, `core/search.py`, `core/structure.py`, `core/batch.py`, `core/data.py`, `core/workbook.py`, `utils/sheet_utils.py`, and `utils/validation_utils.py` onto `workbook_session`, `atomic_save`, and `resolve_sheet` to eliminate `openpyxl` descriptor leaks and duplicate sheet-resolution helpers.
- **Standardized Tool Guard Coverage**: Applied `@tool_guard()` across `read_data_from_excel`, `excel_tools`, `format_tools`, `formulas_excel_tools`, and `graphics_tools` (resolving `filename` cleanly across `@validate_file_access` positional bindings), removing redundant `try/except` handlers from `format_tools`, `formulas_excel_tools`, `create_excel_worksheet`, and `list_excel_documents` while preserving legacy `read_data_from_excel`, `write_data_to_excel`, `create_excel_workbook`, and `graphics_tools` return contracts.

### Removed
- **Docker Container Packaging**: Removed `Dockerfile`, `.dockerignore`, `docs/Docker.md`, and Docker references from `README.md` and CI path filters.

---

## [2.3.0] - 2026-11-18

### Added
- **Unified Workbook Session Lifecycle**: Added `workbook_session` and `resolve_sheet` in `excel_mcp.core._io` with atomic save on clean write exits, strict guards against `mode="read", save=True` and `data_only=True` writes, and sheet-name validation; migrated `protection`, `page_setup`, `data_validation`, and `workbook` context managers.
- **Root CLI `--tool-mode` Support**: Added root Typer callback support for `--tool-mode` and `--path` forwarding across `server` and `list` commands in `excel_mcp.__main__`.

### Changed
- **Standardized Environment Variables (`EXCEL_MCP_DIRS`, `EXCEL_MCP_TOOLS`)**: Renamed `MCP_ALLOWED_DIRECTORIES` to `EXCEL_MCP_DIRS` and `EXCEL_MCP_MODE` to `EXCEL_MCP_TOOLS` across runtime, CLI, manifests, docs, and tests with no legacy fallback.
- **Compound Tool Mode Naming (`compound`)**: Renamed default `compose` mode (`compose_tools.py`, `COMPOSE_REGISTRY`, `ToolMode.COMPOSE`) to `compound` (`compound_tools.py`, `COMPOUND_REGISTRY`, `ToolMode.COMPOUND`) with no legacy fallback.
- **Single-Source Tool Registry & Mode-Aware Instructions**: Consolidated `full` and `compound` tool metadata into `COMPOUND_REGISTRY` (`CompoundActionSpec`) with `ToolMode` (`StrEnum`), removed duplicated registration tables from `register_tools.py`, and derived `server.py` MCP instructions dynamically from `COMPOUND_REGISTRY`.
- **Documentation & Skill Synchronization**: Updated `README.md` and `skills/excel-tools-reference/SKILL.md` to document `compound` mode (8 default tools) and `full` mode opt-in (`EXCEL_MCP_TOOLS=full`, `server --tool-mode full`), and extended `tests/test_skills_in_sync.py` to guard both tool surfaces.

### Removed
- **Legacy Unprefixed Tool Aliases**: Removed `EXCEL_MCP_LEGACY_ALIASES`, `should_register_legacy_aliases()`, and `legacy_aliases` metadata so only canonical `excel_*` tool names are registered.

---

## [2.2.0] - 2026-11-16

### Added
- **Default Compact `compose` Tool Surface**: Added `compose_tools.py` grouping 30 granular operations into 5 action-dispatched domain tools (`excel_workbook`, `excel_data`, `excel_format`, `excel_formula`, `excel_analytics`) alongside the 3 single-pass tools (`8` tools by default), with `EXCEL_MCP_MODE=full` and `--tool-mode full` opt-in (PR #33).

### Changed
- **Lightweight Release Pipeline**: Streamlined `release.yml` by removing redundant multi-version test execution during tag deployment, relying on the mandatory CI quality gate on PRs and `master`, reducing release jobs from 8 to 4.

---

## [2.1.6] - 2026-11-12

### Added
- **Uniform Typography & Sheet Navigation in `excel_build_workbook`**: Added optional `font_name` and `font_size` overrides plus per-sheet `freeze_panes` and `auto_filter` support in `build_workbook` / `build_single_pass_workbook`, and capped title banner sizing at `12pt`.

### Changed
- **Agent Skills Standardization (`skills/`)**: Updated `excel-spreadsheet-creator`, `excel-spreadsheet-editor`, `excel-spreadsheet-analyzer`, and `excel-tools-reference` to mandate uniform typography (defaulting to `Calibri 11pt` / `Arial 11pt` while honoring user-requested fonts), uncluttered multi-tab/spaced layouts, diverse structural templates, and 100% dynamic Excel formulas with zero hardcoded calculations.
- **Tag-Driven Release Engineering**: Aligned release architecture with PyPA (`pip`) and VS Code standards. Deployment to PyPI and GitHub Releases is triggered exclusively by pushing semver Git tags (`v[0-9]+.[0-9]+.[0-9]+`), eliminating unintended deployments from direct commits on `master`.
- **CI Version Guard**: Added automated check in `ci.yml` that blocks regular PRs attempting direct version bumps in `pyproject.toml`.
- **GitHub Actions Modernization**: Upgraded all pinned workflow actions (`actions/checkout` v7.0.1, `astral-sh/setup-uv` v10.3.3, `github/codeql-action` v4.39.3, `actions/upload-artifact` v7.0.2, `actions/download-artifact` v8.0.2, `actions/attest-build-provenance` v4.2.2) to resolve Node.js runtime deprecation warnings while preserving strict SHA pinning.
- **Static Analysis & Security Linting**: Expanded Ruff rules in `pyproject.toml` to enforce `flake8-bugbear` (`B`), `flake8-bandit` (`S`), and `pyupgrade` (`UP`), and added `uv lock --check` to CI and release workflows.
- **Configuration Validation**: Removed filesystem directory-creation side effects from `FileConfig.validate_directory`.

### Fixed
- **Formula & DDE Injection Hardening**: Added `validate_cell_value_secure` in `excel_mcp.core.calculations` and wired it across `write_data`, `batch_write_matrix`, `build_single_pass_sheet`, and `append_rows_to_sheet` to block leading-whitespace formulas (`" =..."`) and DDE command prefixes (`+cmd|`, `-cmd|`, `@...`).
- **Active Worksheet Guards**: Added explicit `wb.active is None` checks in `batch_write_matrix` and `structure._get_sheet`.

---

## [2.1.5] - 2026-11-11

### Changed
- **Release Pipeline**: `release.yml` is now one ordered pipeline (detect, verify on Python 3.11 to 3.14, build once, GitHub Release, PyPI) that runs only when a push to `master` changes the package version and the tag does not exist. The same artifacts are published to GitHub and PyPI.
- **Manual PyPI Recovery**: `publish-pypi.yml` is manual-only and requires an existing tag.
- **Release PR Checks**: `bump-and-release.yml` runs the full CI on the release branch through `workflow_call`; the CI-skip marker and direct-tag mode were removed.
- **Supply Chain**: GitHub Actions are pinned to commit SHAs, jobs have timeouts and least-privilege permissions.

### Added
- **Release Guards**: `scripts/release-state.py` and `scripts/check-repository-consistency.py --tag` block mismatched tags and duplicate releases.

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
