# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

---

## [Unreleased]

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
