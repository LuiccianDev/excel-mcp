# AGENTS.md — Rules for AI Assistants

Project: **excel-mcp** (High-Performance Excel .xlsx MCP Server • PyPI: `luiccian-excel-mcp`).

---

## 0. Language, Aesthetics & Documentation Policy

- **Repository Language:** All official repository documentation, code comments, pull requests, issue templates, workflows, and git commits **MUST** be written in **Technical English**.
- **User Interaction:** When interacting with the user, the assistant should reply in **Spanish** (or the user's preferred language) while preserving technical names in English and maintaining a formal engineering tone.
- **Strict Aesthetic Standard (Zero Informal Emojis):** Prohibit informal emojis (such as rockets, party poppers, fire, tool icons) in markdown documentation, section headers, release notes, issue templates, and commit messages. Maintain a sober engineering format utilizing markdown tables and GitHub alerts (`> [!NOTE]`, `> [!TIP]`, `> [!IMPORTANT]`, `> [!WARNING]`).

---

## 1. Zero-Install Tier 1 Priority

When assisting users or configuring clients (Claude Desktop, Cursor, Cline, Copilot), assistants must adhere to the following priority tiers:

- **Tier 1 (Recommended — Zero-Install):** Prioritize direct execution with `uvx luiccian-excel-mcp`. Users do not need to clone the repository or manually manage Python virtual environments.
  ```json
  {
    "mcpServers": {
      "excel-mcp": {
        "command": "uvx",
        "args": ["luiccian-excel-mcp"],
        "env": {
          "EXCEL_MCP_DIRS": "/path/to/spreadsheets"
        }
      }
    }
  }
  ```
- **Tier 2 (Advanced / Local Development):** For contributors developing features or modifying tools locally, use Astral `uv`:
  ```bash
  uv sync --all-groups
  uv run excel-mcp
  ```

---

## 2. Single-Pass Token Optimization & Tool Surface Directive

- **Default `compound` Surface (8 Tools) vs `full` Mode (33 Tools):**
  - By default (`EXCEL_MCP_TOOLS=compound`), the server exposes **8 tools**: 3 single-pass tools (`excel_build_workbook`, `excel_get_workbook_outline`, `excel_batch_write_data`) plus 5 action-dispatched domain tools (`excel_workbook`, `excel_data`, `excel_format`, `excel_formula`, `excel_analytics`).
  - Set `EXCEL_MCP_TOOLS=full` only when a client explicitly requires all 33 individual `excel_*` tool schemas.
  - Compound dispatchers transparently normalize `range_string` and `cell_range` aliases (rejecting ambiguous calls that pass both).
- **Prohibition of Chatty Micro-Iterations:** AI assistants using `excel-mcp` must avoid making consecutive micro-calls (e.g., calling `create_workbook`, then 50 separate calls to `write_data_to_excel`, followed by 20 calls to `format_range_excel`). This legacy anti-pattern consumes tens of thousands of tokens and multiple conversational round-trips.
- **Mandatory Single-Pass Usage:** Always prefer:
  1. `excel_build_workbook`: Construct complete, publication-ready workbooks with multiple sheets, title banners, table headers, alternating row colors (zebra striping), accounting borders, configurable `auto_totals` aggregations (`SUM`, `AVERAGE`, `COUNT`, `COUNTA`, `MIN`, `MAX`, `SUBTOTAL` via bool, list, or per-column dict), currency/number formats, inline `data_validations` rules, and multiple native `charts` (`chart` or `charts: list[dict]`) in **one single MCP tool call** (>90% token savings).
  2. `excel_get_workbook_outline`: Extract structural bounds, dimensions, column headers, and formula summaries first without loading full cell matrices into the context window.
  3. `excel_batch_write_data`: Ingest 2D data matrices and ranges into existing sheets in a single atomic pass.

---

## 3. Package Scope & Isolation Boundary

The distribution boundary between the PyPI package and the repository is strictly defined:

- **PyPI Package (`luiccian-excel-mcp`):** Bundles **exclusively** the runtime code located in `src/excel_mcp/`, accompanied by `README.md`, `LICENSE`, and `pyproject.toml`.
- **Repository Scope:** Test suites (`tests/`), automation scripts (`scripts/`), agent skills (`skills/`), documentation manuals (`docs/`), and GitHub Actions workflows (`.github/`) belong solely to repository infrastructure and are explicitly excluded from the PyPI wheel and sdist distributions via Hatchling build targets.

---

## 4. Official Naming & Hierarchy Standards

To guarantee consistency, cleanliness, and maintainability across the project:

### 1. Root & Module Hierarchy
- **Source Module:** Python source resides strictly in `src/excel_mcp/`.
- **Package Name on PyPI:** `luiccian-excel-mcp` (Distribution identifier on PyPI).
- **Command-line Scripts (Entrypoints):** Registered in `pyproject.toml` pointing to `excel_mcp.__main__:main`:
  - `luiccian-excel-mcp` (exact distribution mirror CLI)
  - `excel-mcp` (standard CLI shorthand)
- **Canonical Naming & Usage Matrix:**
  | Context | Official Name | Standard Invocation / Config |
  |---|---|---|
  | **PyPI Package** | `luiccian-excel-mcp` | `pip install luiccian-excel-mcp` / `uvx luiccian-excel-mcp` |
  | **CLI Command** | `luiccian-excel-mcp` & `excel-mcp` | `uvx --from luiccian-excel-mcp luiccian-excel-mcp` |
  | **Python Import** | `excel_mcp` | `import excel_mcp` |
  | **MCP Server ID** | `excel-mcp` | Configuration key in `claude_desktop_config.json` |
- **Support Folders (Lowercase):**
  - `docs/` (technical documentation and tool specifications)
  - `scripts/` (release, versioning, and build automation)
  - `skills/` (modular skills for LLMs: `excel-spreadsheet-creator`, `excel-spreadsheet-editor`, `excel-spreadsheet-analyzer`, `excel-tools-reference`)
  - `tests/` (automated test suites)
  - `.github/` (CI/CD and security workflows)

### 2. Python Code Standards (PEP 8)
- **Module & File Names:** Strictly `snake_case.py` (e.g., `batch.py`, `presets.py`, `response_models.py`). Never use PascalCase or camelCase for `.py` files.
- **Classes:** `PascalCase` (`WorkbookOutline`, `BuildWorkbookResult`).
- **Functions & Variables:** `snake_case` (`build_workbook()`, `get_preset()`).
- **Constants:** `SCREAMING_SNAKE_CASE` (`PRESETS`, `DEFAULT_ALLOWED_DIRS`).
- **Automation Scripts:** Strictly `kebab-case.py` in `scripts/` (e.g., `bump-version.py`, `extract-release-notes.py`).

### 3. Workbook Lifecycle, Error Boundaries & Layering Discipline
- **Mandatory `workbook_session` & `resolve_sheet` (`src/excel_mcp/core/_io.py`):**
  - Never call raw `openpyxl.load_workbook()` or `openpyxl.Workbook()` directly in feature code without `workbook_session(filename, mode="read"|"edit"|"create"|"upsert")`.
  - Always resolve worksheets via `resolve_sheet(wb, sheet_name, create_if_missing=...)` and rely on `workbook_session` / `atomic_save` so file handles (`wb.close()`) are deterministically closed on every path and writes commit via atomic `os.replace`.
  - For empty-input short-circuits (`not query`, `not rows`), validate file existence with `Path(filename).resolve().exists()` instead of opening and parsing the `.xlsx` archive.
- **Standardized `@tool_guard()` Coverage & Unified Error Contract (`src/excel_mcp/utils/guards.py`):**
  - Decorate dict-returning MCP tool functions in `src/excel_mcp/tools/` with `@tool_guard()`.
  - Every new or refactored tool in `src/excel_mcp/tools/` must return `dict[str, Any]` with `status: "success" | "error"` and delegate exception handling to `@tool_guard()` (`ExceptionTool.handle_error`) rather than duplicating ad-hoc `try/except` blocks.
  - When combined with `@validate_file_access(...)`, place `@validate_file_access(...)` above `@tool_guard()` so unhandled core exceptions preserve `error_type`, `operation`, and positional `filename` via `inspect.signature.bind_partial`.
- **Strict Layering, Core Consolidation & Circular-Import Prevention (`core/` vs `utils/`):**
  - `src/excel_mcp/core/` owns all workbook I/O, worksheet mutations, `DataValidation` rules (`apply_data_validation_to_sheet`), and domain exceptions (`excel_mcp.exceptions.exception_core`).
  - Any new operation that opens, inspects, or mutates `.xlsx` files must reside in `src/excel_mcp/core/`; `src/excel_mcp/utils/` is reserved for pure coordinate/path/security helpers and execution guards without direct `openpyxl` workbook I/O.
  - Modules inside `src/excel_mcp/utils/` must **never** import `excel_mcp.core.*` at module top level (preventing circular imports when `core/__init__.py` loads before `utils/__init__.py`).
- **Single-Source Tool Inventory Synchronization:**
  - `COMPOUND_REGISTRY` in `src/excel_mcp/tools/compound_tools.py` is the single source of truth for the 30 domain-dispatched tools (plus the 3 single-pass tools = 33 `full` / 8 `compound`).
  - Whenever a tool signature or docstring changes, regenerate `docs/TOOLS.md` via `python scripts/generate-tools-doc.py` and verify synchronization with `python scripts/check-repository-consistency.py` and `pytest tests/test_skills_in_sync.py`.

### 4. Modularity, Size Limits & Pre-PR Verification Gate
- **Module & Function Size Limits:**
  - Keep modules in `src/excel_mcp/` under **600 lines** (hard ceiling: 1,000 lines) and individual functions under **150 lines**.
  - Extract focused, pure helpers for parameter normalization, column resolution (`_resolve_col_index`), and rule validation (`apply_data_validation_to_sheet`) instead of growing monolithic functions or copy-pasting logic across modules.
- **Graphify Navigation & Documentation Synchronization:**
  - Query Graphify MCP tools (`query_graph`, `get_node`, `get_neighbors`, `shortest_path`, `god_nodes`) and `graphify-out/GRAPH_REPORT.md` first to locate architectural hubs instead of reading entire source files.
  - Whenever modules, tools, or behavior change, update `README.md`, `docs/`, `skills/`, and the knowledge graph (`uvx --from graphifyy graphify update .`).
- **Mandatory Pre-PR Audit & Quality Gate:**
  - Before opening any Pull Request that touches `src/` or `tests/`, verify that every parameter documented in `compound_tools.py` docstrings matches the underlying handler signature, run a branch diff audit (e.g., Thermos bug/security + code-quality review), and pass the full local verification gate:
    1. `uv run python scripts/generate-tools-doc.py`
    2. `uv run ruff format --check .` & `uv run ruff check .`
    3. `uv run mypy src`
    4. `uv run python scripts/check-repository-consistency.py`
    5. `uv run pytest -q`

---

## 5. Optimization of Minutes and Compute in GitHub Actions

To preserve compute quotas and prevent unnecessary runner execution in GitHub Actions:

1. **Exclusion of Non-executable Files (`paths-ignore`):**
   - Workflows (`ci.yml`, `codeql.yml`) must ignore changes that exclusively touch documentation (`**.md`, `docs/**`, `skills/**`), licensing, or git metadata (`.gitignore`, `.gitattributes`).
2. **Mandatory Use of `[skip ci]`:**
   - In commits that are strictly documentary, markdown-related, or metadata refinements that do not affect runtime Python code or tests, always include `[skip ci]` or `[ci skip]` in the commit message.
3. **Automatic Cancellation of Stale Runs (`concurrency`):**
   - Workflows must maintain `concurrency: cancel-in-progress: true` to immediately terminate superseded jobs when new commits are pushed to the same branch.

---

## 6. Semantic Versioning & CHANGELOG Maintenance

1. **SemVer 2.0.0 & Keep a Changelog:**
   - Adhere strictly to [Semantic Versioning](https://semver.org/) (`MAJOR.MINOR.PATCH`) and [Keep a Changelog](https://keepachangelog.com/).
   - Valid changelog subsections: `Added`, `Changed`, `Deprecated`, `Removed`, `Fixed`, `Security`.
2. **Unreleased Section (`[Unreleased]`):**
   - Active developments not yet published with a git tag must be recorded under `## [Unreleased]`.
3. **Release Execution Paths (Tag-Driven Deployment):**
   - Releases to PyPI and GitHub Releases are triggered exclusively by pushing an official Git tag `vX.Y.Z` (`push: tags: ['v[0-9]+.[0-9]+.[0-9]+']`), ensuring zero accidental PyPI deployments on ordinary `master` commits or PR merges.
   - **PR Protection Guard:** Regular Pull Requests are strictly prohibited from bumping `version = ` in `pyproject.toml` (enforced by `ci.yml`). Version increments must only happen in official release branches (`release/v*`).
   - **Path A (Automated Bump & Release PR — Recommended):**
     1. Run `.github/workflows/bump-and-release.yml` (specifying `bump_type`: `patch`, `minor`, `major`, or `custom`).
     2. The workflow atomically bumps `pyproject.toml`, `manifest.json`, `uv.lock`, and `CHANGELOG.md`, runs full CI on the release branch (`workflow_call`), and opens a release PR.
     3. Review and merge the release PR to `master`.
     4. Tag the verified `master` commit (`git tag vX.Y.Z && git push origin vX.Y.Z`) to trigger `.github/workflows/release.yml`.
   - **Path B (Direct Tag Push):**
     - Once `master` is verified and synchronized, creating and pushing the tag `vX.Y.Z` triggers `release.yml` directly.
   - **Execution Sequence:**
     - The tag push triggers `release.yml`, which verifies tag consistency against `pyproject.toml` via `scripts/check-repository-consistency.py --tag "$TAG"`, validates lockfile consistency, builds single-pass distributions, creates the GitHub Release, and automatically publishes to PyPI via Trusted Publishing OIDC (with the multi-version test suite across Python 3.11–3.14 strictly enforced upstream by CI on PRs and master).
   - **Emergency Recovery:**
     - If PyPI upload fails after the GitHub Release tag already exists, use `.github/workflows/publish-pypi.yml` manually, specifying the existing tag.
