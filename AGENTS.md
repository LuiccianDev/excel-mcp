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
          "MCP_ALLOWED_DIRECTORIES": "/path/to/spreadsheets"
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

## 2. Single-Pass Token Optimization Directive

- **Prohibition of Chatty Micro-Iterations:** AI assistants using `excel-mcp` must avoid making consecutive micro-calls (e.g., calling `create_workbook`, then 50 separate calls to `write_data_to_excel`, followed by 20 calls to `format_range_excel`). This legacy anti-pattern consumes tens of thousands of tokens and multiple conversational round-trips.
- **Mandatory Single-Pass Usage:** Always prefer:
  1. `excel_build_workbook`: Construct complete, publication-ready workbooks with multiple sheets, title banners, table headers, alternating row colors (zebra striping), accounting borders, auto-calculated totals, number/currency formats, and native charts in **one single MCP tool call** (>90% token savings).
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
     - The tag push triggers `release.yml`, which verifies tag consistency against `pyproject.toml` via `scripts/check-repository-consistency.py --tag "$TAG"`, executes multi-version test suites (Python 3.11–3.14), builds single-pass distributions, creates the GitHub Release, and automatically publishes to PyPI via Trusted Publishing OIDC.
   - **Emergency Recovery:**
     - If PyPI upload fails after the GitHub Release tag already exists, use `.github/workflows/publish-pypi.yml` manually, specifying the existing tag.
