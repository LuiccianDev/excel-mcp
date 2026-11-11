# Release Pipeline Automation Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** One ordered, automatic and verified release chain: merge of a release PR to `master` produces the tag, GitHub Release and PyPI publication, and nothing is released on any other commit.

**Architecture:** Replace the three loosely coupled workflows (`release.yml` tag trigger plus manual dispatch, `publish-pypi.yml` double trigger) with a single `release.yml` of sequential jobs: `detect` (did `pyproject.toml` version change and is the tag missing?) then `verify` (full quality gate on Python 3.11 to 3.14) then `build` (build once, upload artifact) then `github-release` (create tag and release from that artifact) then `publish-pypi` (publish the same artifact via OIDC with PEP 740 attestations). `bump-and-release.yml` keeps preparing the release PR. `publish-pypi.yml` becomes a manual-only recovery workflow.

**Tech Stack:** GitHub Actions, uv, hatchling, pytest, PyPI Trusted Publishing (OIDC, PEP 740).

**Spec:** Audit of `.github/workflows/*` at master `a3aed28` (release v2.1.4) plus the decisions below.

## Decisions (from user)

- Python 3.14 stays supported: matrix and classifiers keep 3.11, 3.12, 3.13, 3.14. The failing symlink test seen on PR #14 was fixed before merge; Task 1 confirms 3.14 is green.
- Release is automatic, ordered and tied to the version-bump commit only.

## Audit Findings (current state)

| # | Workflow | Finding | Risk |
|---|----------|---------|------|
| 1 | `publish-pypi.yml` | Publishes without tests, ruff, mypy or consistency check | Broken build reaches PyPI |
| 2 | `publish-pypi.yml` | Rebuilds the package; Release assets are a different build | Assets and PyPI files may differ |
| 3 | `publish-pypi.yml` | Triggered by `release: published` and by dispatch from `release.yml` | Double publish (failed run 43048323254) |
| 4 | `release.yml` | No check that tag equals `pyproject.toml` version | Tag `v2.1.5` could publish 2.1.4 |
| 5 | `release.yml` | Manual dispatch with free `tag_name`; tag created by the release action on branch HEAD | Tag may point at wrong commit |
| 6 | `release.yml` | `gh workflow run ... \|\| echo` swallows failure | Silent: Release exists, PyPI never published |
| 7 | `bump-and-release.yml` | Release PR created with `GITHUB_TOKEN` plus `[skip ci]` | Release PR has no checks (PR #16 had none) |
| 8 | all | Actions pinned by mutable tag; no `timeout-minutes`; no job-level least privilege | Supply-chain and hang risk |
| 9 | `ci.yml` | CHANGELOG says matrix is 3.11 to 3.13 but matrix has 3.14 | Docs drift |

## Global Constraints

- Technical English for code, comments, commits (AGENTS.md section 0). No informal emojis.
- Workflow-only commits use `ci:`; docs-only commits include `[skip ci]`.
- Package `luiccian-excel-mcp`; source only in `src/excel_mcp/`.
- Releases are prepared only by `bump-and-release.yml` plus `scripts/bump-version.py` (AGENTS.md section 6).
- CI workflows keep `concurrency: cancel-in-progress: true`; the release workflow uses `cancel-in-progress: false`.
- Python floor `>=3.11`; matrix and classifiers 3.11, 3.12, 3.13, 3.14.
- PyPI authentication only through Trusted Publishing (no API tokens).

## Review Focus

1. Docs-only or non-version commit on `master`: expected no tag, no release, no publish.
2. Version bumped but tag `vX.Y.Z` already exists (re-run, retry): expected clean no-op, not a failure or duplicate.
3. Version bumped in `pyproject.toml` but `manifest.json` or `CHANGELOG.md` out of sync: expected failure at `verify`, before any tag exists.
4. PyPI upload fails after the GitHub Release exists: expected a failed run plus a documented manual recovery path (`publish-pypi.yml`), no second release.
5. Two release commits pushed close together: expected sequential runs (concurrency group), never parallel publishes.

---

### Task 1: Pin invariants with tests and confirm Python 3.14

**Files:**
- Create: `tests/test_release_workflows.py`
- Test: `tests/test_release_workflows.py`

**Interfaces:**
- Produces: tests `test_release_triggers_only_on_master_push_and_dispatch`, `test_release_jobs_are_ordered` (needs chain `detect` -> `verify` -> `build` -> `github-release` -> `publish-pypi`), `test_publish_pypi_is_manual_only`, `test_all_jobs_have_timeout`, `test_actions_pinned_to_sha`, `test_matrix_matches_classifiers` (matrix versions equal `Programming Language :: Python :: 3.x` classifiers).

- [ ] **Step 1: Run `uv run --python 3.14 pytest -q`.** Expected: PASS. If it fails, stop and fix before continuing.
- [ ] **Step 2: Write the tests above** parsing YAML with `yaml.safe_load` (add `pyyaml` to the dev group if missing).
- [ ] **Step 3: Run** `uv run pytest tests/test_release_workflows.py -v`. Expected: FAIL.
- [ ] **Step 4: Commit** `test: pin release workflow invariants`.

### Task 2: Version/tag consistency in the script

**Files:**
- Modify: `scripts/check-repository-consistency.py`
- Create: `scripts/release-state.py`
- Test: `tests/test_release_metadata.py`

**Interfaces:**
- Produces: `check-repository-consistency.py --tag vX.Y.Z` exits 1 with `Tag vX does not match pyproject version Y`.
- Produces: `release-state.py --previous-ref <sha>` prints `version=<X.Y.Z>`, `tag=vX.Y.Z`, `should_release=true|false` to stdout in `GITHUB_OUTPUT` format. `should_release` is true only when the `pyproject.toml` version differs from its value at `<sha>` and `git tag -l vX.Y.Z` is empty.

- [ ] **Step 1: Write failing tests** for both scripts (mismatched tag exits 1; unchanged version gives `should_release=false`; existing tag gives `false`; changed version and no tag gives `true`), using a temporary git repository fixture.
- [ ] **Step 2: Run** them. Expected: FAIL.
- [ ] **Step 3: Implement** the two behaviors with the signatures above.
- [ ] **Step 4: Run** `uv run pytest tests/test_release_metadata.py -v`. Expected: PASS.
- [ ] **Step 5: Commit** `ci: add release-state and tag consistency checks`.

### Task 3: Single sequential `release.yml`

**Files:**
- Modify: `.github/workflows/release.yml`

**Interfaces:**
- Consumes: Task 2 scripts.
- Produces: jobs `detect` (outputs `should_release`, `tag`, `version`), `verify` (needs detect, `if should_release`, matrix 3.11 to 3.14, same commands as `ci.yml` plus `--tag`), `build` (single `uv build`, `actions/upload-artifact` as `dist`), `github-release` (`permissions: contents: write`, downloads `dist`, creates tag on `${{ github.sha }}` and release with notes from `scripts/extract-release-notes.py`), `publish-pypi` (`environment: pypi`, `permissions: id-token: write, attestations: write`, downloads `dist`, attests, publishes).

- [ ] **Step 1: Set triggers** to `push` on `master` with the same `paths-ignore` as `ci.yml`, plus `workflow_dispatch` (re-run for the current `pyproject.toml` version; `detect` still refuses an existing tag).
- [ ] **Step 2: Set workflow-level** `permissions: contents: read`, `concurrency: group: release, cancel-in-progress: false`, and `timeout-minutes` on every job.
- [ ] **Step 3: Implement the five jobs** per the Interfaces block; remove the `gh workflow run` dispatch and the `|| echo`.
- [ ] **Step 4: Run** `uv run pytest tests/test_release_workflows.py -v`. Expected: ordering, trigger and timeout tests PASS.
- [ ] **Step 5: Commit** `ci: replace release workflows with one ordered pipeline`.

### Task 4: Manual recovery workflow and release PR checks

**Files:**
- Modify: `.github/workflows/publish-pypi.yml`, `.github/workflows/bump-and-release.yml`

- [ ] **Step 1: `publish-pypi.yml`:** keep only `workflow_dispatch` with required input `tag` (checks out that tag, runs `uv build`, verifies `--tag`, publishes). Add a header comment: recovery only, normal path is `release.yml`.
- [ ] **Step 2: `bump-and-release.yml`:** use `${{ secrets.RELEASE_PAT || secrets.GITHUB_TOKEN }}` for checkout and `gh pr create`; drop `[skip ci]` from the bump commit so CI runs on the release PR; remove the "direct tag" mode (tags are created only by `release.yml`); add `timeout-minutes`.
- [ ] **Step 3: Update** `tests/test_release_metadata.py` expectations that mention `[skip ci]` or direct tagging.
- [ ] **Step 4: Run** `uv run pytest -q`. Expected: PASS.
- [ ] **Step 5: Commit** `ci: make PyPI publish manual-only and run checks on release PRs`.

### Task 5: Action pinning, Dependabot, docs

**Files:**
- Modify: `.github/workflows/*.yml`, `CHANGELOG.md`, `CONTRIBUTING.md`
- Create: `.github/dependabot.yml`

- [ ] **Step 1: Pin by full commit SHA** (with `# vX` comment): `actions/checkout`, `astral-sh/setup-uv`, `actions/upload-artifact`, `actions/download-artifact`, `actions/attest-build-provenance`, `pypa/gh-action-pypi-publish`, `github/codeql-action/*`.
- [ ] **Step 2: Add Dependabot** `github-actions` ecosystem, weekly.
- [ ] **Step 3: Document** the release flow in `CONTRIBUTING.md` (run `Automated Version Bump` -> merge release PR -> pipeline publishes automatically) and fix the CHANGELOG matrix note to 3.11 to 3.14 under `[Unreleased]`.
- [ ] **Step 4: Run** `uv run pytest -q`. Expected: all PASS, including `test_actions_pinned_to_sha`.
- [ ] **Step 5: Commit** `ci: pin actions, add dependabot, document release flow`.

### Task 6: Dry run

- [ ] **Step 1:** Push a docs-only commit to a test branch merged path and confirm `release.yml` shows `should_release=false` (no later jobs).
- [ ] **Step 2:** Dispatch `release.yml` on a version already tagged (`v2.1.4`) and confirm clean no-op.
- [ ] **Step 3:** Cut the next patch through `bump-and-release.yml`, merge the PR, and confirm tag, Release assets and PyPI file hashes match.

---

## Self-Review

- **Coverage:** findings 1 to 9 map to Tasks 3 to 5; Review Focus 1 to 5 map to Tasks 2 and 3 (focus 1, 2, 5), Task 2 and 3 `verify` (focus 3), Task 4 (focus 4).
- **Remaining choice for you:** create the `RELEASE_PAT` secret so release PRs run CI (without it the fallback is `GITHUB_TOKEN` and PRs again have no checks).
