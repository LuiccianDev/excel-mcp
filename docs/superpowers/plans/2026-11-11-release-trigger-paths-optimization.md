# Release Trigger Optimization & Filter Hardening Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Optimize `release.yml` so that it ONLY triggers on pushes to `master` that modify `pyproject.toml` (`paths: ['pyproject.toml']`), eliminating all unnecessary workflow wakeups, detected/skipped checks, and runner usage on standard code, test, documentation, or maintenance commits.

**Architecture:**
1. In `.github/workflows/release.yml`, replace the broad `paths-ignore` trigger with a strict `paths: ['pyproject.toml']` filter under `push: branches: [ master ]`.
2. Keep `scripts/release-state.py` as the second deterministic defense line in `detect` to verify whether the version actually changed and tag does not exist.
3. Keep `workflow_dispatch` for manual triggers when needed.
4. Update `tests/test_release_workflows.py` to assert that `paths: ['pyproject.toml']` is enforced.
5. Verify tests and commit.

**Tech Stack:** GitHub Actions workflow syntax, Python test suite (`pytest`).

---

### Task 1: Update workflow test assertions to require `paths: ['pyproject.toml']`

**Files:**
- Modify: `tests/test_release_workflows.py`

- [ ] **Step 1: Add/update test in `tests/test_release_workflows.py`**
  Assert that `paths:` includes `pyproject.toml` under `push:`.
- [ ] **Step 2: Run test to see failure (RED)**
  Run: `pytest tests/test_release_workflows.py -k test_release_triggers -q`
  Expected: FAIL (currently using `paths-ignore`).

### Task 2: Configure `paths: ['pyproject.toml']` in `release.yml`

**Files:**
- Modify: `.github/workflows/release.yml`

- [ ] **Step 1: Replace trigger block in `release.yml`**
  ```yaml
  on:
    push:
      branches: [ master ]
      paths:
        - 'pyproject.toml'
    workflow_dispatch:
  ```
- [ ] **Step 2: Run tests to verify pass (GREEN)**
  Run: `pytest tests/test_release_workflows.py tests/test_ci_configuration.py tests/test_release_metadata.py -q`
  Expected: All PASS.

### Task 3: Commit and Push to Master

- [ ] **Step 1: Commit changes**
  ```bash
  git commit -am "ci(release): restrict push triggers strictly to pyproject.toml changes"
  ```
- [ ] **Step 2: Push to master**
  Run: `git push origin master`
- [ ] **Step 3: Confirm that `release.yml` does NOT wake up or produce skipped checks on this push**
  Run: `gh run list --limit 3`
