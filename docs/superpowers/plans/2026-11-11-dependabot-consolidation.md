# Dependabot Consolidation & Action Update Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Clean up the 5 scattered open Dependabot PRs (#18 to #22), configure Dependabot to run strictly once a month and produce at most a single grouped PR (`open-pull-requests-limit: 1`, `groups:`), and upgrade existing workflow actions to their latest stable node24-compatible releases pinned by SHA.

**Architecture:** 
1. Close PRs #18, #19, #20, #21, #22 on GitHub.
2. Update `.github/dependabot.yml` to:
   - Interval: `monthly` (first day of the month).
   - Grouping: `groups: github-actions: patterns: ["*"]` so all updates are merged into 1 single PR.
   - Limit: `open-pull-requests-limit: 1`.
3. Directly update workflows to latest SHAs for actions (`actions/checkout`, `astral-sh/setup-uv`, `actions/upload-artifact`, `actions/download-artifact`, `github/codeql-action`) so warnings about Node.js 20 deprecation disappear.
4. Verify tests (`tests/test_release_workflows.py`, `tests/test_ci_configuration.py`) pass.

**Tech Stack:** GitHub Actions, Dependabot YAML v2, GitHub CLI (`gh`), pytest.

**Spec:** User directive: "solo una vez por mes para actualizar de manera adecuada y evitar errores... en un solo PR".

## Global Constraints

- Technical English for git commits and workflow documentation (AGENTS.md section 0).
- Prohibit informal emojis.
- Actions pinned to 40-character commit SHA (enforced by `test_actions_pinned_to_sha`).
- Keep monthly grouped configuration in `.github/dependabot.yml`.

## Review Focus

1. Dependabot opening multiple PRs: prevented by `groups` and `open-pull-requests-limit: 1`.
2. Frequent weekly interruptions: prevented by `interval: "monthly"`.
3. Tests checking action SHA pinning: must pass with new SHAs.
4. CI workflow runs: must continue executing cleanly on push.

---

### Task 1: Close current spam PRs from Dependabot

**Files:**
- GitHub PRs #18, #19, #20, #21, #22

- [ ] **Step 1: Close PRs #18 to #22 using GitHub CLI**
  ```bash
  gh pr close 18 --comment "Superseded by consolidated monthly dependabot configuration"
  gh pr close 19 --comment "Superseded by consolidated monthly dependabot configuration"
  gh pr close 20 --comment "Superseded by consolidated monthly dependabot configuration"
  gh pr close 21 --comment "Superseded by consolidated monthly dependabot configuration"
  gh pr close 22 --comment "Superseded by consolidated monthly dependabot configuration"
  ```
- [ ] **Step 2: Confirm open PR list is 0**
  Run: `gh pr list`
  Expected: No open PRs.

### Task 2: Configure Dependabot for 1 grouped PR per month

**Files:**
- Modify: `.github/dependabot.yml`

- [ ] **Step 1: Update `.github/dependabot.yml`**
  ```yaml
  version: 2
  updates:
    - package-ecosystem: "github-actions"
      directory: "/"
      schedule:
        interval: "monthly"
      open-pull-requests-limit: 1
      groups:
        github-actions:
          patterns:
            - "*"
      commit-message:
        prefix: "ci"
  ```
- [ ] **Step 2: Verify YAML structure and commit**
  ```bash
  git add .github/dependabot.yml
  git commit -m "ci(dependabot): configure monthly grouped updates with max 1 open PR"
  ```

### Task 3: Update action pins to latest stable Node 24 versions

**Files:**
- Modify: `.github/workflows/ci.yml`
- Modify: `.github/workflows/release.yml`
- Modify: `.github/workflows/bump-and-release.yml`
- Modify: `.github/workflows/codeql.yml`
- Modify: `.github/workflows/publish-pypi.yml`

- [ ] **Step 1: Check latest commit SHAs for target actions**
  Resolve latest git SHAs for:
  - `actions/checkout` (v4/v5/v6 latest)
  - `astral-sh/setup-uv`
  - `actions/upload-artifact` & `download-artifact`
  - `github/codeql-action`
  - `actions/attest-build-provenance`
  - `pypa/gh-action-pypi-publish`
- [ ] **Step 2: Update workflow files with the new SHAs**
- [ ] **Step 3: Run workflow and repo tests**
  Run: `pytest tests/test_release_workflows.py tests/test_ci_configuration.py -q`
  Expected: All PASS.
- [ ] **Step 4: Commit**
  ```bash
  git commit -am "ci: upgrade workflow actions to latest stable node24-compatible releases"
  ```

### Task 4: Push to master and verify CI runs clean

- [ ] **Step 1: Push changes to master**
  Run: `git push origin master`
- [ ] **Step 2: Verify `release.yml` detect step skips release cleanly (no version change)**
  Run: `gh run list --limit 3`
- [ ] **Step 3: Confirm no new spam PRs are opened by Dependabot**
