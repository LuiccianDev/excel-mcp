"""Contracts for the ordered, automatic release pipeline."""

from __future__ import annotations

import re
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
WORKFLOWS = ROOT / ".github" / "workflows"
RELEASE_JOBS = ["verify-tag", "build", "github-release", "publish-pypi"]


def _read(name: str) -> str:
    return (WORKFLOWS / name).read_text(encoding="utf-8")


def _jobs(text: str) -> dict[str, str]:
    """Return job name -> job block for a workflow (indent-based, no YAML dep)."""
    body = text.split("\njobs:\n", 1)[1]
    parts = re.split(r"^  ([A-Za-z0-9_-]+):[ \t]*$", body, flags=re.MULTILINE)
    return {parts[i]: parts[i + 1] for i in range(1, len(parts), 2)}


def _triggers(text: str) -> str:
    return text.split("\njobs:", 1)[0].split("\non:\n", 1)[1]


def test_release_triggers_only_on_tags_and_dispatch() -> None:
    triggers = _triggers(_read("release.yml"))
    assert "push:" in triggers
    assert "tags:" in triggers
    assert re.search(
        r"tags:\s*\n\s*-\s*['\"]?v\[0-9\]\+\.\[0-9\]\+\.\[0-9\]\+['\"]?", triggers
    )
    assert "workflow_dispatch:" in triggers
    for forbidden in (
        "branches:",
        "paths:",
        "paths-ignore:",
        "release:",
        "pull_request:",
        "schedule:",
    ):
        assert forbidden not in triggers


def test_release_jobs_are_ordered() -> None:
    jobs = _jobs(_read("release.yml"))
    assert list(jobs) == RELEASE_JOBS
    for previous, current in zip(RELEASE_JOBS, RELEASE_JOBS[1:], strict=False):
        assert re.search(rf"needs:\s*\[?[^\n]*\b{previous}\b", jobs[current]), (
            f"{current} must need {previous}"
        )


def test_release_serializes_runs_without_cancelling() -> None:
    release = _read("release.yml")
    assert "cancel-in-progress: false" in release
    assert "permissions:\n  contents: read" in release


def test_publish_pypi_is_manual_only() -> None:
    triggers = _triggers(_read("publish-pypi.yml"))
    assert "workflow_dispatch:" in triggers
    for forbidden in ("release:", "push:", "pull_request:", "tags:"):
        assert forbidden not in triggers
    assert "tag:" in triggers


def test_all_jobs_have_timeout() -> None:
    for workflow in sorted(WORKFLOWS.glob("*.yml")):
        for name, block in _jobs(workflow.read_text(encoding="utf-8")).items():
            if re.search(r"^    uses:", block, flags=re.MULTILINE):
                continue  # reusable-workflow call: timeout belongs to the callee
            assert "timeout-minutes:" in block, f"{workflow.name}:{name}"


def test_actions_pinned_to_sha() -> None:
    pattern = re.compile(r"^\s*-?\s*uses:\s*([^\s#]+)", re.MULTILINE)
    for workflow in sorted(WORKFLOWS.glob("*.yml")):
        for ref in pattern.findall(workflow.read_text(encoding="utf-8")):
            if ref.startswith("./"):
                continue
            assert re.search(r"@[0-9a-f]{40}$", ref), f"{workflow.name}: {ref}"


def test_matrix_matches_classifiers() -> None:
    ci = _read("ci.yml")
    matrix = re.search(r"python-version:\s*\[([^\]]+)\]", ci)
    assert matrix
    versions = {v.strip().strip("\"'") for v in matrix.group(1).split(",")}
    pyproject = (ROOT / "pyproject.toml").read_text(encoding="utf-8")
    classifiers = set(
        re.findall(r"Programming Language :: Python :: (3\.\d+)\"", pyproject)
    )
    assert versions == classifiers


def test_release_commit_does_not_skip_ci() -> None:
    bump = _read("bump-and-release.yml")
    assert "[skip ci]" not in bump
    assert "[ci skip]" not in bump


def test_release_pr_runs_ci_through_workflow_call() -> None:
    bump = _read("bump-and-release.yml")
    ci = _read("ci.yml")
    assert "workflow_call:" in _triggers(ci)
    assert "uses: ./.github/workflows/ci.yml" in bump
    assert "create_pr" not in bump
    assert "--tags" not in bump


def test_bump_and_release_uses_force_with_lease_and_checks_existing_pr() -> None:
    bump = _read("bump-and-release.yml")
    assert "--force-with-lease" in bump
    assert "--force\n" not in bump
    assert "gh pr list" in bump
    assert "|| echo" not in bump


def test_publish_pypi_validates_tag_format_and_master_ancestry() -> None:
    publish = _read("publish-pypi.yml")
    assert "^v[0-9]+\\.[0-9]+\\.[0-9]+$" in publish
    assert "fetch-depth: 0" in publish
    assert "git merge-base --is-ancestor" in publish


def test_release_binds_build_and_github_release_to_verified_sha() -> None:
    release = _read("release.yml")
    jobs = _jobs(release)
    assert "^v[0-9]+\\.[0-9]+\\.[0-9]+$" in jobs["verify-tag"]
    assert "refs/tags/${TAG}^{commit}" in jobs["verify-tag"]
    assert "ref: ${{ needs.verify-tag.outputs.sha }}" in jobs["build"]
    assert "ref: ${{ needs.verify-tag.outputs.sha }}" in jobs["github-release"]
    assert "enable-cache: false" in jobs["verify-tag"]
    assert "enable-cache: false" in jobs["build"]


def test_publish_pypi_qualifies_tag_ref_checks_lockfile_and_disables_cache() -> None:
    publish = _read("publish-pypi.yml")
    assert "ref: refs/tags/${{ github.event.inputs.tag }}" in publish
    assert "uv lock --check" in publish
    assert "enable-cache: false" in publish


def test_ci_version_bump_guard_rejects_cross_repo_forks_and_codeql_permissions() -> (
    None
):
    ci = _read("ci.yml")
    codeql = _read("codeql.yml")
    assert "github.event.pull_request.head.repo.full_name == github.repository" in ci
    assert "permissions:\n  contents: read" in codeql
