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
    # A squash-merged release PR reuses the bump commit message; "[skip ci]" in it
    # would also suppress the push-triggered release workflow.
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
