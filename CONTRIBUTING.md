# Contributing to Excel MCP

## Development setup

Use Astral `uv` for all local commands:

```bash
uv sync --all-groups
uv run pytest -q
uv run ruff check .
uv run ruff format --check .
uv run mypy src
```

## Pull requests

1. Create a focused branch from `master`.
2. Add or update tests before changing behavior.
3. Run the complete `uv` quality gates locally.
4. Update generated documentation with its generator when tool signatures change.
5. Describe security, compatibility, and performance impact in the pull request.

Do not add runtime dependencies without documenting the reason and updating
`uv.lock` with `uv sync --all-groups`.

## Releasing

1. Run the **Automated Version Bump & Release Preparation** workflow (patch, minor, major or custom).
2. Review the release pull request. The full CI runs on the release branch.
3. Merge the pull request. `release.yml` detects the version change, verifies the commit on Python 3.11 to 3.14, builds once, creates the tag and GitHub Release, and publishes the same artifacts to PyPI through Trusted Publishing.

Commits that do not change the version never release. If only the PyPI upload fails, run **Publish to PyPI (manual recovery)** with the existing tag.
