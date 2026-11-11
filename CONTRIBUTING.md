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
