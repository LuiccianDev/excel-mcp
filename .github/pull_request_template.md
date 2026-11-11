## Summary

Describe the change and the user-visible behavior it adds or fixes.

## Verification

- [ ] `uv run ruff check .`
- [ ] `uv run ruff format --check .`
- [ ] `uv run mypy src`
- [ ] `uv run pytest -q --cov --cov-report=term-missing`
- [ ] Packaging or documentation checks were run when applicable.

## Reproduction and test evidence

For bug fixes, include the original reproduction and the regression test.
For features, include the important input cases and expected output.

## Security impact

Describe any path, file, formula, secret, permission, or dependency impact.
Write `None` when the change has no security impact.

## Compatibility and release notes

- Breaking change: Yes / No
- [ ] `CHANGELOG.md` updated when required.
- [ ] Generated documentation and manifest are synchronized.
