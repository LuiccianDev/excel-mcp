# Security Policy

## Supported Versions

Security updates and patches are applied to the following versions of `excel-mcp` (PyPI: `luiccian-excel-mcp`):

| Version | Supported          |
| ------- | ------------------ |
| 2.0.x   | :white_check_mark: |
| < 2.0.0 | :x:                |

---

## Reporting a Vulnerability

We take the security of `excel-mcp` seriously. If you discover a potential vulnerability, please report it responsibly:

1. **Do not create public GitHub issues** for security vulnerabilities.
2. Email the maintainer directly at **luicciandev@gmail.com** with:
   - A detailed description of the vulnerability.
   - Steps or proof-of-concept script to reproduce the issue.
   - The affected version(s) and environment details.
3. You will receive an acknowledgment within 48 hours.
4. We will coordinate a patched release before any public disclosure.

---

## Security Boundaries & Best Practices

- **Allowed Directories (`MCP_ALLOWED_DIRECTORIES`):**
  The server strictly validates all file operations against permitted directories configured in `MCP_ALLOWED_DIRECTORIES`. Attempted path traversals (`../`) or operations outside permitted roots are rejected.
- **Formula Security:**
  All formulas are sanitized and validated against malicious payload injection before application to spreadsheet cells.
- **Trusted Publishing (OIDC):**
  All releases to PyPI use OpenID Connect (OIDC) Trusted Publishing with PEP 740 cryptographic attestation—no static API tokens are stored in repository secrets.
