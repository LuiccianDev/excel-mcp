# MCPB Packaging & Distribution Guide

This document describes how to package, test, and distribute `excel-mcp` (PyPI: `luiccian-excel-mcp`) using the Model Context Protocol Binary/Bundle standard.

---

## 1. Zero-Install Execution (`uvx`)

The primary delivery mechanism for `excel-mcp` is direct execution via `uvx`:

```bash
uvx luiccian-excel-mcp
```

This dynamically fetches the latest tested wheel from PyPI, creates an ephemeral virtual environment, and launches the FastMCP server over `stdio`.

---

## 2. Desktop & IDE Integration

### Claude Desktop Configuration

Add the following to `claude_desktop_config.json`:

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

### Cursor Configuration (`.cursor/mcp.json`)

```json
{
  "mcpServers": {
    "excel-mcp": {
      "command": "uvx",
      "args": ["luiccian-excel-mcp"],
      "env": {
        "MCP_ALLOWED_DIRECTORIES": "${workspaceFolder}"
      }
    }
  }
}
```

---

## 3. Bundle Manifest (`manifest.json`)

The server exposes metadata and tool descriptions compliant with MCP manifest specification v0.4.
Local clients can launch the server using the bundled `manifest.json`.
