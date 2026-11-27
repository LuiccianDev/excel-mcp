# MCPB Packaging & Distribution Guide

This document describes how to package, test, and distribute `excel-mcp` (PyPI: `luiccian-excel-mcp`) using the Model Context Protocol Binary/Bundle standard.

---

## 1. Zero-Install Execution (`uvx`)

The primary delivery mechanism for `excel-mcp` is direct execution via `uvx`:

```bash
uvx luiccian-excel-mcp
```

This dynamically fetches the latest tested wheel from PyPI, creates an ephemeral virtual environment, and launches the FastMCP server over `stdio` in `compound` mode (8 tools by default; set `EXCEL_MCP_TOOLS=full` for all 33 granular tools).

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
        "EXCEL_MCP_DIRS": "/path/to/spreadsheets"
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
        "EXCEL_MCP_DIRS": "${workspaceFolder}"
      }
    }
  }
}
```

---

## 3. Bundle Manifest (`manifest.json`)

The server exposes metadata and tool descriptions compliant with MCP manifest specification v0.4.
Local clients can launch the server using the bundled `manifest.json`.
