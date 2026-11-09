# Excel MCP Server - Docker Deployment Guide

This document describes how to containerize and run the `excel-mcp` (PyPI: `luiccian-excel-mcp`) Model Context Protocol server using Docker, enabling secure `stdio` communication with MCP clients such as Claude Desktop and Cursor.

---

## 1. Build the Docker Image

1. Open a terminal in the project root directory.

2. Build the Docker container image:

   ```bash
   docker build -t excel-mcp .
   ```

   This builds the `excel-mcp` image with all required dependencies installed via Astral `uv`.

---

## 2. Run the Excel MCP Server with Docker

To launch the container interactively over standard input/output (`stdio`):

```bash
docker run --rm -i \
  -v /path/to/your/spreadsheets:/spreadsheets \
  -e MCP_ALLOWED_DIRECTORIES="/spreadsheets" \
  excel-mcp
```

- `-i`: Kept interactive so the container accepts JSON-RPC messages via `stdio`.
- `-v`: Mounts the local directory containing spreadsheets. **Replace `/path/to/your/spreadsheets` with your local folder path**.
- `-e MCP_ALLOWED_DIRECTORIES`: Restricts server file operations strictly to the mounted volume for defense-in-depth security.

---

## 3. Integration with Claude Desktop

Add the following block to your `claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "excel-mcp": {
      "command": "docker",
      "args": [
        "run", "--rm", "-i",
        "-v", "/path/to/your/spreadsheets:/spreadsheets",
        "-e", "MCP_ALLOWED_DIRECTORIES=/spreadsheets",
        "excel-mcp"
      ]
    }
  }
}
```

> [!NOTE]
> Adjust `/path/to/your/spreadsheets` according to your local operating system path.

---

## 4. Usage Flow

1. Launch Claude Desktop or any MCP-compliant client.
2. The client spawns the Docker process using the configured arguments.
3. Communication occurs over `stdio` without exposing network ports.
4. You can create, read, batch-edit, format, and audit Excel (`.xlsx`) workbooks using natural language.

---

## 5. Security & Isolation

- **Directory Boundary:** Always specify `MCP_ALLOWED_DIRECTORIES` to confine file operations strictly to mounted directories.
- **Symlink Protection:** The server resolves all paths and symlinks to prevent path traversal outside designated volumes.
- **Least Privilege:** Mount volumes as read-only (`:ro`) if only inspection and analysis tools are required.

---

## 6. Resources

- [Model Context Protocol Specification](https://modelcontextprotocol.io)
- [excel-mcp GitHub Repository](https://github.com/LuiccianDev/excel-mcp)
- [luiccian-excel-mcp on PyPI](https://pypi.org/project/luiccian-excel-mcp/)
