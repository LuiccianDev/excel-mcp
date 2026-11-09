
# MCP Excel Office

This repository implements an MCP (Model Context Protocol) component designed to read, create, and modify Microsoft Excel files automatically.
Below you will find the steps to build the Docker image and connect it with Claude Desktop as an MCP tool via `stdio`.

---

## 1. **Build the Docker Image**

1. Open a terminal in the project root.

2. Run the following command to build the Docker image:

   ```bash
   docker build -t mcp-excel-server .
   ```

   This creates the image named `mcp-excel-server` with all requirements installed.

---

## 2. **Run the MCP Excel Office Server with Docker**

To launch the container and allow communication via standard input/output (`stdio`):

```bash
docker run --rm -i \
  -v /path/to/your/documents:/documents \
  -e MCP_ALLOWED_DIRECTORIES="/documents" \
  mcp-excel-server
```

- `-i`: Allows the container to accept communication via stdio.
- `-v`: Mounts the folder with Excel documents. **Change `/path/to/your/documents` to the path where your files are located**.
- `-e MCP_ALLOWED_DIRECTORIES`: Restricts access to the mounted folder for greater security.

---

## 3. **Integration with Claude Desktop**

For Claude Desktop (or any MCP-compatible client via stdio) to automatically launch the MCP Word Office server and communicate via standard input/output, add the following configuration to the MCP tools file in Claude Desktop:

```json
{
  "mcp-excel": {
    "command": "docker",
    "args": [
      "run", "--rm", "-i",
      "-v", "/path/to/your/documents:/documents",
      "--name", "mcp-excel",
      "-e", "MCP_ALLOWED_DIRECTORIES=/documents",
      "mcp-excel-server"
    ],
    "type": "stdio"
  }
}
```

- **Adjust the path `/path/to/your/documents` according to your system.**
- The type `"stdio"` indicates that communication will be via standard input/output, not over the network.

---

## 4. **Usage Flow**

1. Start Claude Desktop.
2. Claude runs the above Docker command to launch the MCP Excel Office server.
3. All communication occurs via stdio (no need to open or configure ports).
4. You can manipulate Excel (.xlsx) files directly from Claude Desktop using natural commands.

---

## 5. **Recommendations and Security**

- Use the environment variable `MCP_ALLOWED_DIRECTORIES` to restrict access only to the necessary folder.
- Share only the required volumes.
- Check the permissions of shared files.
- There is no need to open firewall ports or configure networks.

---

## 6. **Resources**

- [MCP Protocol](https://modelcontextprotocol.io)
- [Claude Desktop](https://www.anthropic.com/)
- [MCP Word Office repository on GitHub](https://github.com/LuiccianDev/mcp_word_office)

---

Do you have questions or need practical examples of Word manipulation?
Leave them in the Issues section of the repository.
