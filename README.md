# AutoCAD MCP Server

A [Model Context Protocol](https://modelcontextprotocol.io) server that lets
Claude (or any MCP client) draw and control **AutoCAD** on Windows through the
AutoCAD ActiveX/COM automation interface.

```
MCP client  <-- stdio JSON-RPC -->  server.py  <-- COM -->  AutoCAD
```

## Requirements

- Windows
- AutoCAD 2021 or newer, installed and licensed (the COM object
  `AutoCAD.Application` must be registered — it is, after a normal install)
- Python 3.10+
- [`uv`](https://docs.astral.sh/uv/)

## Install

```powershell
# 1. Install uv (then restart your terminal)
powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"

# 2. Clone and enter the repo
git clone https://github.com/YOUR_USERNAME/autocad-mcp.git
cd autocad-mcp

# 3. Create the environment and install dependencies
uv venv
.venv\Scripts\activate
uv pip install -e .

# 4. Register pywin32's COM DLLs (one time, required for AutoCAD COM)
python .venv\Scripts\pywin32_postinstall.py -install
```

## Connect to Claude Desktop

Open your config file (create it if missing):

```powershell
code $env:AppData\Claude\claude_desktop_config.json
```

Add the server, using the **absolute path** to where you cloned the repo
(double backslashes in JSON):

```json
{
  "mcpServers": {
    "autocad": {
      "command": "uv",
      "args": [
        "--directory",
        "C:\\Users\\YOUR_USERNAME\\autocad-mcp",
        "run",
        "server.py"
      ]
    }
  }
}
```

If Claude can't find `uv`, run `where uv` in PowerShell and use that full path
in the `command` field. Then **fully quit** Claude Desktop (right-click the
tray icon → Quit) and reopen it.

## Tools

| Tool | What it does |
| --- | --- |
| `connect` | Attach to AutoCAD (launch if needed), report active drawing |
| `draw_line` | Line from (x1,y1) to (x2,y2) |
| `draw_circle` | Circle at a center with a radius |
| `draw_rectangle` | Closed rectangle between two corners |
| `draw_polyline` | Polyline through a list of points |
| `add_text` | Single-line text at a point |
| `create_layer` | New layer with an AutoCAD color index |
| `set_active_layer` | Switch the active layer |
| `list_layers` | List all layers |
| `insert_block` | Insert a block reference |
| `zoom_extents` | Zoom to fit all objects |
| `save_drawing` | Save, or Save As to a path |

Add your own tools by writing a function and decorating it with `@mcp.tool()`.
The SDK turns the type hints and docstring into the schema Claude sees.

## Notes

- This is a **stdio** server, so the code never uses `print()` — that would
  corrupt the JSON-RPC stream. Logging goes to stderr.
- AutoCAD COM only works on Windows. The code imports `win32com` lazily so the
  file can still be read/linted elsewhere.

## License

MIT
