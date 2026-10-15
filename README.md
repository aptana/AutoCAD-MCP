# AutoCAD MCP Server

A Windows-only Model Context Protocol (MCP) server that lets an MCP client work with the active AutoCAD drawing through the AutoCAD COM automation interface.

## Requirements

- Windows
- Python 3.12 or newer
- AutoCAD 2021 (the server currently connects to the COM ProgID `AutoCAD.Application.24`)
- [uv](https://docs.astral.sh/uv/getting-started/installation/)
- An MCP-compatible client, such as Visual Studio Code with MCP support

AutoCAD must be installed on the same Windows machine as this server. Keep AutoCAD open when using the tools. Most drawing tools also require an active drawing.

## Installation

1. Clone this repository:

   ```powershell
   git clone https://github.com/aptana/AutoCAD-MCP.git
   cd AutoCAD-MCP
   ```

2. Install the project dependencies:

   ```powershell
   uv sync
   ```

   This project uses `mcp` and `pywin32`; `uv sync` creates the virtual environment and installs the dependencies from `pyproject.toml`.

3. Start AutoCAD and open a drawing. The server attaches to the running AutoCAD application through COM; it does not launch AutoCAD itself.

## Configure Visual Studio Code

Create or update `.vscode/mcp.json` in your workspace. Replace the directory below with the absolute path to your local clone:

```json
{
  "servers": {
    "autocad": {
      "type": "stdio",
      "command": "uv",
      "args": [
        "--directory",
        "C:/path/to/AutoCAD-MCP",
        "run",
        "autocad_mcp_server.py"
      ]
    }
  }
}
```

If `uv` is not on your `PATH`, set `command` to the full path of `uv.exe`. Start or restart the `autocad` MCP server from the MCP controls in VS Code after saving the configuration.

## Run Manually

From the repository directory, run:

```powershell
uv run autocad_mcp_server.py
```

The server uses standard input/output (stdio) for MCP communication. When launched by an MCP client, do not type into its terminal process; interact with it through the client.

## Available Tools

### Drawings and views

- `get_active_document`, `list_documents`, `activate_document`
- `open_document`, `new_document`, `save_document`, `close_document`
- `set_view`
- `set_active_layout`, `list_layouts`, `plot_active_layout_to_pdf`

### Geometry and annotations

- `create_layer`, `create_line`, `create_circle`, `create_rectangle`, `create_polyline`
- `add_text`, `create_aligned_dimension`, `create_linear_dimension`
- `insert_block`, `insert_external_block`, `arrange_blocks_with_left_connector`
- `attach_xref`, `import_geojson`

### Inspect and edit entities

- `list_blocks`, `list_modelspace_entities`
- `get_entity_bounding_box`, `get_entity_properties`, `set_entity_properties`
- `transform_entities`, `erase_entity`
- `send_autocad_command`

## Coordinate and File Notes

- Geometry tools use World Coordinate System (WCS) coordinates unless their descriptions state otherwise.
- File paths passed to tools, including drawing, block, Xref, PDF, and GeoJSON paths, must be accessible to the Windows machine running AutoCAD and the MCP server.
- `import_geojson` accepts a GeoJSON `FeatureCollection` containing point, line, or polygon features. Feature properties can optionally supply layer, block, and label values.
- AutoCAD COM availability and behavior depend on the installed AutoCAD version. If using a version other than AutoCAD 2021, verify its COM ProgID and update `get_autocad()` in `autocad_mcp_server.py` if needed.

## Safety

Some tools modify drawings, save files, close documents, erase entities, send native AutoCAD commands, or create output files. Review the requested operation and active drawing before invoking these tools. Save a backup of important drawings before making substantial changes.

## Troubleshooting

- **AutoCAD is not found:** Open AutoCAD first. This server currently requests `AutoCAD.Application.24` and reports an error if it cannot attach.
- **No active drawing:** Open a drawing in AutoCAD before calling tools that operate on the active document.
- **Dependency or import errors:** Run `uv sync` from the repository directory, then restart the MCP server.
- **MCP server does not start in VS Code:** Confirm that `uv` is available to the VS Code process, the configured directory points to this clone, and the JSON in `.vscode/mcp.json` is valid. Restart the server after changing the configuration.