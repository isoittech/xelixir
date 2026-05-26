# xelixir: Excel Operations Tool

This repository provides an MCP server that exposes Excel manipulation tools to MCP clients
(such as Microsoft Copilot, Claude, or LibreChat) by wrapping Java-based tools from Python.

The core Excel operations are implemented in Java using Apache POI.
On top of that, we build the MCP server in Python, which gives us the following advantages:

- Apache POI is a very mature library for Excel operations, with rich features and good compatibility.
- It is easier to implement an MCP server in Python and to leverage libraries like `fastmcp`.

By adding a Python wrapper around the Java logic, we can:

- Let Java/Apache POI handle the heavy lifting for Excel processing.
- Keep the MCP server itself lightweight and easy to extend in Python.

## Overview & Goals

- Provide robust, scriptable Excel operations (read/write, formatting, charts, pivot tables, sheet management, etc.).
- Make these operations callable from modern MCP clients.
- Separate concerns so that the Java/Apache POI layer focuses on Excel semantics, while the Python layer focuses on MCP transport and ergonomics.

## Architecture & Tech Stack

- **Java core**

  - Excel operations are implemented in Java 11+ using Apache POI and related libraries.
  - Source code lives under [`java/src/xelixir`](java/src/xelixir/ExcelUtils.java:1).
  - The build script [`compile.sh`](java/tools/compile.sh:1) compiles `*.java` into a `dist` directory with the JARs in [`java/jars`](java/jars/poi-5.4.1.jar:1) on the classpath.

- **Python MCP server**

  - Requires Python 3.11 or later.
  - Main package is `xelixir`, defined in [`pyproject.toml`](mcp/pyproject.toml:1).
  - MCP entrypoint is [`xelixir.py`](mcp/xelixir.py:1), which:

    - registers tools with `FastMCP`,
    - exposes them over `stdio`, `http`, or `sse`,
    - optionally serves static files from a shared directory under `/files/...`.

  - Thin wrappers around the Java tools live in [`mcp/src/excel/wrapper.py`](mcp/src/excel/wrapper.py:1), and are re-exported via [`mcp/src/excel/__init__.py`](mcp/src/excel/__init__.py:1).

- **Transports**

  - `stdio` (typical for desktop MCP runtimes like Claude Desktop).
  - `http` (streamable HTTP).
  - `sse` (Server-Sent Events) for long-lived connections, often used from web UIs.

- **File sharing pattern**

  - Shared directory inside the container: `/mnt/data` by default, or the value of the `WORKSPACE_DIR` environment variable.
  - Public download URLs constructed as `<EXCEL_PUBLIC_BASE_URL>/files/<relative-path>` for files under the shared directory.

## Table of Contents

1. [Requirements](#requirements)
2. [Installation](#installation)
3. [Using the standalone CLI](#using-the-standalone-cli)
4. [Using with MCP clients](#using-with-mcp-clients)
5. [Using Docker](#using-docker)
6. [File sharing & download URLs](#file-sharing--download-urls)
7. [Available tools](#available-tools)
8. [Notes](#notes)
9. [Development & contributing](#development--contributing)
10. [License](#license)

## Requirements

- Java 11 or later
- Python 3.11 or later
- Excel file format: `.xlsx`
- Supported OS: Windows 10/11, macOS 12+, Linux (Ubuntu 20.04+)

## Installation

### 1. Clone the repository

```bash
git clone git@github.com:isoittech/xelixir.git
cd xelixir
```

### 2. Build the Java tools

You can compile the Java-side Excel tools using [`java/tools/compile.sh`](java/tools/compile.sh:1):

```bash
cd java
./tools/compile.sh
cd ..
```

This produces compiled classes under a `dist` directory, which are then used by the Python wrappers.

### 3. Set up the Python MCP server

The Python MCP server lives under the [`mcp/`](mcp) directory.
It is designed to be run locally using `uv`:

```bash
cd mcp
uv run xelixir --help
```

On the first run, dependencies defined in [`pyproject.toml`](mcp/pyproject.toml:1) will be installed automatically.
The main runtime dependency is `fastmcp>=0.3.0`; it brings in the HTTP/SSE stack used by the server.

For details that are specific to the MCP package layout and local development, see [`mcp/README.md`](mcp/README.md:1).

## Using the standalone CLI

In addition to the MCP server, this repository provides a standalone CLI package under [`cli/`](cli).
The CLI reuses the same Excel wrapper layer as the MCP server, so you can run Excel operations directly from a terminal without an MCP client.

The CLI package definition lives in [`cli/pyproject.toml`](cli/pyproject.toml:1), and the main command implementation lives in [`cli/src/xelixir_cli/tools.py`](cli/src/xelixir_cli/tools.py:1).

Run it from the repository root with `uv`:

```bash
uv run --project cli xelixir-tool --help
uv run --project cli xelixir-tool tool --help
```

Typical examples:

```bash
# Create a workbook
uv run --project cli xelixir-tool tool create-excel \
  --path ./example.xlsx \
  --sheet-name Data

# List sheets
uv run --project cli xelixir-tool tool list-sheets \
  --path ./example.xlsx

# Write a 2D array from inline JSON
uv run --project cli xelixir-tool tool write-excel \
  --path ./example.xlsx \
  --sheet-name Data \
  --data-json '[["A1","B1"],["A2","B2"]]'
```

The CLI prints JSON to stdout, matching the MCP tool response style as closely as possible.
For array-based inputs, you can pass either inline JSON or a JSON file path to options such as `--data-json` and `--rows-json`.

## Using with MCP clients

### Example: Claude Desktop

1. Point your MCP configuration to the `xelixir/mcp` directory.
2. The file [`mcp/mcp-config.json`](mcp/mcp-config.json:1) contains the definition for this MCP server.

Example `mcp-config.json`:

```json
{
  "mcpServers": {
    "xelixir": {
      "command": "uv",
      "args": [
        "run",
        "xelixir"
      ],
      "env": {
        "PYTHONPATH": "./src"
      }
    }
  }
}
```

After restarting your MCP client, you should see a server named `xelixir`
with a set of tools for working with Excel files.

## Using Docker

The [`mcp/`](mcp) directory also includes a [`Dockerfile`](mcp/Dockerfile:1) and [`docker-compose.yml`](mcp/docker-compose.yml:1).
You can use them if you prefer to connect via HTTP/SSE from your MCP client.

Build and run the container manually:

```bash
cd mcp
docker build -t xelixir .
docker run --rm -p 8585:8585 xelixir xelixir -t sse -p 8585
```

Or use docker compose:

```bash
cd mcp
docker compose up -d
```

From the MCP client side, configure an SSE endpoint such as
`http://host.docker.internal:8585/sse`.

## File sharing & download URLs

Some MCP runtimes (for example LibreChat) expose a shared directory inside the
tool container as `/mnt/data/...` for uploaded files and generated artifacts.

This server follows a common MCP pattern:

- **Excel files are read/written directly under a shared directory** inside the container
  (default `/mnt/data` or `WORKSPACE_DIR`).
- Tools that create or modify files return a **public `download_url`** pointing to
  `/files/...` so that the end user can download the result from the browser.

### Shared directory (`WORKSPACE_DIR` and `MCP_SHARED_DIR`)

- Container-side shared directory:

  - Default is `/mnt/data`.
  - Can be overridden by setting the `WORKSPACE_DIR` environment variable.

- In [`mcp/docker-compose.yml`](mcp/docker-compose.yml:1), the host directory is mounted into `/mnt/data`:

  ```yaml
  services:
    xelixir:
      volumes:
        - ${MCP_SHARED_DIR:-./workspace}:/mnt/data
  ```

  - Set `MCP_SHARED_DIR` on the host to control where uploaded/generated files
    are stored.
  - Inside the container, the server looks at `WORKSPACE_DIR` (or falls back
    to `/mnt/data`) when validating downloadable paths.

### Public base URL (`EXCEL_PUBLIC_BASE_URL`)

To let end users download generated files, the server needs to know the base URL
by which it is reached from the outside:

- Set `EXCEL_PUBLIC_BASE_URL` to something like:

  - `https://your-domain.example`, or
  - `http://your-host:8585`.

- When a tool writes to a file under the shared directory, it will attempt to
  build a public URL of the form:

  - `<EXCEL_PUBLIC_BASE_URL>/files/<relative-path-under-shared-dir>`.

- If `EXCEL_PUBLIC_BASE_URL` is not configured, tools that rely on it
  (such as `tool_create_excel`, `tool_write_excel`, `tool_write_range`,
  `tool_append_rows`, etc.) will raise an error instead of returning a
  container-only path.

This behavior is implemented in [`xelixir.py`](mcp/xelixir.py:76) via
`get_shared_directory()`, `get_public_base_url()`, and
`build_download_url_for_path()`.

### Editing the same file across multiple tool invocations

If your MCP client exposes files under paths like `/mnt/data/xxx.xlsx`,
you can:

- pass that same absolute path as the `path` argument to all tools, and
- keep editing the same workbook in place, without creating additional copies.

For example, tools such as `tool_write_excel`, `tool_write_range`, and
`tool_append_rows` will overwrite/update the existing file at the given path.
By contrast, `tool_create_excel` is intentionally implemented so that it **does
not overwrite an existing file**; it is meant for new workbook creation.

## Available tools

Below are some examples of the tools exposed to MCP clients.
The argument names shown here match the Python tool signatures:

- `tool_create_excel(path, sheet_name="Sheet1")`
- `tool_read_excel(path, sheet_name, range_str)`
- `tool_write_excel(path, sheet_name, data)`
- `tool_write_range(path, sheet_name, start_cell, data)`
- `tool_append_rows(path, sheet_name, rows, anchor_column="A")`
- `tool_create_sheet(path, sheet_name)`
- `tool_create_excel(path, sheet_name="Sheet1")`
- `tool_get_workbook_metadata(...)`
- `tool_rename_worksheet(...)`
- `tool_delete_worksheet(...)`
- `tool_copy_worksheet(...)`
- `tool_apply_formula(...)`
- `tool_validate_formula_syntax(...)`
- `tool_format_range(...)`
- `tool_merge_cells(...)`
- `tool_unmerge_cells(...)`
- `tool_copy_range(...)`
- `tool_delete_range(...)`
- `tool_validate_excel_range(...)`
- `tool_create_chart(...)`
- `tool_create_pivot_table(...)`
- `tool_list_sheets(path)`

The following JSON snippets show typical requests from an MCP client:

### Read from an Excel file

```json
{
  "server_name": "xelixir",
  "tool_name": "read_excel",
  "arguments": {
    "filePath": "/path/to/file.xlsx",
    "sheetName": "Sheet1",
    "range": "A1:C10"
  }
}
```

### Write to an Excel file

```json
{
  "server_name": "xelixir",
  "tool_name": "write_excel",
  "arguments": {
    "filePath": "/path/to/file.xlsx",
    "sheetName": "Sheet1",
    "data": [
      ["A1", "B1", "C1"],
      ["A2", "B2", "C2"]
    ]
  }
}
```

### Create a new sheet

```json
{
  "server_name": "xelixir",
  "tool_name": "create_sheet",
  "arguments": {
    "filePath": "/path/to/file.xlsx",
    "sheetName": "NewSheet"
  }
}
```

### Create a new Excel file

```json
{
  "server_name": "xelixir",
  "tool_name": "create_excel",
  "arguments": {
    "filePath": "/path/to/new_file.xlsx",
    "sheetName": "Sheet1"
  }
}
```

### Get workbook metadata

```json
{
  "server_name": "xelixir",
  "tool_name": "get_workbook_metadata",
  "arguments": {
    "filePath": "/path/to/file.xlsx",
    "includeRanges": false
  }
}
```

### Rename a worksheet

```json
{
  "server_name": "xelixir",
  "tool_name": "rename_worksheet",
  "arguments": {
    "filePath": "/path/to/file.xlsx",
    "oldName": "Sheet1",
    "newName": "NewName"
  }
}
```

### Delete a worksheet

```json
{
  "server_name": "xelixir",
  "tool_name": "delete_worksheet",
  "arguments": {
    "filePath": "/path/to/file.xlsx",
    "sheetName": "Sheet1"
  }
}
```

### Copy a worksheet

```json
{
  "server_name": "xelixir",
  "tool_name": "copy_worksheet",
  "arguments": {
    "filePath": "/path/to/file.xlsx",
    "sourceSheet": "Sheet1",
    "targetSheet": "Sheet1Copy"
  }
}
```

### Apply a formula to a cell

```json
{
  "server_name": "xelixir",
  "tool_name": "apply_formula",
  "arguments": {
    "filePath": "/path/to/file.xlsx",
    "sheetName": "Sheet1",
    "cell": "C1",
    "formula": "=SUM(A1:B1)"
  }
}
```

### Validate formula syntax

```json
{
  "server_name": "xelixir",
  "tool_name": "validate_formula_syntax",
  "arguments": {
    "filePath": "/path/to/file.xlsx",
    "sheetName": "Sheet1",
    "cell": "C1",
    "formula": "=SUM(A1:B1)"
  }
}
```

### Format a range

```json
{
  "server_name": "xelixir",
  "tool_name": "format_range",
  "arguments": {
    "filePath": "/path/to/file.xlsx",
    "sheetName": "Sheet1",
    "startCell": "A1",
    "endCell": "C3",
    "bold": true,
    "italic": false,
    "fontSize": 12,
    "fontColor": "#FF0000",
    "bgColor": "#FFFF00"
  }
}
```

### Merge cells

```json
{
  "server_name": "xelixir",
  "tool_name": "merge_cells",
  "arguments": {
    "filePath": "/path/to/file.xlsx",
    "sheetName": "Sheet1",
    "startCell": "A1",
    "endCell": "C1"
  }
}
```

### Unmerge cells

```json
{
  "server_name": "xelixir",
  "tool_name": "unmerge_cells",
  "arguments": {
    "filePath": "/path/to/file.xlsx",
    "sheetName": "Sheet1",
    "startCell": "A1",
    "endCell": "C1"
  }
}
```

### Copy a range

```json
{
  "server_name": "xelixir",
  "tool_name": "copy_range",
  "arguments": {
    "filePath": "/path/to/file.xlsx",
    "sheetName": "Sheet1",
    "sourceStart": "A1",
    "sourceEnd": "C3",
    "targetStart": "D1",
    "targetSheet": "Sheet2"
  }
}
```

### Delete a range

```json
{
  "server_name": "xelixir",
  "tool_name": "delete_range",
  "arguments": {
    "filePath": "/path/to/file.xlsx",
    "sheetName": "Sheet1",
    "startCell": "A1",
    "endCell": "C3",
    "shiftDirection": "up"
  }
}
```

### Validate an Excel range

```json
{
  "server_name": "xelixir",
  "tool_name": "validate_excel_range",
  "arguments": {
    "filePath": "/path/to/file.xlsx",
    "sheetName": "Sheet1",
    "startCell": "A1",
    "endCell": "C3"
  }
}
```

### Create a chart

```json
{
  "server_name": "xelixir",
  "tool_name": "create_chart",
  "arguments": {
    "filePath": "/path/to/file.xlsx",
    "sheetName": "Sheet1",
    "dataRange": "A1:C10",
    "chartType": "column",
    "targetCell": "E1",
    "title": "Sample Chart",
    "xAxis": "X Axis",
    "yAxis": "Y Axis"
  }
}
```

### Create a pivot table

```json
{
  "server_name": "xelixir",
  "tool_name": "create_pivot_table",
  "arguments": {
    "filePath": "/path/to/file.xlsx",
    "sheetName": "Sheet1",
    "dataRange": "A1:D100",
    "rows": ["Category"],
    "values": ["Sales"],
    "columns": ["Region"],
    "aggFunc": "sum"
  }
}
```

### List sheet names

```json
{
  "server_name": "xelixir",
  "tool_name": "list_sheets",
  "arguments": {
    "filePath": "/path/to/book.xlsx"
  }
}
```

The `list_sheets` tool returns:

```json
{
  "path": "/path/to/book.xlsx",
  "sheets": ["Sheet1", "Data", "Summary"]
}
```

## Notes

- Always use absolute paths for file paths.
- If you omit the sheet name (in tools that permit omission), the first sheet will be used by default.
- Ranges should be specified in A1 notation such as `"A1:C10"`.
- `create_excel` will fail if the target file already exists.
- Depending on the current implementation, pivot table support may only build metadata
  and may not create a full Excel pivot table object in the file.

## Development & contributing

There is currently no dedicated CONTRIBUTING document.
When changing or extending this project, please:

- keep the Java tools buildable via [`java/tools/compile.sh`](java/tools/compile.sh:1),
- keep the Python wrappers in [`mcp/src/excel`](mcp/src/excel/__init__.py:1) and
  the MCP entrypoint [`mcp/xelixir.py`](mcp/xelixir.py:1) in sync,
- update both [`README.md`](README.md:1) and [`README.JA.md`](README.JA.md:1) when you add
  or change user-facing behavior,
- describe your changes in English or Japanese in commit messages and PR descriptions.

Pull requests and issue reports are welcome.

## License

This repository includes an Apache License 2.0 text in [`LICENSE`](LICENSE:1).
Unless otherwise noted in individual files, the project is intended to be used
under those terms.

> Note: the Python package metadata in [`mcp/pyproject.toml`](mcp/pyproject.toml:7)
> currently declares an MIT license string. This will be aligned with the
> repository-level Apache License 2.0 in a future revision; for now, treat
> the root `LICENSE` file as canonical.
