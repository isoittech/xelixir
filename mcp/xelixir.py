#!/usr/bin/env python
"""MCP server for Excel operations backed by Java tools.

This server exposes Excel manipulation capabilities implemented in Java
(classes under ``java/src/xelixir``) via thin Python wrappers
(``mcp/src/excel``).

In addition to MCP (SSE/stdio), this server can expose a static file download
endpoint when running in SSE mode.

This follows a common MCP pattern: **tools return a downloadable URL** for
generated artifacts, instead of returning a local/container-only file path.

- SSE MCP endpoint: ``/sse``
- Static downloads: ``/files/<relative-path>``

Related OSS pattern reference:
- ``temp-file-share-mcp`` uploads a local file to an external service
  (tmpfile.link) and returns a shareable download URL.
- This server returns the same kind of shareable URL, but hosts the file on the
  MCP server itself under ``/files/...`` (rather than using an external service).
"""

from __future__ import annotations

import argparse
import os
import sys
from pathlib import Path
from typing import Any, Dict
from urllib.parse import quote

from mcp.server.fastmcp import FastMCP

import uvicorn
from starlette.applications import Starlette
from starlette.routing import Mount
from starlette.staticfiles import StaticFiles

from excel import (
    create_excel,
    read_excel,
    write_excel,
    write_range,
    append_rows,
    create_sheet,
    rename_worksheet,
    delete_worksheet,
    copy_worksheet,
    apply_formula,
    validate_formula_syntax,
    format_range,
    merge_cells,
    unmerge_cells,
    copy_range,
    delete_range,
    validate_excel_range,
    create_chart,
    create_pivot_table,
    list_sheets,
)

# NOTE:
# FastMCP may enable extra localhost-only host-header protections when bound to
# localhost. In container deployments (e.g., accessed via host.docker.internal),
# that can cause "Invalid Host header" responses. We bind to 0.0.0.0 to avoid
# those localhost-only restrictions.
app = FastMCP(name="xelixir", host="0.0.0.0")


# ---------------------------------------------------------------------------
# File sharing & download URL helpers
# ---------------------------------------------------------------------------


DEFAULT_SHARED_DIR = "/mnt/data"
# Intentionally empty by default: deployments must set EXCEL_PUBLIC_BASE_URL
# to a user-reachable origin (reverse proxy / domain / host:port).
DEFAULT_PUBLIC_BASE_URL = ""


def get_shared_directory() -> str:
    """Return the directory that should be exposed under `/files`.

    Priority:
        1) WORKSPACE_DIR env var (container path)
        2) DEFAULT_SHARED_DIR

    The directory is created if it does not exist.
    """
    shared_dir = os.environ.get("WORKSPACE_DIR", DEFAULT_SHARED_DIR)
    os.makedirs(shared_dir, exist_ok=True)
    return shared_dir


def get_public_base_url() -> str:
    """Return the public base URL used to build file download URLs.

    Priority:
        1) EXCEL_PUBLIC_BASE_URL env var
        2) DEFAULT_PUBLIC_BASE_URL (empty by default)

    Trailing slash is stripped for normalization.

    Notes:
        This should be set to a user-reachable origin, e.g.:
        - https://your-domain.example
        - http://your-host-or-proxy:8585
    """
    return os.environ.get("EXCEL_PUBLIC_BASE_URL", DEFAULT_PUBLIC_BASE_URL).rstrip("/")


def build_download_url_for_path(file_path: str) -> str:
    """Build a public download URL for a file under the shared directory.

    This maps:
        <shared_dir>/<relpath>  ->  <public_base_url>/files/<relpath>

    This function is intentionally strict: if the path is not under the shared
    directory, or if the base URL is not configured, it raises an error.
    The goal is to avoid returning container-only paths that end users cannot
    download.

    Raises:
        ValueError: when `file_path` is not under WORKSPACE_DIR.
        RuntimeError: when EXCEL_PUBLIC_BASE_URL is empty/unset after normalization.
    """
    base_url = get_public_base_url()
    if not base_url:
        raise RuntimeError(
            "EXCEL_PUBLIC_BASE_URL is not configured. Set it to the public base URL for this MCP server "
            "(e.g. https://your-domain.example or http://your-host-or-proxy:8585)."
        )

    shared_dir = Path(get_shared_directory()).resolve()
    p = Path(file_path).resolve()

    try:
        rel = p.relative_to(shared_dir)
    except ValueError as e:
        raise ValueError(
            f"File path must be under the shared directory to be downloadable. "
            f"WORKSPACE_DIR={shared_dir} but got path={p}"
        ) from e

    # Use URL-safe POSIX path
    rel_posix = rel.as_posix()
    rel_quoted = quote(rel_posix)
    return f"{base_url}/files/{rel_quoted}"


def build_asgi_app_for_sse(sse_app) -> Starlette:
    """Wrap the FastMCP SSE ASGI app and expose `/files` for downloads."""
    shared_dir = get_shared_directory()
    static_files_app = StaticFiles(directory=shared_dir, check_dir=False)

    return Starlette(
        routes=[
            Mount("/files", app=static_files_app, name="files"),
            Mount("/", app=sse_app),
        ]
    )


# ---------------------------------------------------------------------------
# Tool registrations
# ---------------------------------------------------------------------------


@app.tool()
async def tool_create_excel(path: str, sheet_name: str = "Sheet1") -> Dict[str, Any]:
    """Create a new Excel workbook (.xlsx) at the given path (fails if it already exists).

    In chat UIs, users often cannot access container paths directly. If the created file is
    under the shared directory (see WORKSPACE_DIR), this tool also returns a public
    `download_url` that points to `/files/...` on this server.

    This "return a shareable URL for an artifact" pattern is common in MCP servers.
    For example, `temp-file-share-mcp` uploads a local file to an external service
    and returns a shareable download URL. This server uses the same idea, but serves
    the file from the MCP server host itself (under `/files/...`) instead of uploading
    it to an external file-sharing service.

    Args:
        path:
            Destination file path to create (e.g. `/mnt/data/report.xlsx`).
            When running in Docker, `/mnt/data` is commonly mounted as a shared directory.
        sheet_name:
            Name of the initial worksheet (default: `Sheet1`).

    Returns:
        A JSON-serializable dict containing:
        - `message`: confirmation text
        - `path`: the file path that was created
        - `download_url`: public URL (requires EXCEL_PUBLIC_BASE_URL; path must be under WORKSPACE_DIR)

    Raises:
        RuntimeError: if EXCEL_PUBLIC_BASE_URL is not configured.
        ValueError: if `path` is not under WORKSPACE_DIR.

    Notes:
        - This tool does not overwrite existing files.
        - It will fail if the parent directory does not exist or is not writable.
    """

    create_excel(path, sheet_name)
    return {
        "message": f"Created Excel workbook at {path} with sheet '{sheet_name}'",
        "path": path,
        "download_url": build_download_url_for_path(path),
    }


@app.tool()
async def tool_read_excel(path: str, sheet_name: str, range_str: str) -> Dict[str, Any]:
    """Read a rectangular cell range from a worksheet and return it as a 2D array.

    Args:
        path:
            Target workbook path (e.g. `/mnt/data/book.xlsx`).
        sheet_name:
            Worksheet name (required).
        range_str:
            A1-style rectangular range (e.g. `A1:C10`).

    Returns:
        A JSON-serializable dict containing:
        - `data`: 2D array (rows x columns). Values are JSON-compatible types.
        - `path`, `range`, and other metadata.

    Examples:
        - path="/mnt/data/book.xlsx", sheet_name="Sheet1", range_str="A1:C10"

    Notes:
        - Raises an error if the file/sheet does not exist or the range is invalid.
    """

    data = read_excel(path, sheet_name, range_str)
    return {"path": path, "+sheet": sheet_name, "range": range_str, "data": data}


@app.tool()
async def tool_write_excel(path: str, sheet_name: str, data: Any) -> Dict[str, Any]:
    """Write a 2D array into a worksheet starting at A1 (fixed start cell).

    If the workbook is under EXCEL_SHARED_DIR, this also returns a public `download_url`
    for user-facing downloads.

    This "return a shareable URL for an artifact" pattern is common in MCP servers.
    For example, `temp-file-share-mcp` uploads a local file to an external service
    and returns a shareable download URL. This server returns a similar shareable URL,
    but hosts the file on the MCP server itself under `/files/...`.

    Args:
        path:
            Target workbook path.
        sheet_name:
            Worksheet name.
        data:
            A 2D array to write (e.g. `[[\"A1\", \"B1\"], [\"A2\", \"B2\"]]`).
            The payload is passed to the Java layer as JSON, so it must be JSON-serializable.

    Returns:
        A JSON-serializable dict containing:
        - `message`: confirmation text
        - `path`: the workbook path
        - `download_url`: public URL (requires EXCEL_PUBLIC_BASE_URL; path must be under EXCEL_SHARED_DIR)

    Raises:
        RuntimeError: if EXCEL_PUBLIC_BASE_URL is not configured.
        ValueError: if `path` is not under EXCEL_SHARED_DIR.
    """

    write_excel(path, sheet_name, data)
    return {
        "message": f"Wrote {len(data)} rows to {path}:{sheet_name}!A1",
        "path": path,
        "download_url": build_download_url_for_path(path),
    }


@app.tool()
async def tool_write_range(path: str, sheet_name: str, start_cell: str, data: Any) -> Dict[str, Any]:
    """Write a 2D array into a worksheet starting at `start_cell`.

    This is like `write_excel` but allows specifying the top-left cell.

    If the workbook is under EXCEL_SHARED_DIR, this also returns a public `download_url`
    for user-facing downloads.

    Args:
        path: Target workbook path.
        sheet_name: Worksheet name.
        start_cell: A1-style top-left cell address (e.g. `B3`).
        data: A 2D array to write (JSON-serializable). `null` writes blanks.

    Returns:
        A JSON-serializable dict containing:
        - `message`
        - `path`
        - `start_cell`
        - `download_url`
    """
    write_range(path, sheet_name, start_cell, data)
    return {
        "message": f"Wrote {len(data)} rows to {path}:{sheet_name}!{start_cell}",
        "path": path,
        "start_cell": start_cell,
        "download_url": build_download_url_for_path(path),
    }


@app.tool()
async def tool_append_rows(
    path: str,
    sheet_name: str,
    rows: Any,
    anchor_column: str = "A",
) -> Dict[str, Any]:
    """Append rows at the first empty row determined by scanning `anchor_column`.

    The anchor column is scanned top-to-bottom; the first row where the anchor cell is
    missing/blank is considered the append position.

    Args:
        path: Target workbook path.
        sheet_name: Worksheet name.
        rows: 2D array (rows) to append (JSON-serializable).
        anchor_column: Column letter used to determine the first empty row (default: "A").

    Returns:
        A JSON-serializable dict containing:
        - `message`
        - `path`
        - `anchor_column`
        - `start_row` (0-based)
        - `download_url`
    """
    start_row = append_rows(path, sheet_name, rows, anchor_column=anchor_column)
    return {
        "message": f"Appended {len(rows)} rows to {path}:{sheet_name} at row {start_row + 1} (anchor={anchor_column})",
        "path": path,
        "anchor_column": anchor_column,
        "start_row": start_row,
        "download_url": build_download_url_for_path(path),
    }


@app.tool()
async def tool_create_sheet(path: str, sheet_name: str) -> Dict[str, Any]:
    """Add a new worksheet to an existing workbook.

    If the workbook is under EXCEL_SHARED_DIR, this also returns a public `download_url`.

    Args:
        path: Target workbook path.
        sheet_name: New worksheet name (fails if a sheet with the same name already exists).

    Returns:
        A JSON-serializable dict containing `message`, `path`, and `download_url`.
    """
    create_sheet(path, sheet_name)
    return {
        "message": f"Created sheet '{sheet_name}' in {path}",
        "path": path,
        "download_url": build_download_url_for_path(path),
    }


@app.tool()
async def tool_rename_worksheet(path: str, old_name: str, new_name: str) -> Dict[str, Any]:
    """Rename a worksheet.

    If the workbook is under EXCEL_SHARED_DIR, this also returns a public `download_url`.

    Args:
        path: Target workbook path.
        old_name: Current worksheet name.
        new_name: New worksheet name (fails if it collides with an existing sheet).

    Returns:
        A JSON-serializable dict containing `message`, `path`, and `download_url`.
    """
    rename_worksheet(path, old_name, new_name)
    return {
        "message": f"Renamed sheet '{old_name}' to '{new_name}' in {path}",
        "path": path,
        "download_url": build_download_url_for_path(path),
    }


@app.tool()
async def tool_delete_worksheet(path: str, sheet_name: str) -> Dict[str, Any]:
    """Delete a worksheet by name.

    If the workbook is under EXCEL_SHARED_DIR, this also returns a public `download_url`.

    Args:
        path: Target workbook path.
        sheet_name: Worksheet name to delete.

    Returns:
        A JSON-serializable dict containing `message`, `path`, and `download_url`.

    Notes:
        - Raises an error if the worksheet does not exist.
        - Some workbooks may disallow deleting the last remaining sheet.
    """
    delete_worksheet(path, sheet_name)
    return {
        "message": f"Deleted sheet '{sheet_name}' in {path}",
        "path": path,
        "download_url": build_download_url_for_path(path),
    }


@app.tool()
async def tool_copy_worksheet(path: str, source_sheet: str, target_sheet: str) -> Dict[str, Any]:
    """Copy a worksheet (creates a new duplicated sheet).

    If the workbook is under EXCEL_SHARED_DIR, this also returns a public `download_url`.

    Args:
        path: Target workbook path.
        source_sheet: Source worksheet name.
        target_sheet: Target (new) worksheet name.

    Returns:
        A JSON-serializable dict containing `message`, `path`, and `download_url`.
    """
    copy_worksheet(path, source_sheet, target_sheet)
    return {
        "message": f"Copied sheet '{source_sheet}' to '{target_sheet}' in {path}",
        "path": path,
        "download_url": build_download_url_for_path(path),
    }


@app.tool()
async def tool_apply_formula(path: str, sheet_name: str, cell: str, formula: str) -> Dict[str, Any]:
    """Set an Excel formula on a specific cell.

    If the workbook is under EXCEL_SHARED_DIR, this also returns a public `download_url`.

    Args:
        path: Target workbook path.
        sheet_name: Worksheet name.
        cell: A1-style cell address (e.g. `C1`).
        formula:
            Excel-style formula string (e.g. `=SUM(A1:B1)`).
            Include the leading `=`.

    Returns:
        A JSON-serializable dict containing `message`, `path`, and `download_url`.
    """
    apply_formula(path, sheet_name, cell, formula)
    return {
        "message": f"Applied formula '{formula}' to {path}:{sheet_name}!{cell}",
        "path": path,
        "download_url": build_download_url_for_path(path),
    }


@app.tool()
async def tool_validate_formula_syntax(path: str, sheet_name: str, formula: str) -> Dict[str, Any]:
    """Validate the syntax of an Excel formula (does not modify the workbook).

    Args:
        path: Target workbook path.
        sheet_name: Worksheet name (used as validation context).
        formula: Formula to validate (e.g. `=SUM(A1:B1)`).

    Returns:
        A JSON-serializable dict:
        - `valid`: boolean (true if syntactically valid)
        - `formula`: the input formula

    Notes:
        - This is primarily a syntax check; it does not guarantee referenced cells exist or that the formula is semantically valid.
    """
    ok = validate_formula_syntax(path, sheet_name, formula)
    return {"valid": ok, "formula": formula}


@app.tool()
async def tool_format_range(
    path: str,
    sheet_name: str,
    start_cell: str,
    end_cell: str,
    bold: bool = False,
    italic: bool = False,
    font_size: int = 0,
    font_color: str = "",
    bg_color: str = "",
) -> Dict[str, Any]:
    """Apply basic formatting to a rectangular cell range.

    If the workbook is under EXCEL_SHARED_DIR, this also returns a public `download_url`.

    Args:
        path: Target workbook path.
        sheet_name: Worksheet name.
        start_cell: Range start cell (A1 style).
        end_cell: Range end cell (A1 style).
        bold: Set true for bold.
        italic: Set true for italic.
        font_size: Font size (0 means "unspecified").
        font_color: Font color in hex (e.g. `#FF0000`). Empty string means "unspecified".
        bg_color: Background color in hex (e.g. `#FFFF00`). Empty string means "unspecified".

    Returns:
        A JSON-serializable dict containing `message`, `path`, and `download_url`.
    """
    format_range(path, sheet_name, start_cell, end_cell, bold, italic, font_size, font_color, bg_color)
    return {
        "message": f"Formatted range {sheet_name}!{start_cell}:{end_cell} in {path}",
        "path": path,
        "download_url": build_download_url_for_path(path),
    }


@app.tool()
async def tool_merge_cells(path: str, sheet_name: str, start_cell: str, end_cell: str) -> Dict[str, Any]:
    """Merge cells in the specified rectangular range.

    If the workbook is under EXCEL_SHARED_DIR, this also returns a public `download_url`.

    Args:
        path: Target workbook path.
        sheet_name: Worksheet name.
        start_cell: Merge range start cell (A1 style).
        end_cell: Merge range end cell (A1 style).

    Returns:
        A JSON-serializable dict containing `message`, `path`, and `download_url`.
    """
    merge_cells(path, sheet_name, start_cell, end_cell)
    return {
        "message": f"Merged cells {sheet_name}!{start_cell}:{end_cell} in {path}",
        "path": path,
        "download_url": build_download_url_for_path(path),
    }


@app.tool()
async def tool_unmerge_cells(path: str, sheet_name: str, start_cell: str, end_cell: str) -> Dict[str, Any]:
    """Unmerge cells in the specified rectangular range.

    If the workbook is under EXCEL_SHARED_DIR, this also returns a public `download_url`.

    Args:
        path: Target workbook path.
        sheet_name: Worksheet name.
        start_cell: Unmerge range start cell (A1 style).
        end_cell: Unmerge range end cell (A1 style).

    Returns:
        A JSON-serializable dict containing `message`, `path`, and `download_url`.
    """
    unmerge_cells(path, sheet_name, start_cell, end_cell)
    return {
        "message": f"Unmerged cells {sheet_name}!{start_cell}:{end_cell} in {path}",
        "path": path,
        "download_url": build_download_url_for_path(path),
    }


@app.tool()
async def tool_copy_range(
    path: str,
    sheet_name: str,
    source_start: str,
    source_end: str,
    target_start: str,
    target_sheet: str | None = None,
    copy_style: bool = True,
) -> Dict[str, Any]:
    """Copy a cell range to a target location.

    If the workbook is under EXCEL_SHARED_DIR, this also returns a public `download_url`.

    Args:
        path: Target workbook path.
        sheet_name: Source worksheet name.
        source_start: Source range start cell (A1 style).
        source_end: Source range end cell (A1 style).
        target_start: Target start cell (A1 style).
        target_sheet:
            Target worksheet name (optional). If omitted, uses the source worksheet.
        copy_style:
            When true (default), copy cell styles (borders/fills/font/alignment) in addition to values.
            When false, copy values only.

    Returns:
        A JSON-serializable dict containing `message`, `path`, and `download_url`.
    """
    copy_range(path, sheet_name, source_start, source_end, target_start, target_sheet, copy_style=copy_style)
    return {
        "message": (
            f"Copied range {sheet_name}!{source_start}:{source_end} "
            f"to {target_sheet or sheet_name}!{target_start} in {path} (copy_style={copy_style})"
        ),
        "path": path,
        "download_url": build_download_url_for_path(path),
    }


@app.tool()
async def tool_delete_range(
    path: str,
    sheet_name: str,
    start_cell: str,
    end_cell: str,
    shift_direction: str = "up",
) -> Dict[str, Any]:
    """Delete a cell range and shift surrounding cells.

    If the workbook is under EXCEL_SHARED_DIR, this also returns a public `download_url`.

    Args:
        path: Target workbook path.
        sheet_name: Worksheet name.
        start_cell: Delete range start cell (A1 style).
        end_cell: Delete range end cell (A1 style).
        shift_direction:
            Direction to shift cells after deletion: `up` or `left` (default: `up`).

    Returns:
        A JSON-serializable dict containing `message`, `path`, and `download_url`.
    """
    delete_range(path, sheet_name, start_cell, end_cell, shift_direction)
    return {
        "message": f"Deleted range {sheet_name}!{start_cell}:{end_cell} in {path} (shift={shift_direction})",
        "path": path,
        "download_url": build_download_url_for_path(path),
    }


@app.tool()
async def tool_validate_excel_range(
    path: str,
    sheet_name: str,
    start_cell: str,
    end_cell: str | None = None,
) -> Dict[str, Any]:
    """Validate a cell address/range specification (does not modify the workbook).

    Args:
        path: Target workbook path.
        sheet_name: Worksheet name.
        start_cell: Start cell (A1 style).
        end_cell: End cell (A1 style, optional).

    Returns:
        A JSON-serializable dict:
        - `valid`: boolean
        - `sheet`, `start`, `end`: echo of the inputs
    """
    ok = validate_excel_range(path, sheet_name, start_cell, end_cell)
    return {"valid": ok, "sheet": sheet_name, "start": start_cell, "end": end_cell}


@app.tool()
async def tool_create_chart(
    path: str,
    sheet_name: str,
    data_range: str,
    chart_type: str,
    target_cell: str,
    title: str | None = None,
    x_axis: str | None = None,
    y_axis: str | None = None,
) -> Dict[str, Any]:
    """Create a chart from a data range and place it on the worksheet.

    If the workbook is under EXCEL_SHARED_DIR, this also returns a public `download_url`.

    Args:
        path: Target workbook path.
        sheet_name: Worksheet name (both data source and placement sheet).
        data_range: Data range (e.g. `A1:C10`).
        chart_type:
            Chart type (e.g. `column`, `line`, `bar`, `area`, `scatter`, `pie`).
        target_cell: Top-left cell where the chart will be placed (e.g. `E1`).
        title: Chart title (optional).
        x_axis: X-axis label (optional).
        y_axis: Y-axis label (optional).

    Returns:
        A JSON-serializable dict containing `message`, `path`, and `download_url`.

    Notes:
        - Exact chart appearance can vary by Excel compatibility and defaults/templates.
    """
    create_chart(path, sheet_name, data_range, chart_type, target_cell, title, x_axis, y_axis)
    return {
        "message": f"Created chart of type '{chart_type}' at {sheet_name}!{target_cell} in {path}",
        "path": path,
        "download_url": build_download_url_for_path(path),
    }


@app.tool()
async def tool_list_sheets(path: str) -> Dict[str, Any]:
    """List worksheet names in a workbook.

    Args:
        path: Target workbook path (e.g. `/mnt/data/book.xlsx`).

    Returns:
        A JSON-serializable dict:
        - `path`: echo of the workbook path
        - `sheets`: list of worksheet names in order
    """

    sheets = list_sheets(path)
    return {"path": path, "sheets": sheets}


@app.tool()
async def tool_create_pivot_table(
    path: str,
    sheet_name: str,
    data_range: str,
    rows: list[str],
    values: list[str],
    columns: list[str] | None = None,
    agg_func: str = "sum",
) -> Dict[str, Any]:
    """Create a pivot-table instruction from a source data range.

    Args:
        path: Target workbook path.
        sheet_name: Worksheet name.
        data_range: Source data range (e.g. `A1:D100`).
        rows: Column names to use as row labels (e.g. `["Category"]`).
        values: Column names to aggregate (e.g. `["Sales"]`).
        columns: Column names to use as column labels (optional).
        agg_func:
            Aggregation function (e.g. `sum`, `count`, `average`, `max`, `min`).

    Returns:
        A JSON-serializable dict: `{ "message": "..." }`.

    Notes:
        - Depending on the implementation state, this may build metadata only and may not create a native Excel pivot table object.
    """
    message = create_pivot_table(path, sheet_name, data_range, rows, values, columns, agg_func)
    return {"message": message}


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------


# ---- Main Function ----
def main(transport: str = "stdio", port: int = 8000) -> None:
    """Entry point used by the ``xelixir`` console script.

    When running in SSE mode, this entry point wraps FastMCP's SSE ASGI app with Starlette
    and exposes:
        - `/sse` and `/messages` (MCP)
        - `/files/...` (static downloads from EXCEL_SHARED_DIR)

    This enables tools to return user-downloadable URLs instead of container file paths.
    """

    # setuptools console_scripts does NOT pass CLI arguments to this function.
    # When called via the entry-point wrapper (e.g. ``uv run xelixir -t sse``),
    # sys.argv contains the real CLI args, so we parse it ourselves here.
    # When called programmatically (``from xelixir import main; main(...)``),
    # the caller-supplied values are respected because they become the defaults
    # that argparse falls back to when the corresponding option is absent.
    if len(sys.argv) > 1:
        parser = argparse.ArgumentParser(description="MCP Server for Excel manipulation")
        parser.add_argument(
            "-t",
            "--transport",
            type=str,
            default=transport,
            choices=["stdio", "http", "sse"],
            help="Transport method for the MCP server (default: stdio)",
        )
        parser.add_argument(
            "-p",
            "--port",
            type=int,
            default=port,
            help="Port to run the MCP server on (default: 8000)",
        )
        args = parser.parse_args()
        transport = args.transport
        port = args.port

    if transport == "http":
        import asyncio

        # Bind on all interfaces in containers and respect the requested port
        app.settings.host = "0.0.0.0"
        app.settings.port = port
        try:
            app.run(transport="streamable-http")
        except asyncio.exceptions.CancelledError:
            print("Server stopped by user.")
        except KeyboardInterrupt:
            print("Server stopped by user.")
        except Exception as e:  # pragma: no cover - defensive logging
            print(f"Error starting server: {e}")

    elif transport == "sse":
        # Prefer ASGI + uvicorn so we can also serve /files
        try:
            sse_app = app.sse_app()
        except AttributeError:
            # Fallback: run FastMCP directly (no /files endpoint)
            print("FastMCP.sse_app not available; falling back to app.run('sse')", file=sys.stderr)
            app.settings.host = "0.0.0.0"
            app.settings.port = port
            app.run(transport="sse")
            return

        asgi_app = build_asgi_app_for_sse(sse_app)
        print(f"Starting SSE MCP server (SSE + /files) on 0.0.0.0:{port} via uvicorn...", file=sys.stderr)
        uvicorn.run(asgi_app, host="0.0.0.0", port=port)

    else:
        # Default: stdio transport for MCP tooling
        app.run(transport="stdio")


if __name__ == "__main__":  # pragma: no cover
    main()
