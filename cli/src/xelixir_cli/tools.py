"""CLI implementations for xelixir Excel tools.

Each function here corresponds to an MCP tool defined in ``mcp/xelixir.py``.
The CLI package lives under the repository-root ``cli/`` directory and reuses
Excel wrapper functions from ``mcp/src/excel``.
"""

from __future__ import annotations

import argparse
import json
import os
import sys
from pathlib import Path
from typing import Any
from urllib.parse import quote

# Make mcp/src importable so we can reuse the existing Excel wrappers.
_REPO_ROOT = Path(__file__).resolve().parents[3]
_MCP_SRC_DIR = _REPO_ROOT / "mcp" / "src"
if str(_MCP_SRC_DIR) not in sys.path:
    sys.path.insert(0, str(_MCP_SRC_DIR))

from excel import (  # type: ignore
    append_rows,
    apply_formula,
    copy_range,
    copy_worksheet,
    create_chart,
    create_excel,
    create_pivot_table,
    create_sheet,
    delete_range,
    delete_worksheet,
    format_range,
    list_sheets,
    merge_cells,
    read_excel,
    rename_worksheet,
    unmerge_cells,
    validate_excel_range,
    validate_formula_syntax,
    write_excel,
    write_range,
)

DEFAULT_SHARED_DIR = "/mnt/data"
DEFAULT_PUBLIC_BASE_URL = ""


def get_shared_directory() -> str:
    """Return the directory that should be exposed under `/files`."""
    shared_dir = os.environ.get("WORKSPACE_DIR", DEFAULT_SHARED_DIR)
    os.makedirs(shared_dir, exist_ok=True)
    return shared_dir


def get_public_base_url() -> str:
    """Return the public base URL used to build file download URLs."""
    return os.environ.get(
        "EXCEL_PUBLIC_BASE_URL",
        DEFAULT_PUBLIC_BASE_URL,
    ).rstrip("/")


def build_download_url_for_path(file_path: str) -> str:
    """Build a public download URL for a file under the shared directory."""
    base_url = get_public_base_url()
    if not base_url:
        raise RuntimeError(
            "EXCEL_PUBLIC_BASE_URL is not configured. "
            "Set it to the public base URL for this CLI/server."
        )

    shared_dir = Path(get_shared_directory()).resolve()
    target_path = Path(file_path).resolve()

    try:
        rel = target_path.relative_to(shared_dir)
    except ValueError as exc:
        raise ValueError(
            "File path must be under the shared directory to be downloadable. "
            f"WORKSPACE_DIR={shared_dir} but got path={target_path}"
        ) from exc

    rel_posix = rel.as_posix()
    return f"{base_url}/files/{quote(rel_posix)}"


def output_json(data: dict[str, Any], pretty: bool = False) -> None:
    """Write JSON to stdout."""
    if pretty:
        print(json.dumps(data, indent=2, ensure_ascii=False))
        return
    print(json.dumps(data, ensure_ascii=False))


def error_exit(
    error_type: str,
    message: str,
    *,
    details: str | None = None,
    command: str | None = None,
    exit_code: int = 4,
) -> None:
    """Write structured error JSON to stderr and exit."""
    error_obj: dict[str, Any] = {
        "error_type": error_type,
        "message": message,
    }
    if details is not None:
        error_obj["details"] = details
    if command is not None:
        error_obj["command"] = command
    print(json.dumps(error_obj, indent=2, ensure_ascii=False), file=sys.stderr)
    raise SystemExit(exit_code)


def parse_json_input(value: str, label: str) -> Any:
    """Parse JSON from a file path or inline JSON string."""
    candidate = Path(value)
    if candidate.exists() and candidate.is_file():
        with candidate.open("r", encoding="utf-8") as fh:
            return json.load(fh)

    try:
        return json.loads(value)
    except json.JSONDecodeError as exc:
        error_exit(
            "InvalidArgument",
            f"Failed to parse {label}: {exc}",
            details=value,
            exit_code=3,
        )


def maybe_add_download_url(result: dict[str, Any], path: str) -> dict[str, Any]:
    """Add download_url when the current environment allows it."""
    try:
        result["download_url"] = build_download_url_for_path(path)
    except (RuntimeError, ValueError):
        pass
    return result


def cmd_create_excel(args: argparse.Namespace) -> None:
    try:
        sheet_name = args.sheet_name or "Sheet1"
        create_excel(args.path, sheet_name)
        output_json(
            maybe_add_download_url(
                {
                    "message": (
                        f"Created Excel workbook at {args.path} "
                        f"with sheet '{sheet_name}'"
                    ),
                    "path": args.path,
                },
                args.path,
            ),
            args.pretty,
        )
    except Exception as exc:
        error_exit("CreateExcelError", str(exc), command="create-excel")


def cmd_read_excel(args: argparse.Namespace) -> None:
    try:
        data = read_excel(args.path, args.sheet_name, args.range_str)
        output_json(
            {
                "path": args.path,
                "+sheet": args.sheet_name,
                "range": args.range_str,
                "data": data,
            },
            args.pretty,
        )
    except Exception as exc:
        error_exit("ReadExcelError", str(exc), command="read-excel")


def cmd_write_excel(args: argparse.Namespace) -> None:
    try:
        data = parse_json_input(args.data_json, "--data-json")
        write_excel(args.path, args.sheet_name, data)
        output_json(
            maybe_add_download_url(
                {
                    "message": (
                        f"Wrote {len(data)} rows to "
                        f"{args.path}:{args.sheet_name}!A1"
                    ),
                    "path": args.path,
                },
                args.path,
            ),
            args.pretty,
        )
    except Exception as exc:
        error_exit("WriteExcelError", str(exc), command="write-excel")


def cmd_write_range(args: argparse.Namespace) -> None:
    try:
        data = parse_json_input(args.data_json, "--data-json")
        write_range(args.path, args.sheet_name, args.start_cell, data)
        output_json(
            maybe_add_download_url(
                {
                    "message": (
                        f"Wrote {len(data)} rows to "
                        f"{args.path}:{args.sheet_name}!{args.start_cell}"
                    ),
                    "path": args.path,
                    "start_cell": args.start_cell,
                },
                args.path,
            ),
            args.pretty,
        )
    except Exception as exc:
        error_exit("WriteRangeError", str(exc), command="write-range")


def cmd_append_rows(args: argparse.Namespace) -> None:
    try:
        rows = parse_json_input(args.rows_json, "--rows-json")
        anchor = args.anchor_column or "A"
        start_row = append_rows(
            args.path,
            args.sheet_name,
            rows,
            anchor_column=anchor,
        )
        output_json(
            maybe_add_download_url(
                {
                    "message": (
                        f"Appended {len(rows)} rows to "
                        f"{args.path}:{args.sheet_name} at row "
                        f"{start_row + 1} (anchor={anchor})"
                    ),
                    "path": args.path,
                    "anchor_column": anchor,
                    "start_row": start_row,
                },
                args.path,
            ),
            args.pretty,
        )
    except Exception as exc:
        error_exit("AppendRowsError", str(exc), command="append-rows")


def cmd_create_sheet(args: argparse.Namespace) -> None:
    try:
        create_sheet(args.path, args.sheet_name)
        output_json(
            maybe_add_download_url(
                {
                    "message": f"Created sheet '{args.sheet_name}' in {args.path}",
                    "path": args.path,
                },
                args.path,
            ),
            args.pretty,
        )
    except Exception as exc:
        error_exit("CreateSheetError", str(exc), command="create-sheet")


def cmd_rename_worksheet(args: argparse.Namespace) -> None:
    try:
        rename_worksheet(args.path, args.old_name, args.new_name)
        output_json(
            maybe_add_download_url(
                {
                    "message": (
                        f"Renamed sheet '{args.old_name}' to "
                        f"'{args.new_name}' in {args.path}"
                    ),
                    "path": args.path,
                },
                args.path,
            ),
            args.pretty,
        )
    except Exception as exc:
        error_exit(
            "RenameWorksheetError",
            str(exc),
            command="rename-worksheet",
        )


def cmd_delete_worksheet(args: argparse.Namespace) -> None:
    try:
        delete_worksheet(args.path, args.sheet_name)
        output_json(
            maybe_add_download_url(
                {
                    "message": f"Deleted sheet '{args.sheet_name}' in {args.path}",
                    "path": args.path,
                },
                args.path,
            ),
            args.pretty,
        )
    except Exception as exc:
        error_exit(
            "DeleteWorksheetError",
            str(exc),
            command="delete-worksheet",
        )


def cmd_copy_worksheet(args: argparse.Namespace) -> None:
    try:
        copy_worksheet(args.path, args.source_sheet, args.target_sheet)
        output_json(
            maybe_add_download_url(
                {
                    "message": (
                        f"Copied sheet '{args.source_sheet}' to "
                        f"'{args.target_sheet}' in {args.path}"
                    ),
                    "path": args.path,
                },
                args.path,
            ),
            args.pretty,
        )
    except Exception as exc:
        error_exit("CopyWorksheetError", str(exc), command="copy-worksheet")


def cmd_apply_formula(args: argparse.Namespace) -> None:
    try:
        apply_formula(args.path, args.sheet_name, args.cell, args.formula)
        output_json(
            maybe_add_download_url(
                {
                    "message": (
                        f"Applied formula '{args.formula}' to "
                        f"{args.path}:{args.sheet_name}!{args.cell}"
                    ),
                    "path": args.path,
                },
                args.path,
            ),
            args.pretty,
        )
    except Exception as exc:
        error_exit("ApplyFormulaError", str(exc), command="apply-formula")


def cmd_validate_formula_syntax(args: argparse.Namespace) -> None:
    try:
        ok = validate_formula_syntax(args.path, args.sheet_name, args.formula)
        output_json({"valid": ok, "formula": args.formula}, args.pretty)
    except Exception as exc:
        error_exit(
            "ValidateFormulaSyntaxError",
            str(exc),
            command="validate-formula-syntax",
        )


def cmd_format_range(args: argparse.Namespace) -> None:
    try:
        format_range(
            args.path,
            args.sheet_name,
            args.start_cell,
            args.end_cell,
            args.bold,
            args.italic,
            args.font_size,
            args.font_color,
            args.bg_color,
        )
        output_json(
            maybe_add_download_url(
                {
                    "message": (
                        f"Formatted range {args.sheet_name}!"
                        f"{args.start_cell}:{args.end_cell} in {args.path}"
                    ),
                    "path": args.path,
                },
                args.path,
            ),
            args.pretty,
        )
    except Exception as exc:
        error_exit("FormatRangeError", str(exc), command="format-range")


def cmd_merge_cells(args: argparse.Namespace) -> None:
    try:
        merge_cells(args.path, args.sheet_name, args.start_cell, args.end_cell)
        output_json(
            maybe_add_download_url(
                {
                    "message": (
                        f"Merged cells {args.sheet_name}!"
                        f"{args.start_cell}:{args.end_cell} in {args.path}"
                    ),
                    "path": args.path,
                },
                args.path,
            ),
            args.pretty,
        )
    except Exception as exc:
        error_exit("MergeCellsError", str(exc), command="merge-cells")


def cmd_unmerge_cells(args: argparse.Namespace) -> None:
    try:
        unmerge_cells(args.path, args.sheet_name, args.start_cell, args.end_cell)
        output_json(
            maybe_add_download_url(
                {
                    "message": (
                        f"Unmerged cells {args.sheet_name}!"
                        f"{args.start_cell}:{args.end_cell} in {args.path}"
                    ),
                    "path": args.path,
                },
                args.path,
            ),
            args.pretty,
        )
    except Exception as exc:
        error_exit("UnmergeCellsError", str(exc), command="unmerge-cells")


def cmd_copy_range(args: argparse.Namespace) -> None:
    try:
        copy_range(
            args.path,
            args.sheet_name,
            args.source_start,
            args.source_end,
            args.target_start,
            args.target_sheet,
            copy_style=args.copy_style,
        )
        output_json(
            maybe_add_download_url(
                {
                    "message": (
                        f"Copied range {args.sheet_name}!"
                        f"{args.source_start}:{args.source_end} to "
                        f"{args.target_sheet or args.sheet_name}!"
                        f"{args.target_start} in {args.path} "
                        f"(copy_style={args.copy_style})"
                    ),
                    "path": args.path,
                },
                args.path,
            ),
            args.pretty,
        )
    except Exception as exc:
        error_exit("CopyRangeError", str(exc), command="copy-range")


def cmd_delete_range(args: argparse.Namespace) -> None:
    try:
        delete_range(
            args.path,
            args.sheet_name,
            args.start_cell,
            args.end_cell,
            args.shift_direction,
        )
        output_json(
            maybe_add_download_url(
                {
                    "message": (
                        f"Deleted range {args.sheet_name}!"
                        f"{args.start_cell}:{args.end_cell} in "
                        f"{args.path} (shift={args.shift_direction})"
                    ),
                    "path": args.path,
                },
                args.path,
            ),
            args.pretty,
        )
    except Exception as exc:
        error_exit("DeleteRangeError", str(exc), command="delete-range")


def cmd_validate_excel_range(args: argparse.Namespace) -> None:
    try:
        ok = validate_excel_range(
            args.path,
            args.sheet_name,
            args.start_cell,
            args.end_cell,
        )
        output_json(
            {
                "valid": ok,
                "sheet": args.sheet_name,
                "start": args.start_cell,
                "end": args.end_cell,
            },
            args.pretty,
        )
    except Exception as exc:
        error_exit(
            "ValidateExcelRangeError",
            str(exc),
            command="validate-excel-range",
        )


def cmd_create_chart(args: argparse.Namespace) -> None:
    try:
        create_chart(
            args.path,
            args.sheet_name,
            args.data_range,
            args.chart_type,
            args.target_cell,
            args.title,
            args.x_axis,
            args.y_axis,
        )
        output_json(
            maybe_add_download_url(
                {
                    "message": (
                        f"Created chart of type '{args.chart_type}' at "
                        f"{args.sheet_name}!{args.target_cell} in {args.path}"
                    ),
                    "path": args.path,
                },
                args.path,
            ),
            args.pretty,
        )
    except Exception as exc:
        error_exit("CreateChartError", str(exc), command="create-chart")


def cmd_list_sheets(args: argparse.Namespace) -> None:
    try:
        sheets = list_sheets(args.path)
        output_json({"path": args.path, "sheets": sheets}, args.pretty)
    except Exception as exc:
        error_exit("ListSheetsError", str(exc), command="list-sheets")


def cmd_create_pivot_table(args: argparse.Namespace) -> None:
    try:
        rows = parse_json_input(args.rows_json, "--rows-json")
        values = parse_json_input(args.values_json, "--values-json")
        columns = None
        if args.columns_json:
            columns = parse_json_input(args.columns_json, "--columns-json")
        message = create_pivot_table(
            args.path,
            args.sheet_name,
            args.data_range,
            rows,
            values,
            columns,
            args.agg_func,
        )
        output_json({"message": message}, args.pretty)
    except Exception as exc:
        error_exit(
            "CreatePivotTableError",
            str(exc),
            command="create-pivot-table",
        )


def create_parser() -> argparse.ArgumentParser:
    """Create the top-level parser for xelixir-tool."""
    parser = argparse.ArgumentParser(
        prog="xelixir-tool",
        description="CLI for Excel operations backed by Java tools.",
    )
    parser.add_argument(
        "--pretty",
        action="store_true",
        default=False,
        help="Pretty-print JSON output",
    )

    subparsers = parser.add_subparsers(dest="command", help="Available commands")
    tool_parser = subparsers.add_parser("tool", help="Run Excel tool commands")
    tool_subparsers = tool_parser.add_subparsers(
        dest="tool_command",
        help="Available tool commands",
    )

    p = tool_subparsers.add_parser("create-excel", help="Create a new Excel workbook")
    p.add_argument("--path", required=True, help="Destination file path")
    p.add_argument("--sheet-name", default="Sheet1", help="Initial worksheet name")
    p.set_defaults(func=cmd_create_excel)

    p = tool_subparsers.add_parser("read-excel", help="Read from an Excel file")
    p.add_argument("--path", required=True, help="Workbook path")
    p.add_argument("--sheet-name", required=True, help="Worksheet name")
    p.add_argument("--range-str", required=True, help="A1-style range (e.g. A1:C10)")
    p.set_defaults(func=cmd_read_excel)

    p = tool_subparsers.add_parser("write-excel", help="Write to an Excel file")
    p.add_argument("--path", required=True, help="Workbook path")
    p.add_argument("--sheet-name", required=True, help="Worksheet name")
    p.add_argument("--data-json", required=True, help="2D array JSON or file path")
    p.set_defaults(func=cmd_write_excel)

    p = tool_subparsers.add_parser("write-range", help="Write to a cell range")
    p.add_argument("--path", required=True, help="Workbook path")
    p.add_argument("--sheet-name", required=True, help="Worksheet name")
    p.add_argument("--start-cell", required=True, help="Top-left cell (e.g. B3)")
    p.add_argument("--data-json", required=True, help="2D array JSON or file path")
    p.set_defaults(func=cmd_write_range)

    p = tool_subparsers.add_parser("append-rows", help="Append rows to first empty row")
    p.add_argument("--path", required=True, help="Workbook path")
    p.add_argument("--sheet-name", required=True, help="Worksheet name")
    p.add_argument("--rows-json", required=True, help="Rows JSON array or file path")
    p.add_argument("--anchor-column", default="A", help="Anchor column letter")
    p.set_defaults(func=cmd_append_rows)

    p = tool_subparsers.add_parser("create-sheet", help="Create a new worksheet")
    p.add_argument("--path", required=True, help="Workbook path")
    p.add_argument("--sheet-name", required=True, help="New worksheet name")
    p.set_defaults(func=cmd_create_sheet)

    p = tool_subparsers.add_parser("rename-worksheet", help="Rename a worksheet")
    p.add_argument("--path", required=True, help="Workbook path")
    p.add_argument("--old-name", required=True, help="Current sheet name")
    p.add_argument("--new-name", required=True, help="New sheet name")
    p.set_defaults(func=cmd_rename_worksheet)

    p = tool_subparsers.add_parser("delete-worksheet", help="Delete a worksheet")
    p.add_argument("--path", required=True, help="Workbook path")
    p.add_argument("--sheet-name", required=True, help="Sheet name to delete")
    p.set_defaults(func=cmd_delete_worksheet)

    p = tool_subparsers.add_parser("copy-worksheet", help="Copy a worksheet")
    p.add_argument("--path", required=True, help="Workbook path")
    p.add_argument("--source-sheet", required=True, help="Source sheet name")
    p.add_argument("--target-sheet", required=True, help="Target sheet name")
    p.set_defaults(func=cmd_copy_worksheet)

    p = tool_subparsers.add_parser("apply-formula", help="Apply a formula to a cell")
    p.add_argument("--path", required=True, help="Workbook path")
    p.add_argument("--sheet-name", required=True, help="Worksheet name")
    p.add_argument("--cell", required=True, help="Cell address (e.g. C1)")
    p.add_argument("--formula", required=True, help="Formula string")
    p.set_defaults(func=cmd_apply_formula)

    p = tool_subparsers.add_parser(
        "validate-formula-syntax",
        help="Validate formula syntax",
    )
    p.add_argument("--path", required=True, help="Workbook path")
    p.add_argument("--sheet-name", required=True, help="Worksheet name")
    p.add_argument("--formula", required=True, help="Formula to validate")
    p.set_defaults(func=cmd_validate_formula_syntax)

    p = tool_subparsers.add_parser("format-range", help="Format a cell range")
    p.add_argument("--path", required=True, help="Workbook path")
    p.add_argument("--sheet-name", required=True, help="Worksheet name")
    p.add_argument("--start-cell", required=True, help="Start cell")
    p.add_argument("--end-cell", required=True, help="End cell")
    p.add_argument("--bold", action="store_true", default=False, help="Apply bold")
    p.add_argument("--italic", action="store_true", default=False, help="Apply italic")
    p.add_argument("--font-size", type=int, default=0, help="Font size")
    p.add_argument("--font-color", default="", help="Font color hex")
    p.add_argument("--bg-color", default="", help="Background color hex")
    p.set_defaults(func=cmd_format_range)

    p = tool_subparsers.add_parser("merge-cells", help="Merge cells")
    p.add_argument("--path", required=True, help="Workbook path")
    p.add_argument("--sheet-name", required=True, help="Worksheet name")
    p.add_argument("--start-cell", required=True, help="Start cell")
    p.add_argument("--end-cell", required=True, help="End cell")
    p.set_defaults(func=cmd_merge_cells)

    p = tool_subparsers.add_parser("unmerge-cells", help="Unmerge cells")
    p.add_argument("--path", required=True, help="Workbook path")
    p.add_argument("--sheet-name", required=True, help="Worksheet name")
    p.add_argument("--start-cell", required=True, help="Start cell")
    p.add_argument("--end-cell", required=True, help="End cell")
    p.set_defaults(func=cmd_unmerge_cells)

    p = tool_subparsers.add_parser("copy-range", help="Copy a cell range")
    p.add_argument("--path", required=True, help="Workbook path")
    p.add_argument("--sheet-name", required=True, help="Source sheet name")
    p.add_argument("--source-start", required=True, help="Source start cell")
    p.add_argument("--source-end", required=True, help="Source end cell")
    p.add_argument("--target-start", required=True, help="Target start cell")
    p.add_argument("--target-sheet", default=None, help="Target sheet name")
    p.add_argument(
        "--copy-style",
        action="store_true",
        default=True,
        help="Copy styles too",
    )
    p.set_defaults(func=cmd_copy_range)

    p = tool_subparsers.add_parser("delete-range", help="Delete a cell range")
    p.add_argument("--path", required=True, help="Workbook path")
    p.add_argument("--sheet-name", required=True, help="Worksheet name")
    p.add_argument("--start-cell", required=True, help="Start cell")
    p.add_argument("--end-cell", required=True, help="End cell")
    p.add_argument(
        "--shift-direction",
        default="up",
        choices=["up", "left"],
        help="Shift direction",
    )
    p.set_defaults(func=cmd_delete_range)

    p = tool_subparsers.add_parser(
        "validate-excel-range",
        help="Validate Excel range",
    )
    p.add_argument("--path", required=True, help="Workbook path")
    p.add_argument("--sheet-name", required=True, help="Worksheet name")
    p.add_argument("--start-cell", required=True, help="Start cell")
    p.add_argument("--end-cell", default=None, help="End cell (optional)")
    p.set_defaults(func=cmd_validate_excel_range)

    p = tool_subparsers.add_parser("create-chart", help="Create a chart")
    p.add_argument("--path", required=True, help="Workbook path")
    p.add_argument("--sheet-name", required=True, help="Worksheet name")
    p.add_argument("--data-range", required=True, help="Data range")
    p.add_argument("--chart-type", required=True, help="Chart type")
    p.add_argument("--target-cell", required=True, help="Target cell")
    p.add_argument("--title", default=None, help="Chart title")
    p.add_argument("--x-axis", default=None, help="X-axis label")
    p.add_argument("--y-axis", default=None, help="Y-axis label")
    p.set_defaults(func=cmd_create_chart)

    p = tool_subparsers.add_parser("list-sheets", help="List worksheet names")
    p.add_argument("--path", required=True, help="Workbook path")
    p.set_defaults(func=cmd_list_sheets)

    p = tool_subparsers.add_parser(
        "create-pivot-table",
        help="Create a pivot table",
    )
    p.add_argument("--path", required=True, help="Workbook path")
    p.add_argument("--sheet-name", required=True, help="Worksheet name")
    p.add_argument("--data-range", required=True, help="Data range")
    p.add_argument("--rows-json", required=True, help="Rows JSON")
    p.add_argument("--values-json", required=True, help="Values JSON")
    p.add_argument("--columns-json", default=None, help="Columns JSON")
    p.add_argument(
        "--agg-func",
        default="sum",
        choices=["sum", "count", "average", "max", "min"],
        help="Aggregation function",
    )
    p.set_defaults(func=cmd_create_pivot_table)

    return parser


def main() -> None:
    """CLI main entry point."""
    parser = create_parser()
    args = parser.parse_args()

    if not args.command:
        parser.print_help()
        raise SystemExit(0)

    if args.command != "tool":
        parser.print_help()
        raise SystemExit(0)

    if not args.tool_command:
        print("Error: Please specify a tool command.", file=sys.stderr)
        raise SystemExit(2)

    if not hasattr(args, "func"):
        print(
            f"Error: Unknown tool command: {args.tool_command}",
            file=sys.stderr,
        )
        raise SystemExit(2)

    try:
        args.func(args)
    except SystemExit:
        raise
    except Exception as exc:
        error_exit(
            "InternalError",
            str(exc),
            command=f"tool {args.tool_command}",
            exit_code=5,
        )
