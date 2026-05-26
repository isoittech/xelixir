"""Python wrappers around the Java-based Excel tools.

This package will later be used as the implementation of an MCP server.
The current design is simple: each public function corresponds to one
Java CLI tool under ``java/src/xelixir`` and invokes it via
``subprocess.run``.
"""

from .wrapper import (
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

__all__ = [
    "create_excel",
    "read_excel",
    "write_excel",
    "write_range",
    "append_rows",
    "create_sheet",
    "rename_worksheet",
    "delete_worksheet",
    "copy_worksheet",
    "apply_formula",
    "validate_formula_syntax",
    "format_range",
    "merge_cells",
    "unmerge_cells",
    "copy_range",
    "delete_range",
    "validate_excel_range",
    "create_chart",
    "create_pivot_table",
    "list_sheets",
]
