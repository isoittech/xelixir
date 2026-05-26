"""Namespace package wrapper for excel tools when installed via uv.

This package simply re-exports the excel wrappers that live in
`mcp/src/excel`. It allows the `xelixir` package to import
`excel` without having to change the existing layout of the repository.
"""

from pathlib import Path
import sys

# Dynamically add the src directory (where the real `excel` package lives)
# to sys.path so that `from excel import ...` inside xelixir.py
# resolves to `mcp/src/excel`.
_this_dir = Path(__file__).resolve().parent
_repo_root = _this_dir.parent
_src_dir = _repo_root / "src"

# Ensure the src directory is on sys.path so that the real ``excel``
# package (under src/excel) is importable when this shim is imported
# as the top-level ``excel`` package within the uv-managed environment.

if str(_src_dir) not in sys.path:
    sys.path.insert(0, str(_src_dir))

from src.excel import (
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

# Import the real excel package (located in src/excel)
from excel import *  # type: ignore  # noqa: F401,F403
