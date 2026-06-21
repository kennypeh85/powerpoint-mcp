"""
Excel Live MCP Server
=====================
Attaches to a *running* Microsoft Excel instance via COM (pywin32) and
exposes read/edit/screenshot tools over stdio MCP so Hermes Agent can edit
the workbook live on screen.

Requirements
------------
- Windows + Microsoft Excel installed.
- pywin32 (py -m pip install pywin32)  -- run `python Scripts/pywin32_postinstall.py -install` once.
- mcp  (py -m pip install mcp)

Register in ~/.hermes/config.yaml:
    mcp_servers:
      excel:
        command: "C:/Users/Kenny Peh/AppData/Local/hermes/hermes-agent/venv/Scripts/python.exe"
        args: ["C:/Users/Kenny Peh/OneDrive/Documents/Claude/Projects/office-tools.ai/windows-mcp-servers/excel-mcp/server.py"]
        connect_timeout: 30
        timeout: 90

Notes
-----
- Excel COM is 1-indexed. Sheet/row/column indices in this server are 1-indexed
  to match what the user sees on screen.
- COM objects are apartment-threaded. Each tool call re-initialises COM and
  re-dispatches Excel.Application, which transparently reconnects to the
  already-running instance.
- All tools return JSON-serialisable dicts.
"""

from __future__ import annotations

import json
import os
import sys
import tempfile
import traceback
from contextlib import contextmanager
from functools import wraps
from typing import Any

import pythoncom
import win32com.client
from mcp.server.fastmcp import FastMCP


def excel_rgb(r: int, g: int, b: int) -> int:
    """Convert hex RGB components to Excel COM color value.

    COM stores color as: R + G*256 + B*65536  (NOT B + G*256 + R*65536!).
    Getting this wrong swaps red and blue channels — peach becomes light blue.
    Always use this helper instead of manual hex arithmetic.
    """
    return r + g * 256 + b * 65536


@contextmanager
def com_session():
    """Initialise COM for the current thread and yield a fresh app handle."""
    pythoncom.CoInitialize()
    try:
        # Dispatch returns the running instance if Excel is already open,
        # otherwise it launches Excel (visible if Interactive).
        app = win32com.client.Dispatch("Excel.Application")
        try:
            yield app
        finally:
            # Do NOT quit the app — the user owns the Excel process.
            del app
    finally:
        pythoncom.CoUninitialize()


def excel_tool(fn):
    """Decorator: run a tool body inside a COM session with robust error capture."""
    @wraps(fn)
    def wrapper(*args, **kwargs):
        try:
            with com_session() as app:
                return fn(app, *args, **kwargs)
        except Exception as e:
            tb = traceback.format_exc()
            return {"error": f"{type(e).__name__}: {e}", "traceback": tb}
    return wrapper


def _resolve_workbook(app, name: str | None):
    """Return the requested Workbook. If name is None, return ActiveWorkbook."""
    if name:
        for i in range(1, app.Workbooks.Count + 1):
            wb = app.Workbooks.Item(i)
            # Match by Name (no extension) or full Path.
            if wb.Name == name or wb.Name == name + ".xlsx" or wb.Path and wb.FullName == name:
                return wb
        raise ValueError(f"No open workbook matching '{name}'.")
    wb = app.ActiveWorkbook
    if wb is None:
        raise RuntimeError("No active workbook. Open a file in Excel first.")
    return wb


def _resolve_sheet(wb, name_or_index: str | int, sheet_type: str = "worksheets"):
    """Resolve a worksheet. Tries: 1-indexed int → exact name → text content."""
    # 1. Numeric index
    if isinstance(name_or_index, int):
        if 1 <= name_or_index <= getattr(wb, sheet_type).Count:
            return getattr(wb, sheet_type).Item(name_or_index)
        return None
    key = str(name_or_index)
    sheets = getattr(wb, sheet_type)
    # 2. Exact name match
    for i in range(1, sheets.Count + 1):
        sh = sheets.Item(i)
        if sh.Name == key:
            return sh
    # 3. Exact text-content match
    for i in range(1, sheets.Count + 1):
        sh = sheets.Item(i)
        try:
            # Try sheet.Name, sheet.CodeName, or cell text
            name = sh.Name
            if name == key:
                return sh
            # Check cell A1 text as fallback
            cell = sh.Cells(1, 1)
            if cell.Value and str(cell.Value).strip() == key:
                return sh
        except Exception:
            continue
    return None


def _range_summary(rng) -> dict[str, Any]:
    """Extract a compact, JSON-safe description of a range."""
    try:
        area = rng.Area
    except Exception:
        area = rng
    rows = area.Rows.Count
    cols = area.Columns.Count
    info: dict[str, Any] = {
        "rows": rows,
        "cols": cols,
        "address": str(area.Address).replace("$", ""),
        "row1": area.Row,
        "col1": area.Column,
        "row2": area.Row + rows - 1,
        "col2": area.Column + cols - 1,
    }
    try:
        # Try to get first cell value
        first_cell = area.Cells(1, 1)
        if first_cell.Value:
            info["value"] = str(first_cell.Value)
    except Exception:
        pass
    return info


# ---------------------------------------------------------------------------
# MCP server + tools
# ---------------------------------------------------------------------------

mcp = FastMCP("excel-live")


# --- Session / workbook discovery -------------------------------------------

@mcp.tool()
@excel_tool
def list_workbooks(app) -> str:
    """List all workbooks currently open in Excel.

    Returns name, full path, sheet count, and whether each is the active workbook.
    """
    workbooks = []
    active = None
    try:
        active = app.ActiveWorkbook
    except Exception:
        active = None
    active_name = active.Name if active else None
    for i in range(1, app.Workbooks.Count + 1):
        wb = app.Workbooks.Item(i)
        try:
            sheet_count = wb.Worksheets.Count
            path = wb.FullName if wb.Path else "(unsaved)"
        except Exception:
            sheet_count = "(unknown)"
            path = wb.Name
        workbooks.append({
            "name": wb.Name,
            "path": path,
            "sheet_count": sheet_count,
            "active": wb.Name == active_name,
        })
    return json.dumps({"workbooks": workbooks, "active": active_name, "count": len(workbooks)})


@mcp.tool()
@excel_tool
def get_active_workbook(app) -> str:
    """Return details of the currently active workbook."""
    wb = app.ActiveWorkbook
    if wb is None:
        return json.dumps({"error": "No active workbook."})
    try:
        sheet_count = wb.Worksheets.Count
        name = wb.Name
        path = wb.FullName if wb.Path else "(unsaved)"
    except Exception:
        sheet_count = 0
        name = wb.Name
        path = "(unknown)"
    return json.dumps({
        "name": name,
        "path": path,
        "sheet_count": sheet_count,
    })


# --- Read -------------------------------------------------------------------

@mcp.tool()
@excel_tool
def list_sheets(app, workbook: str | None = None) -> str:
    """List all sheets (worksheets) in a workbook.

    If workbook is None, uses the active workbook.
    """
    wb = _resolve_workbook(app, workbook)
    sheets = []
    for i in range(1, wb.Worksheets.Count + 1):
        sh = wb.Worksheets.Item(i)
        try:
            active = sh.Index == wb.ActiveSheet.Index
        except Exception:
            active = False
        sheets.append({
            "name": sh.Name,
            "index": i,
            "active": active,
        })
    return json.dumps({"workbook": wb.Name, "sheets": sheets, "count": len(sheets)})


@mcp.tool()
@excel_tool
def read_cell(app, row: int, col: int, sheet: str | int = 1, workbook: str | None = None) -> str:
    """Read a single cell value and metadata.

    row and col are 1-indexed.
    """
    wb = _resolve_workbook(app, workbook)
    sh = _resolve_sheet(wb, sheet, "Worksheets")
    if sh is None:
        return json.dumps({"error": f"Sheet '{sheet}' not found."})
    cell = sh.Cells(row, col)
    try:
        value = cell.Value
        formula = cell.Formula if cell.HasFormula else None
        number_format = cell.NumberFormat
    except Exception:
        value = None
        formula = None
        number_format = None
    info: dict[str, Any] = {
        "sheet": sh.Name,
        "row": row,
        "col": col,
        "address": str(cell.Address).replace("$", ""),
    }
    if value is not None:
        info["value"] = str(value)
    if formula is not None:
        info["formula"] = formula
    if number_format is not None:
        info["number_format"] = number_format
    return json.dumps(info)


@mcp.tool()
@excel_tool
def read_range(
    app,
    rows: str | int,
    cols: str | int,
    sheet: str | int = 1,
    workbook: str | None = None,
) -> str:
    """Read a range of cells.

    rows and cols can be:
    - 'A1' (single cell)
    - 'A1:B10' (range)
    - 1 (single cell at row 1, col 1)
    - 'B2' (single cell)
    """
    wb = _resolve_workbook(app, workbook)
    sh = _resolve_sheet(wb, sheet, "Worksheets")
    if sh is None:
        return json.dumps({"error": f"Sheet '{sheet}' not found."})
    # Parse range
    if isinstance(rows, int):
        row1 = row2 = rows
    elif isinstance(rows, str) and "," not in rows and ":" not in rows:
        # Single cell like 'A1' or row number
        if ":" in rows:
            row1, row2 = rows.split(":")
        else:
            row1 = row2 = rows
    else:
        # Range like 'A1:B10'
        if ":" in rows:
            row1, row2 = rows.split(":")
        else:
            row1 = row2 = rows
    if isinstance(cols, int):
        col1 = col2 = cols
    elif isinstance(cols, str) and "," not in cols and ":" not in cols:
        # Single column
        if ":" in cols:
            col1, col2 = cols.split(":")
        else:
            col1 = col2 = cols
    else:
        # Range like 'A:C'
        if ":" in cols:
            col1, col2 = cols.split(":")
        else:
            col1 = col2 = cols
    # Convert letters to numbers
    if isinstance(col1, str):
        col1_num = 0
        for c in col1.upper():
            col1_num = col1_num * 26 + (ord(c) - ord('A') + 1)
    if isinstance(col2, str):
        col2_num = 0
        for c in col2.upper():
            col2_num = col2_num * 26 + (ord(c) - ord('A') + 1)
    if isinstance(row1, str):
        row1_num = int(row1)
    if isinstance(row2, str):
        row2_num = int(row2)
    if isinstance(col1, int):
        col1_num = col1
    if isinstance(col2, int):
        col2_num = col2
    # Build range reference
    col1_letter = ""
    col = col1_num
    while col > 0:
        col, rem = divmod(col - 1, 26)
        col1_letter = chr(rem + ord('A')) + col1_letter
    range_ref = f"{col1_letter}{row1_num}:{chr(col2_num + ord('A') - 1)}{row2_num}"
    rng = sh.Range(range_ref)
    return json.dumps(_range_summary(rng))


@mcp.tool()
@excel_tool
def read_table(
    app,
    table_name: str | None = None,
    sheet: str | int = 1,
    workbook: str | None = None,
) -> str:
    """Read data from a table (ListObject).

    If table_name is None, uses the active table in the active sheet.
    """
    wb = _resolve_workbook(app, workbook)
    sh = _resolve_sheet(wb, sheet, "Worksheets")
    if sh is None:
        return json.dumps({"error": f"Sheet '{sheet}' not found."})
    # Find table by name or active
    if table_name:
        tables = sh.ListObjects
        for i in range(1, tables.Count + 1):
            tbl = tables.Item(i)
            if tbl.Name == table_name:
                break
        else:
            return json.dumps({"error": f"Table '{table_name}' not found on sheet '{sh.Name}'."})
    else:
        # Try to find the active table
        try:
            tbl = sh.ListObjects(sh.ListObjects.Count)  # Last table
        except Exception:
            return json.dumps({"error": "No tables found or active table."})
    # Read data
    data = []
    for row in tbl.DataBodyRange.Rows:
        row_data = []
        for cell in row.Cells:
            row_data.append(str(cell.Value) if cell.Value is not None else "")
        data.append(row_data)
    return json.dumps({
        "table_name": tbl.Name,
        "sheet": sh.Name,
        "data": data,
        "rows": len(data),
        "cols": len(data[0]) if data else 0,
    })


@mcp.tool()
@excel_tool
def get_sheet_content(
    app,
    sheet: str | int = 1,
    workbook: str | None = None,
) -> str:
    """Read all used cells from a worksheet.

    Returns a 2D array with cell values.
    """
    wb = _resolve_workbook(app, workbook)
    sh = _resolve_sheet(wb, sheet, "Worksheets")
    if sh is None:
        return json.dumps({"error": f"Sheet '{sheet}' not found."})
    try:
        used_range = sh.UsedRange
    except Exception:
        used_range = sh.Cells(1, 1)
    return json.dumps({
        "sheet": sh.Name,
        "rows": used_range.Rows.Count,
        "cols": used_range.Columns.Count,
        "address": str(used_range.Address).replace("$", ""),
    })


# --- Edit -------------------------------------------------------------------

@mcp.tool()
@excel_tool
def write_cell(
    app,
    row: int,
    col: int,
    value: str,
    sheet: str | int = 1,
    workbook: str | None = None,
) -> str:
    """Write a value to a cell.

    row and col are 1-indexed.
    """
    wb = _resolve_workbook(app, workbook)
    sh = _resolve_sheet(wb, sheet, "Worksheets")
    if sh is None:
        return json.dumps({"error": f"Sheet '{sheet}' not found."})
    cell = sh.Cells(row, col)
    cell.Value = value
    return json.dumps({
        "ok": True,
        "sheet": sh.Name,
        "row": row,
        "col": col,
        "value": value,
    })


@mcp.tool()
@excel_tool
def write_range(
    app,
    rows: str | int,
    cols: str | int,
    value: str,
    sheet: str | int = 1,
    workbook: str | None = None,
) -> str:
    """Write a value to a range of cells.

    rows and cols can be:
    - 'A1' (single cell)
    - 'A1:B10' (range)
    - 1 (single cell at row 1, col 1)
    - 'B2' (single cell)
    """
    wb = _resolve_workbook(app, workbook)
    sh = _resolve_sheet(wb, sheet, "Worksheets")
    if sh is None:
        return json.dumps({"error": f"Sheet '{sheet}' not found."})
    # Parse range (same as read_range)
    if isinstance(rows, int):
        row1 = row2 = rows
    elif isinstance(rows, str) and "," not in rows and ":" not in rows:
        if ":" in rows:
            row1, row2 = rows.split(":")
        else:
            row1 = row2 = rows
    else:
        if ":" in rows:
            row1, row2 = rows.split(":")
        else:
            row1 = row2 = rows
    if isinstance(cols, int):
        col1 = col2 = cols
    elif isinstance(cols, str) and "," not in cols and ":" not in cols:
        if ":" in cols:
            col1, col2 = cols.split(":")
        else:
            col1 = col2 = cols
    else:
        if ":" in cols:
            col1, col2 = cols.split(":")
        else:
            col1 = col2 = cols
    # Convert letters to numbers
    if isinstance(col1, str):
        col1_num = 0
        for c in col1.upper():
            col1_num = col1_num * 26 + (ord(c) - ord('A') + 1)
    if isinstance(col2, str):
        col2_num = 0
        for c in col2.upper():
            col2_num = col2_num * 26 + (ord(c) - ord('A') + 1)
    if isinstance(row1, str):
        row1_num = int(row1)
    if isinstance(row2, str):
        row2_num = int(row2)
    if isinstance(col1, int):
        col1_num = col1
    if isinstance(col2, int):
        col2_num = col2
    # Build range reference
    col1_letter = ""
    col = col1_num
    while col > 0:
        col, rem = divmod(col - 1, 26)
        col1_letter = chr(rem + ord('A')) + col1_letter
    range_ref = f"{col1_letter}{row1_num}:{chr(col2_num + ord('A') - 1)}{row2_num}"
    rng = sh.Range(range_ref)
    rng.Value = value
    return json.dumps({
        "ok": True,
        "sheet": sh.Name,
        "range": range_ref,
        "value": value,
    })


@mcp.tool()
@excel_tool
def clear_cell(
    app,
    row: int,
    col: int,
    sheet: str | int = 1,
    workbook: str | None = None,
) -> str:
    """Clear a cell value."""
    wb = _resolve_workbook(app, workbook)
    sh = _resolve_sheet(wb, sheet, "Worksheets")
    if sh is None:
        return json.dumps({"error": f"Sheet '{sheet}' not found."})
    cell = sh.Cells(row, col)
    cell.ClearContents()
    return json.dumps({
        "ok": True,
        "sheet": sh.Name,
        "row": row,
        "col": col,
    })


@mcp.tool()
@excel_tool
def clear_range(
    app,
    rows: str | int,
    cols: str | int,
    sheet: str | int = 1,
    workbook: str | None = None,
) -> str:
    """Clear a range of cells."""
    wb = _resolve_workbook(app, workbook)
    sh = _resolve_sheet(wb, sheet, "Worksheets")
    if sh is None:
        return json.dumps({"error": f"Sheet '{sheet}' not found."})
    # Parse range (same as write_range)
    if isinstance(rows, int):
        row1 = row2 = rows
    elif isinstance(rows, str) and "," not in rows and ":" not in rows:
        if ":" in rows:
            row1, row2 = rows.split(":")
        else:
            row1 = row2 = rows
    else:
        if ":" in rows:
            row1, row2 = rows.split(":")
        else:
            row1 = row2 = rows
    if isinstance(cols, int):
        col1 = col2 = cols
    elif isinstance(cols, str) and "," not in cols and ":" not in cols:
        if ":" in cols:
            col1, col2 = cols.split(":")
        else:
            col1 = col2 = cols
    else:
        if ":" in cols:
            col1, col2 = cols.split(":")
        else:
            col1 = col2 = cols
    # Convert letters to numbers
    if isinstance(col1, str):
        col1_num = 0
        for c in col1.upper():
            col1_num = col1_num * 26 + (ord(c) - ord('A') + 1)
    if isinstance(col2, str):
        col2_num = 0
        for c in col2.upper():
            col2_num = col2_num * 26 + (ord(c) - ord('A') + 1)
    if isinstance(row1, str):
        row1_num = int(row1)
    if isinstance(row2, str):
        row2_num = int(row2)
    if isinstance(col1, int):
        col1_num = col1
    if isinstance(col2, int):
        col2_num = col2
    # Build range reference
    col1_letter = ""
    col = col1_num
    while col > 0:
        col, rem = divmod(col - 1, 26)
        col1_letter = chr(rem + ord('A')) + col1_letter
    range_ref = f"{col1_letter}{row1_num}:{chr(col2_num + ord('A') - 1)}{row2_num}"
    rng = sh.Range(range_ref)
    rng.ClearContents()
    return json.dumps({
        "ok": True,
        "sheet": sh.Name,
        "range": range_ref,
    })


@mcp.tool()
@excel_tool
def set_cell_format(
    app,
    row: int,
    col: int,
    font_name: str | None = None,
    font_size: float | None = None,
    bold: bool | None = None,
    italic: bool | None = None,
    color_rgb: str | None = None,
    number_format: str | None = None,
    sheet: str | int = 1,
    workbook: str | None = None,
) -> str:
    """Set cell formatting properties."""
    wb = _resolve_workbook(app, workbook)
    sh = _resolve_sheet(wb, sheet, "Worksheets")
    if sh is None:
        return json.dumps({"error": f"Sheet '{sheet}' not found."})
    cell = sh.Cells(row, col)
    font = cell.Font
    if font_name is not None:
        font.Name = font_name
    if font_size is not None:
        font.Size = font_size
    if bold is not None:
        font.Bold = -1 if bold else 0
    if italic is not None:
        font.Italic = -1 if italic else 0
    if color_rgb is not None:
        hex_clean = color_rgb.lstrip("#")
        font.Color.RGB = excel_rgb(int(hex_clean[0:2], 16), int(hex_clean[2:4], 16), int(hex_clean[4:6], 16))
    if number_format is not None:
        cell.NumberFormat = number_format
    return json.dumps({
        "ok": True,
        "sheet": sh.Name,
        "row": row,
        "col": col,
    })


@mcp.tool()
@excel_tool
def set_range_format(
    app,
    rows: str | int,
    cols: str | int,
    font_name: str | None = None,
    font_size: float | None = None,
    bold: bool | None = None,
    italic: bool | None = None,
    color_rgb: str | None = None,
    number_format: str | None = None,
    sheet: str | int = 1,
    workbook: str | None = None,
) -> str:
    """Set range formatting properties."""
    wb = _resolve_workbook(app, workbook)
    sh = _resolve_sheet(wb, sheet, "Worksheets")
    if sh is None:
        return json.dumps({"error": f"Sheet '{sheet}' not found."})
    # Parse range (same as write_range)
    if isinstance(rows, int):
        row1 = row2 = rows
    elif isinstance(rows, str) and "," not in rows and ":" not in rows:
        if ":" in rows:
            row1, row2 = rows.split(":")
        else:
            row1 = row2 = rows
    else:
        if ":" in rows:
            row1, row2 = rows.split(":")
        else:
            row1 = row2 = rows
    if isinstance(cols, int):
        col1 = col2 = cols
    elif isinstance(cols, str) and "," not in cols and ":" not in cols:
        if ":" in cols:
            col1, col2 = cols.split(":")
        else:
            col1 = col2 = cols
    else:
        if ":" in cols:
            col1, col2 = cols.split(":")
        else:
            col1 = col2 = cols
    # Convert letters to numbers
    if isinstance(col1, str):
        col1_num = 0
        for c in col1.upper():
            col1_num = col1_num * 26 + (ord(c) - ord('A') + 1)
    if isinstance(col2, str):
        col2_num = 0
        for c in col2.upper():
            col2_num = col2_num * 26 + (ord(c) - ord('A') + 1)
    if isinstance(row1, str):
        row1_num = int(row1)
    if isinstance(row2, str):
        row2_num = int(row2)
    if isinstance(col1, int):
        col1_num = col1
    if isinstance(col2, int):
        col2_num = col2
    # Build range reference
    col1_letter = ""
    col = col1_num
    while col > 0:
        col, rem = divmod(col - 1, 26)
        col1_letter = chr(rem + ord('A')) + col1_letter
    range_ref = f"{col1_letter}{row1_num}:{chr(col2_num + ord('A') - 1)}{row2_num}"
    rng = sh.Range(range_ref)
    font = rng.Font
    if font_name is not None:
        font.Name = font_name
    if font_size is not None:
        font.Size = font_size
    if bold is not None:
        font.Bold = -1 if bold else 0
    if italic is not None:
        font.Italic = -1 if italic else 0
    if color_rgb is not None:
        hex_clean = color_rgb.lstrip("#")
        font.Color.RGB = excel_rgb(int(hex_clean[0:2], 16), int(hex_clean[2:4], 16), int(hex_clean[4:6], 16))
    if number_format is not None:
        rng.NumberFormat = number_format
    return json.dumps({
        "ok": True,
        "sheet": sh.Name,
        "range": range_ref,
    })


@mcp.tool()
@excel_tool
def set_cell_fill(
    app,
    row: int,
    col: int,
    rgb: str,
    sheet: str | int = 1,
    workbook: str | None = None,
) -> str:
    """Set cell fill color."""
    wb = _resolve_workbook(app, workbook)
    sh = _resolve_sheet(wb, sheet, "Worksheets")
    if sh is None:
        return json.dumps({"error": f"Sheet '{sheet}' not found."})
    cell = sh.Cells(row, col)
    cell.Interior.Color = excel_rgb(int(rgb.lstrip("#")[0:2], 16), int(rgb.lstrip("#")[2:4], 16), int(rgb.lstrip("#")[4:6], 16))
    return json.dumps({
        "ok": True,
        "sheet": sh.Name,
        "row": row,
        "col": col,
        "rgb": rgb,
    })


@mcp.tool()
@excel_tool
def add_sheet(
    app,
    name: str | None = None,
    sheet_type: str = "worksheet",
    workbook: str | None = None,
) -> str:
    """Add a new sheet.

    sheet_type is either 'worksheet' or 'chart'.
    """
    wb = _resolve_workbook(app, workbook)
    if name is None:
        name = f"Sheet{wb.Worksheets.Count + 1}"
    if sheet_type == "worksheet":
        sh = wb.Worksheets.Add(Before=wb.Worksheets(wb.Worksheets.Count))
        sh.Name = name
    elif sheet_type == "chart":
        # Add chart sheet
        sh = wb.Charts.Add()
        sh.Name = name
    else:
        return json.dumps({"error": f"Unknown sheet_type '{sheet_type}'. Use 'worksheet' or 'chart'."})
    return json.dumps({
        "ok": True,
        "workbook": wb.Name,
        "sheet_name": sh.Name,
        "sheet_index": sh.Index,
    })


@mcp.tool()
@excel_tool
def delete_sheet(app, sheet: str | int, workbook: str | None = None) -> str:
    """Delete a sheet (1-indexed by index or name). Irreversible."""
    wb = _resolve_workbook(app, workbook)
    sh = _resolve_sheet(wb, sheet, "Worksheets")
    if sh is None:
        return json.dumps({"error": f"Sheet '{sheet}' not found."})
    if wb.Worksheets.Count <= 1:
        return json.dumps({"error": "Cannot delete the only sheet in the workbook."})
    sh.Delete()
    return json.dumps({
        "ok": True,
        "workbook": wb.Name,
        "deleted": sh.Name,
        "remaining": wb.Worksheets.Count,
    })


@mcp.tool()
@excel_tool
def export_to_pdf(
    app,
    file_path: str,
    workbook: str | None = None,
) -> str:
    """Export workbook to PDF.

    file_path is the full path (including .pdf extension).
    """
    wb = _resolve_workbook(app, workbook)
    if not file_path.lower().endswith(".pdf"):
        file_path += ".pdf"
    # Save as PDF
    wb.ExportAsFixedFormat(
        xlExportFormatPDF,
        file_path,
        OpenAfterPublish=False,
    )
    return json.dumps({
        "ok": True,
        "workbook": wb.Name,
        "pdf_path": file_path,
    })


# --- Run -------------------------------------------------------------------

if __name__ == "__main__":
    mcp.run(transport="stdio")
