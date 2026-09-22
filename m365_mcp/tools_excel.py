"""Excel tools (COM).

Covers the day-to-day surface: workbooks, sheets, ranges, formulas, formatting,
names, charts, macros, PDF export. Anything not covered here is reachable with
`com_eval` / `com_call` from tools_common.
"""

from __future__ import annotations

import os
from typing import Any

from .comcore import (
    ComToolError,
    com_tool,
    drop_handle,
    find_handle_by,
    get_app,
    get_handle,
    handle_meta,
    jsonable,
    list_handles,
    put_handle,
)

# --- Excel enums we rely on (dynamic dispatch has no win32com constants) ---
XL_TYPE_PDF = 0
XL_UP = -4162
XL_TO_LEFT = -4159
XL_CELL_TYPE_LAST_CELL = 11
XL_PART = 2
XL_WHOLE = 1
XL_VALUES = -4163
XL_FORMULAS = -4123

FILE_FORMATS = {
    "xlsx": 51,
    "xlsm": 52,
    "xlsb": 50,
    "xls": 56,
    "csv": 6,
    "txt": -4158,
    "html": 44,
    "pdf": None,  # handled by ExportAsFixedFormat
}

CHART_TYPES = {
    "column": 51,
    "column_stacked": 52,
    "bar": 57,
    "line": 4,
    "line_markers": 65,
    "pie": 5,
    "doughnut": -4120,
    "scatter": -4169,
    "scatter_lines": 74,
    "area": 1,
    "radar": -4151,
    "combo": 51,
}

MAX_CELLS_DEFAULT = 5000


# --------------------------------------------------------------------------
# helpers (all run inside the COM thread)
# --------------------------------------------------------------------------


def _xl(new_instance: bool = False, visible: bool | None = None) -> Any:
    return get_app("excel", new_instance=new_instance, visible=visible)


def _register_wb(wb: Any) -> str:
    existing = find_handle_by("xlwb", "name", wb.Name)
    if existing:
        return existing
    try:
        path = wb.FullName
    except Exception:  # noqa: BLE001
        path = wb.Name
    return put_handle("xlwb", wb, name=wb.Name, path=path)


def _wb(handle: str | None) -> Any:
    """Resolve a workbook from a handle id, a workbook name, a path, or None."""
    app = _xl()
    if not handle:
        try:
            wb = app.ActiveWorkbook
        except Exception:  # noqa: BLE001
            wb = None
        if wb is None:
            raise ComToolError(
                "No active workbook",
                hint="Open one with excel_open / excel_new, or pass a handle.",
            )
        return wb
    if handle.startswith("xlwb:"):
        return get_handle(handle)
    target = handle.strip().lower()
    for wb in app.Workbooks:
        if wb.Name.lower() == target:
            return wb
        try:
            if wb.FullName.lower() == target or os.path.abspath(
                wb.FullName
            ).lower() == os.path.abspath(handle).lower():
                return wb
        except Exception:  # noqa: BLE001
            continue
    raise ComToolError(
        "No open workbook matches %r" % handle,
        hint="Call excel_list_workbooks to see what is open.",
    )


def _sheet(wb: Any, sheet: str | int | None) -> Any:
    if sheet is None or sheet == "":
        return wb.ActiveSheet
    if isinstance(sheet, int):
        return wb.Worksheets(sheet)
    if isinstance(sheet, str) and sheet.isdigit():
        return wb.Worksheets(int(sheet))
    try:
        return wb.Worksheets(sheet)
    except Exception as exc:  # noqa: BLE001
        names = [ws.Name for ws in wb.Worksheets]
        raise ComToolError(
            "No sheet named %r" % sheet, hint="Sheets: " + ", ".join(names)
        ) from exc


def _range(ws: Any, a1: str | None) -> Any:
    if not a1:
        return ws.UsedRange
    try:
        return ws.Range(a1)
    except Exception as exc:  # noqa: BLE001
        raise ComToolError("Bad range %r for sheet %r" % (a1, ws.Name)) from exc


def _addr(rng: Any, absolute: bool = False) -> str:
    """Address is a parameterised property: late-bound COM hands back a plain
    string, so rng.Address(False, False) raises TypeError. Handle both shapes."""
    try:
        return str(rng.Address(absolute, absolute))
    except TypeError:
        text = str(rng.Address)
        return text if absolute else text.replace("$", "")


def _resize(rng: Any, rows: int, cols: int) -> Any:
    """Grow a range to rows x cols, anchored at its top-left cell.

    `Range.Resize` is a parameterised property, and late binding evaluates the
    property first: rng.Resize(2, 3) becomes rng.Resize (-> the same range) and
    then Item(2, 3) -> the SINGLE cell two rows down and three columns across.
    No error, just data written to the wrong place. So build the block from its
    corners instead; Worksheet.Cells(r, c) resolves to Item(r, c), which is what
    we actually want.
    """
    ws = rng.Worksheet
    top, left = int(rng.Row), int(rng.Column)
    return ws.Range(
        ws.Cells(top, left), ws.Cells(top + int(rows) - 1, left + int(cols) - 1)
    )


def _grid(value: Any) -> list[list[Any]]:
    """Range.Value/Value2 -> list of rows (a single cell comes back scalar)."""
    if value is None:
        return [[None]]
    if isinstance(value, tuple):
        if value and isinstance(value[0], tuple):
            return [[jsonable(c) for c in row] for row in value]
        return [[jsonable(c) for c in value]]
    return [[jsonable(value)]]


def _to_com_grid(values: Any) -> list[list[Any]]:
    if not isinstance(values, list):
        values = [[values]]
    elif not values:
        raise ComToolError("values is empty")
    elif not isinstance(values[0], list):
        values = [values]
    width = max(len(r) for r in values)
    return [list(r) + [None] * (width - len(r)) for r in values]


# --------------------------------------------------------------------------
# workbooks
# --------------------------------------------------------------------------


@com_tool
def excel_list_workbooks() -> dict[str, Any]:
    """List the workbooks currently open in Excel, with handles to address them."""
    app = _xl()
    rows = []
    for wb in app.Workbooks:
        try:
            saved = bool(wb.Saved)
        except Exception:  # noqa: BLE001
            saved = None
        rows.append(
            {
                "handle": _register_wb(wb),
                "name": wb.Name,
                "path": getattr(wb, "FullName", wb.Name),
                "saved": saved,
                "sheets": [ws.Name for ws in wb.Worksheets],
            }
        )
    return {"count": len(rows), "workbooks": rows, "visible": bool(app.Visible)}


@com_tool
def excel_open(
    path: str,
    read_only: bool = False,
    visible: bool = True,
    update_links: bool = False,
    password: str | None = None,
) -> dict[str, Any]:
    """Open a workbook (or return the handle if it is already open).

    Args:
        path: full path to .xlsx/.xlsm/.csv/...
        read_only: open without locking the file for others
        visible: show the Excel window
        update_links: refresh external links on open (default off, avoids prompts)
        password: open password, if the file is protected
    """
    full = os.path.abspath(os.path.expanduser(path))
    if not os.path.exists(full):
        raise ComToolError("File not found: " + full)
    app = _xl(visible=visible)
    for wb in app.Workbooks:
        try:
            if os.path.abspath(wb.FullName).lower() == full.lower():
                return {
                    "handle": _register_wb(wb),
                    "name": wb.Name,
                    "path": full,
                    "already_open": True,
                }
        except Exception:  # noqa: BLE001
            continue
    kwargs: dict[str, Any] = {
        "Filename": full,
        "UpdateLinks": 3 if update_links else 0,
        "ReadOnly": bool(read_only),
    }
    if password:
        kwargs["Password"] = password
    wb = app.Workbooks.Open(**kwargs)
    return {
        "handle": _register_wb(wb),
        "name": wb.Name,
        "path": getattr(wb, "FullName", full),
        "sheets": [ws.Name for ws in wb.Worksheets],
        "already_open": False,
    }


@com_tool
def excel_new(visible: bool = True, template: str | None = None) -> dict[str, Any]:
    """Create a new, unsaved workbook (optionally from a template path)."""
    app = _xl(visible=visible)
    wb = app.Workbooks.Add(os.path.abspath(template)) if template else app.Workbooks.Add()
    return {"handle": _register_wb(wb), "name": wb.Name}


@com_tool
def excel_save(
    handle: str | None = None,
    path: str | None = None,
    file_format: str | None = None,
) -> dict[str, Any]:
    """Save a workbook. With `path`, saves a copy under the new name/format.

    file_format: xlsx | xlsm | xlsb | xls | csv | txt | html (inferred from path).
    """
    wb = _wb(handle)
    if not path:
        wb.Save()
        return {"saved": True, "path": getattr(wb, "FullName", wb.Name)}
    full = os.path.abspath(os.path.expanduser(path))
    os.makedirs(os.path.dirname(full) or ".", exist_ok=True)
    fmt_key = (file_format or os.path.splitext(full)[1].lstrip(".")).lower()
    fmt = FILE_FORMATS.get(fmt_key)
    if fmt is None:
        raise ComToolError(
            "Unsupported format %r" % fmt_key,
            hint="Use one of " + ", ".join(k for k, v in FILE_FORMATS.items() if v),
        )
    wb.SaveAs(Filename=full, FileFormat=fmt)
    return {"saved": True, "path": full, "format": fmt_key}


@com_tool
def excel_close(handle: str | None = None, save: bool = False) -> dict[str, Any]:
    """Close a workbook. save=False discards unsaved changes."""
    wb = _wb(handle)
    name = wb.Name
    wb.Close(SaveChanges=bool(save))
    if handle and handle.startswith("xlwb:"):
        drop_handle(handle)
    else:
        hid = find_handle_by("xlwb", "name", name)
        if hid:
            drop_handle(hid)
    return {"closed": name, "saved": bool(save)}


@com_tool
def excel_export_pdf(
    output_path: str,
    handle: str | None = None,
    sheet: str | None = None,
    open_after: bool = False,
) -> dict[str, Any]:
    """Export a workbook (or one sheet) to PDF."""
    wb = _wb(handle)
    full = os.path.abspath(os.path.expanduser(output_path))
    os.makedirs(os.path.dirname(full) or ".", exist_ok=True)
    target = _sheet(wb, sheet) if sheet else wb
    target.ExportAsFixedFormat(
        Type=XL_TYPE_PDF, Filename=full, OpenAfterPublish=bool(open_after)
    )
    return {"pdf": full, "exists": os.path.exists(full)}


# --------------------------------------------------------------------------
# sheets
# --------------------------------------------------------------------------


@com_tool
def excel_list_sheets(handle: str | None = None) -> dict[str, Any]:
    """List worksheets with their used-range dimensions."""
    wb = _wb(handle)
    rows = []
    for ws in wb.Worksheets:
        used = ws.UsedRange
        rows.append(
            {
                "name": ws.Name,
                "index": ws.Index,
                "visible": int(ws.Visible),
                "used_range": _addr(used),
                "rows": int(used.Rows.Count),
                "columns": int(used.Columns.Count),
            }
        )
    return {"workbook": wb.Name, "count": len(rows), "sheets": rows}


@com_tool
def excel_add_sheet(
    name: str | None = None,
    handle: str | None = None,
    after: str | None = None,
) -> dict[str, Any]:
    """Add a worksheet, optionally named and positioned after an existing sheet."""
    wb = _wb(handle)
    ws = (
        wb.Worksheets.Add(After=_sheet(wb, after))
        if after
        else wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
    )
    if name:
        ws.Name = name
    return {"sheet": ws.Name, "index": int(ws.Index)}


@com_tool
def excel_delete_sheet(sheet: str, handle: str | None = None) -> dict[str, Any]:
    """Delete a worksheet (no confirmation prompt - DisplayAlerts is off)."""
    wb = _wb(handle)
    ws = _sheet(wb, sheet)
    name = ws.Name
    ws.Delete()
    return {"deleted": name}


@com_tool
def excel_rename_sheet(sheet: str, new_name: str, handle: str | None = None) -> dict:
    """Rename a worksheet."""
    wb = _wb(handle)
    ws = _sheet(wb, sheet)
    old = ws.Name
    ws.Name = new_name
    return {"renamed": old, "to": new_name}


@com_tool
def excel_copy_sheet(
    sheet: str,
    handle: str | None = None,
    new_name: str | None = None,
    to_workbook: str | None = None,
) -> dict[str, Any]:
    """Duplicate a worksheet inside the workbook, or copy it into another one."""
    wb = _wb(handle)
    ws = _sheet(wb, sheet)
    if to_workbook:
        dest = _wb(to_workbook)
        ws.Copy(After=dest.Worksheets(dest.Worksheets.Count))
        new = dest.ActiveSheet
    else:
        ws.Copy(After=wb.Worksheets(wb.Worksheets.Count))
        new = wb.ActiveSheet
    if new_name:
        new.Name = new_name
    return {"copied": ws.Name, "new_sheet": new.Name}


@com_tool
def excel_activate(
    handle: str | None = None, sheet: str | None = None, cell: str | None = None
) -> dict[str, Any]:
    """Bring a workbook/sheet/cell to the front (useful before a screenshot)."""
    wb = _wb(handle)
    wb.Activate()
    if sheet:
        ws = _sheet(wb, sheet)
        ws.Activate()
        if cell:
            ws.Range(cell).Select()
    return {"workbook": wb.Name, "sheet": wb.ActiveSheet.Name}


# --------------------------------------------------------------------------
# ranges
# --------------------------------------------------------------------------


@com_tool
def excel_read_range(
    handle: str | None = None,
    sheet: str | None = None,
    range_a1: str | None = None,
    mode: str = "values",
    max_cells: int = MAX_CELLS_DEFAULT,
    include_header: bool = False,
) -> dict[str, Any]:
    """Read a range. Omit range_a1 to read the sheet's used range.

    Args:
        mode: values (raw) | display (as formatted on screen) | formulas
        max_cells: safety cap; the result is truncated and flagged
        include_header: also return the first row as a `header` list
    """
    wb = _wb(handle)
    ws = _sheet(wb, sheet)
    rng = _range(ws, range_a1)
    n_rows, n_cols = int(rng.Rows.Count), int(rng.Columns.Count)
    truncated = False
    if n_rows * n_cols > max_cells:
        keep = max(1, max_cells // max(1, n_cols))
        rng = _resize(rng, keep, n_cols)
        truncated = True
    if mode == "display":
        data = _grid(rng.Text if n_rows * n_cols == 1 else rng.Value)
        if n_rows * n_cols != 1:
            data = [
                [
                    jsonable(rng.Cells(r + 1, c + 1).Text)
                    for c in range(int(rng.Columns.Count))
                ]
                for r in range(int(rng.Rows.Count))
            ]
    elif mode == "formulas":
        data = _grid(rng.Formula)
    else:
        data = _grid(rng.Value)
    out: dict[str, Any] = {
        "workbook": wb.Name,
        "sheet": ws.Name,
        "address": _addr(rng),
        "rows": len(data),
        "columns": len(data[0]) if data else 0,
        "truncated": truncated,
        "total_rows": n_rows,
        "values": data,
    }
    if include_header and data:
        out["header"] = data[0]
        out["values"] = data[1:]
        out["rows"] = len(out["values"])
    return out


@com_tool
def excel_write_range(
    values: list,
    handle: str | None = None,
    sheet: str | None = None,
    start_cell: str = "A1",
    as_formula: bool = False,
    number_format: str | None = None,
) -> dict[str, Any]:
    """Write a 2-D array (list of rows) starting at start_cell.

    Args:
        values: [[a, b], [c, d]] - a flat list is treated as one row
        as_formula: write strings as formulas (they must start with "=")
        number_format: apply a format to the written block, e.g. "#,##0.00"
    """
    wb = _wb(handle)
    ws = _sheet(wb, sheet)
    grid = _to_com_grid(values)
    rows, cols = len(grid), len(grid[0])
    rng = _resize(ws.Range(start_cell), rows, cols)
    payload = tuple(tuple(r) for r in grid)
    if as_formula:
        rng.Formula = payload
    else:
        rng.Value = payload
    if number_format:
        rng.NumberFormat = number_format
    return {
        "sheet": ws.Name,
        "address": _addr(rng),
        "cells_written": rows * cols,
    }


@com_tool
def excel_append_rows(
    values: list,
    handle: str | None = None,
    sheet: str | None = None,
    column: str = "A",
) -> dict[str, Any]:
    """Append rows below the last used row of `column` (table-style logging)."""
    wb = _wb(handle)
    ws = _sheet(wb, sheet)
    last = ws.Cells(ws.Rows.Count, ws.Range(column + "1").Column).End(XL_UP).Row
    used_empty = _addr(ws.UsedRange) == "A1" and not ws.Range("A1").Value
    start_row = 1 if used_empty else int(last) + 1
    grid = _to_com_grid(values)
    rng = _resize(
        ws.Cells(start_row, ws.Range(column + "1").Column), len(grid), len(grid[0])
    )
    rng.Value = tuple(tuple(r) for r in grid)
    return {"sheet": ws.Name, "address": _addr(rng), "rows": len(grid)}


@com_tool
def excel_clear_range(
    range_a1: str,
    handle: str | None = None,
    sheet: str | None = None,
    what: str = "all",
) -> dict[str, Any]:
    """Clear a range. what: all | contents | formats | comments."""
    wb = _wb(handle)
    ws = _sheet(wb, sheet)
    rng = _range(ws, range_a1)
    method = {
        "all": rng.Clear,
        "contents": rng.ClearContents,
        "formats": rng.ClearFormats,
        "comments": rng.ClearComments,
    }.get(what.lower())
    if method is None:
        raise ComToolError("what must be all|contents|formats|comments")
    method()
    return {"cleared": _addr(rng), "what": what}


@com_tool
def excel_format_range(
    range_a1: str,
    handle: str | None = None,
    sheet: str | None = None,
    number_format: str | None = None,
    bold: bool | None = None,
    italic: bool | None = None,
    font_size: float | None = None,
    font_name: str | None = None,
    font_color: str | None = None,
    fill_color: str | None = None,
    horizontal_align: str | None = None,
    wrap_text: bool | None = None,
    borders: bool | None = None,
    column_width: float | None = None,
    row_height: float | None = None,
    merge: bool | None = None,
) -> dict[str, Any]:
    """Format a range. Colors are "#RRGGBB"; align is left|center|right."""
    wb = _wb(handle)
    ws = _sheet(wb, sheet)
    rng = _range(ws, range_a1)

    def _bgr(hex_color: str) -> int:
        h = hex_color.lstrip("#")
        if len(h) != 6:
            raise ComToolError("Color must look like #RRGGBB, got %r" % hex_color)
        r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
        return b * 65536 + g * 256 + r  # Excel stores colors as BGR

    applied = []
    if number_format is not None:
        rng.NumberFormat = number_format
        applied.append("number_format")
    if bold is not None:
        rng.Font.Bold = bool(bold)
        applied.append("bold")
    if italic is not None:
        rng.Font.Italic = bool(italic)
        applied.append("italic")
    if font_size is not None:
        rng.Font.Size = float(font_size)
        applied.append("font_size")
    if font_name is not None:
        rng.Font.Name = font_name
        applied.append("font_name")
    if font_color is not None:
        rng.Font.Color = _bgr(font_color)
        applied.append("font_color")
    if fill_color is not None:
        rng.Interior.Color = _bgr(fill_color)
        applied.append("fill_color")
    if horizontal_align is not None:
        rng.HorizontalAlignment = {
            "left": -4131,
            "center": -4108,
            "right": -4152,
        }.get(horizontal_align.lower(), -4131)
        applied.append("align")
    if wrap_text is not None:
        rng.WrapText = bool(wrap_text)
        applied.append("wrap_text")
    if borders is not None and borders:
        for edge in (7, 8, 9, 10, 11, 12):  # left/top/bottom/right/inside v+h
            try:
                rng.Borders(edge).LineStyle = 1
                rng.Borders(edge).Weight = 2
            except Exception:  # noqa: BLE001 - inside borders fail on 1 cell
                pass
        applied.append("borders")
    if column_width is not None:
        rng.ColumnWidth = float(column_width)
        applied.append("column_width")
    if row_height is not None:
        rng.RowHeight = float(row_height)
        applied.append("row_height")
    if merge is not None:
        rng.Merge() if merge else rng.UnMerge()
        applied.append("merge")
    return {"range": _addr(rng), "applied": applied}


@com_tool
def excel_autofit(
    handle: str | None = None,
    sheet: str | None = None,
    range_a1: str | None = None,
    columns: bool = True,
    rows: bool = False,
) -> dict[str, Any]:
    """Autofit column widths and/or row heights."""
    wb = _wb(handle)
    ws = _sheet(wb, sheet)
    rng = _range(ws, range_a1)
    if columns:
        rng.Columns.AutoFit()
    if rows:
        rng.Rows.AutoFit()
    return {"sheet": ws.Name, "range": _addr(rng)}


@com_tool
def excel_insert_delete(
    action: str,
    target: str,
    handle: str | None = None,
    sheet: str | None = None,
) -> dict[str, Any]:
    """Insert or delete whole rows/columns.

    Args:
        action: insert_rows | delete_rows | insert_columns | delete_columns
        target: "3:5" for rows, "B:C" for columns
    """
    wb = _wb(handle)
    ws = _sheet(wb, sheet)
    rng = ws.Range(target)
    act = action.lower()
    if act == "insert_rows":
        rng.EntireRow.Insert()
    elif act == "delete_rows":
        rng.EntireRow.Delete()
    elif act == "insert_columns":
        rng.EntireColumn.Insert()
    elif act == "delete_columns":
        rng.EntireColumn.Delete()
    else:
        raise ComToolError(
            "action must be insert_rows|delete_rows|insert_columns|delete_columns"
        )
    return {"action": act, "target": target, "sheet": ws.Name}


@com_tool
def excel_find(
    what: str,
    handle: str | None = None,
    sheet: str | None = None,
    search_all_sheets: bool = False,
    match_case: bool = False,
    whole_cell: bool = False,
    look_in: str = "values",
    max_results: int = 50,
) -> dict[str, Any]:
    """Find cells containing a value/formula text. look_in: values | formulas."""
    wb = _wb(handle)
    sheets = list(wb.Worksheets) if search_all_sheets else [_sheet(wb, sheet)]
    look = XL_FORMULAS if look_in.lower().startswith("form") else XL_VALUES
    hits: list[dict[str, Any]] = []
    for ws in sheets:
        try:
            first = ws.UsedRange.Find(
                What=what,
                LookIn=look,
                LookAt=XL_WHOLE if whole_cell else XL_PART,
                MatchCase=bool(match_case),
            )
        except Exception:  # noqa: BLE001
            first = None
        if first is None:
            continue
        addr0 = _addr(first)
        cur = first
        while True:
            hits.append(
                {
                    "sheet": ws.Name,
                    "cell": _addr(cur),
                    "value": jsonable(cur.Value),
                    "formula": jsonable(cur.Formula),
                }
            )
            if len(hits) >= max_results:
                return {"query": what, "count": len(hits), "hits": hits,
                        "truncated": True}
            cur = ws.UsedRange.FindNext(cur)
            if cur is None or _addr(cur) == addr0:
                break
    return {"query": what, "count": len(hits), "hits": hits, "truncated": False}


@com_tool
def excel_replace(
    find_text: str,
    replace_text: str,
    handle: str | None = None,
    sheet: str | None = None,
    range_a1: str | None = None,
    match_case: bool = False,
    whole_cell: bool = False,
) -> dict[str, Any]:
    """Find & replace inside a sheet (or a range)."""
    wb = _wb(handle)
    ws = _sheet(wb, sheet)
    rng = _range(ws, range_a1)
    changed = rng.Replace(
        What=find_text,
        Replacement=replace_text,
        LookAt=XL_WHOLE if whole_cell else XL_PART,
        MatchCase=bool(match_case),
    )
    return {"sheet": ws.Name, "replaced": bool(changed)}


@com_tool
def excel_cell_info(
    cell: str, handle: str | None = None, sheet: str | None = None
) -> dict[str, Any]:
    """Everything about one cell: value, formula, format, precedents, comment."""
    wb = _wb(handle)
    ws = _sheet(wb, sheet)
    rng = ws.Range(cell)
    info: dict[str, Any] = {
        "sheet": ws.Name,
        "address": _addr(rng),
        "value": jsonable(rng.Value),
        "text": jsonable(rng.Text),
        "formula": jsonable(rng.Formula),
        "number_format": jsonable(rng.NumberFormat),
        "has_formula": bool(rng.HasFormula),
        "locked": bool(rng.Locked),
    }
    for label, getter in (
        ("comment", lambda: rng.Comment.Text()),
        ("note", lambda: rng.NoteText()),
        ("validation", lambda: rng.Validation.Formula1),
    ):
        try:
            info[label] = jsonable(getter())
        except Exception:  # noqa: BLE001 - absent features raise
            pass
    if info["has_formula"]:
        try:
            info["precedents"] = _addr(rng.Precedents)
        except Exception:  # noqa: BLE001
            info["precedents"] = None
    return info


# --------------------------------------------------------------------------
# names, calculation, macros, charts
# --------------------------------------------------------------------------


@com_tool
def excel_list_names(handle: str | None = None) -> dict[str, Any]:
    """List defined names (named ranges) of a workbook."""
    wb = _wb(handle)
    rows = []
    for nm in wb.Names:
        try:
            rows.append(
                {
                    "name": nm.Name,
                    "refers_to": nm.RefersTo,
                    "visible": bool(nm.Visible),
                }
            )
        except Exception:  # noqa: BLE001
            continue
    return {"workbook": wb.Name, "count": len(rows), "names": rows}


@com_tool
def excel_set_name(
    name: str, refers_to: str, handle: str | None = None
) -> dict[str, Any]:
    """Create/replace a defined name. refers_to looks like "=Sheet1!$A$1:$B$9"."""
    wb = _wb(handle)
    try:
        wb.Names(name).Delete()
    except Exception:  # noqa: BLE001 - not defined yet
        pass
    wb.Names.Add(Name=name, RefersTo=refers_to)
    return {"name": name, "refers_to": refers_to}


@com_tool
def excel_calculate(
    handle: str | None = None, full: bool = False, mode: str | None = None
) -> dict[str, Any]:
    """Recalculate. full=True forces a full rebuild; mode sets auto|manual."""
    app = _xl()
    if mode:
        app.Calculation = {"auto": -4105, "automatic": -4105, "manual": -4135}.get(
            mode.lower(), -4105
        )
    if full:
        app.CalculateFullRebuild()
    else:
        app.Calculate()
    return {"calculated": True, "mode": mode or "unchanged"}


@com_tool
def excel_run_macro(
    macro: str, args: list | None = None, handle: str | None = None
) -> dict[str, Any]:
    """Run a VBA macro/function by name (e.g. "Module1.Refresh" or "MyMacro")."""
    app = _xl()
    if handle:
        _wb(handle).Activate()
    result = app.Run(macro, *(args or []))
    return {"macro": macro, "result": jsonable(result)}


@com_tool
def excel_add_chart(
    data_range: str,
    handle: str | None = None,
    sheet: str | None = None,
    chart_type: str = "column",
    title: str | None = None,
    left: float = 350.0,
    top: float = 20.0,
    width: float = 480.0,
    height: float = 300.0,
    target_sheet: str | None = None,
) -> dict[str, Any]:
    """Add a chart object from a data range.

    chart_type: column | bar | line | line_markers | pie | doughnut | scatter |
    scatter_lines | area | radar
    """
    wb = _wb(handle)
    ws = _sheet(wb, sheet)
    dest = _sheet(wb, target_sheet) if target_sheet else ws
    kind = CHART_TYPES.get(chart_type.lower())
    if kind is None:
        raise ComToolError(
            "Unknown chart_type %r" % chart_type,
            hint="One of: " + ", ".join(sorted(CHART_TYPES)),
        )
    shape = dest.Shapes.AddChart2(-1, kind, left, top, width, height)
    chart = shape.Chart
    chart.SetSourceData(Source=ws.Range(data_range))
    if title:
        chart.HasTitle = True
        chart.ChartTitle.Text = title
    return {
        "sheet": dest.Name,
        "chart_name": shape.Name,
        "chart_type": chart_type,
        "source": data_range,
    }


@com_tool
def excel_list_charts(handle: str | None = None, sheet: str | None = None) -> dict:
    """List chart objects on a sheet (or the whole workbook)."""
    wb = _wb(handle)
    sheets = [_sheet(wb, sheet)] if sheet else list(wb.Worksheets)
    rows = []
    for ws in sheets:
        for co in ws.ChartObjects():
            rows.append(
                {
                    "sheet": ws.Name,
                    "name": co.Name,
                    "title": (
                        co.Chart.ChartTitle.Text if co.Chart.HasTitle else None
                    ),
                }
            )
    return {"count": len(rows), "charts": rows}


@com_tool
def excel_export_chart(
    output_path: str,
    chart_name: str,
    handle: str | None = None,
    sheet: str | None = None,
) -> dict[str, Any]:
    """Export one chart to a PNG file."""
    wb = _wb(handle)
    ws = _sheet(wb, sheet)
    full = os.path.abspath(os.path.expanduser(output_path))
    os.makedirs(os.path.dirname(full) or ".", exist_ok=True)
    ws.ChartObjects(chart_name).Chart.Export(full)
    return {"png": full, "exists": os.path.exists(full)}


@com_tool
def excel_list_open_handles() -> dict[str, Any]:
    """Diagnostics: workbook handles this server has handed out."""
    return {"handles": list_handles("xlwb")}


TOOLS = [
    excel_list_workbooks,
    excel_open,
    excel_new,
    excel_save,
    excel_close,
    excel_export_pdf,
    excel_list_sheets,
    excel_add_sheet,
    excel_delete_sheet,
    excel_rename_sheet,
    excel_copy_sheet,
    excel_activate,
    excel_read_range,
    excel_write_range,
    excel_append_rows,
    excel_clear_range,
    excel_format_range,
    excel_autofit,
    excel_insert_delete,
    excel_find,
    excel_replace,
    excel_cell_info,
    excel_list_names,
    excel_set_name,
    excel_calculate,
    excel_run_macro,
    excel_add_chart,
    excel_list_charts,
    excel_export_chart,
    excel_list_open_handles,
]
