"""Access and Publisher tools (COM / ACE-OLEDB).

Access data access goes through ADODB (ACE OLEDB) rather than the Access UI:
it is faster, needs no window, and works when Access itself is busy. Macros and
UI-level actions still need Access.Application.
"""

from __future__ import annotations

import os
from typing import Any

import win32com.client

from .comcore import (
    ComToolError,
    com_tool,
    drop_handle,
    find_handle_by,
    get_app,
    get_handle,
    jsonable,
    put_handle,
)

ACE_PROVIDERS = ("Microsoft.ACE.OLEDB.16.0", "Microsoft.ACE.OLEDB.12.0",
                 "Microsoft.Jet.OLEDB.4.0")

AD_SCHEMA_TABLES = 20


def _connect(db_path: str, password: str | None = None) -> Any:
    full = os.path.abspath(os.path.expanduser(db_path))
    if not os.path.exists(full):
        raise ComToolError("Database not found: " + full)
    last_error = None
    for provider in ACE_PROVIDERS:
        conn_str = "Provider=%s;Data Source=%s;" % (provider, full)
        if password:
            conn_str += "Jet OLEDB:Database Password=%s;" % password
        conn = win32com.client.Dispatch("ADODB.Connection")
        try:
            conn.Open(conn_str)
            return conn
        except Exception as exc:  # noqa: BLE001 - try the next provider
            last_error = exc
    raise ComToolError(
        "No usable ACE/Jet OLEDB provider for " + full,
        hint="Install the Microsoft Access Database Engine matching your Office "
             "bitness, or open the file with access_run_macro instead. Last error: "
             + str(last_error),
    )


def _rows_from(rs: Any, max_rows: int) -> tuple[list[str], list[list[Any]], bool]:
    columns = [str(rs.Fields(i).Name) for i in range(int(rs.Fields.Count))]
    rows: list[list[Any]] = []
    truncated = False
    while not rs.EOF:
        if len(rows) >= max_rows:
            truncated = True
            break
        rows.append([jsonable(rs.Fields(i).Value) for i in range(len(columns))])
        rs.MoveNext()
    return columns, rows, truncated


# --------------------------------------------------------------------------
# Access
# --------------------------------------------------------------------------


@com_tool
def access_list_objects(db_path: str, password: str | None = None) -> dict[str, Any]:
    """List tables, views and their column definitions in an .accdb/.mdb file."""
    conn = _connect(db_path, password)
    try:
        rs = conn.OpenSchema(AD_SCHEMA_TABLES)
        tables = []
        while not rs.EOF:
            name = str(rs.Fields("TABLE_NAME").Value)
            kind = str(rs.Fields("TABLE_TYPE").Value)
            if not name.startswith("MSys"):
                tables.append({"name": name, "type": kind})
            rs.MoveNext()
        rs.Close()
        for tbl in tables:
            if tbl["type"] not in ("TABLE", "VIEW"):
                continue
            try:
                probe = conn.Execute("SELECT * FROM [%s] WHERE 1=0" % tbl["name"])[0]
                tbl["columns"] = [
                    {
                        "name": str(probe.Fields(i).Name),
                        "type": int(probe.Fields(i).Type),
                    }
                    for i in range(int(probe.Fields.Count))
                ]
                probe.Close()
            except Exception:  # noqa: BLE001
                tbl["columns"] = []
        return {"database": os.path.abspath(db_path), "count": len(tables),
                "objects": tables}
    finally:
        conn.Close()


@com_tool
def access_query(
    db_path: str,
    sql: str,
    max_rows: int = 500,
    password: str | None = None,
) -> dict[str, Any]:
    """Run a SELECT (or a saved query name) and return rows."""
    statement = sql.strip()
    if not statement.lower().startswith(("select", "with", "transform", "parameters")):
        if " " not in statement:  # a saved query name
            statement = "SELECT * FROM [%s]" % statement
        else:
            raise ComToolError(
                "access_query only runs read statements",
                hint="Use access_execute for INSERT/UPDATE/DELETE/DDL.",
            )
    conn = _connect(db_path, password)
    try:
        rs = conn.Execute(statement)[0]
        columns, rows, truncated = _rows_from(rs, max_rows)
        rs.Close()
        return {
            "sql": statement,
            "columns": columns,
            "row_count": len(rows),
            "truncated": truncated,
            "rows": rows,
        }
    finally:
        conn.Close()


@com_tool
def access_execute(
    db_path: str,
    sql: str,
    confirm: bool = False,
    password: str | None = None,
) -> dict[str, Any]:
    """Run a data-modifying statement (INSERT/UPDATE/DELETE/DDL). Needs confirm=True."""
    if not confirm:
        return {
            "ok": False,
            "error": "confirm=False - nothing was executed",
            "sql": sql,
            "hint": "This writes to the database; get the user's approval first.",
        }
    conn = _connect(db_path, password)
    try:
        result = conn.Execute(sql)
        affected = None
        try:
            affected = int(result[1])
        except Exception:  # noqa: BLE001
            pass
        return {"sql": sql, "rows_affected": affected}
    finally:
        conn.Close()


@com_tool
def access_export(
    db_path: str,
    output_path: str,
    table_or_sql: str,
    max_rows: int = 100000,
    password: str | None = None,
) -> dict[str, Any]:
    """Export a table or SELECT result to .csv (UTF-8 with BOM for Excel)."""
    import csv

    statement = table_or_sql.strip()
    if " " not in statement:
        statement = "SELECT * FROM [%s]" % statement
    conn = _connect(db_path, password)
    try:
        rs = conn.Execute(statement)[0]
        columns, rows, truncated = _rows_from(rs, max_rows)
        rs.Close()
    finally:
        conn.Close()
    full = os.path.abspath(os.path.expanduser(output_path))
    os.makedirs(os.path.dirname(full) or ".", exist_ok=True)
    with open(full, "w", newline="", encoding="utf-8-sig") as fh:
        writer = csv.writer(fh)
        writer.writerow(columns)
        writer.writerows(rows)
    return {"csv": full, "rows": len(rows), "truncated": truncated}


@com_tool
def access_run_macro(
    db_path: str, macro: str, visible: bool = False
) -> dict[str, Any]:
    """Open the database in Access and run a macro or VBA procedure."""
    app = get_app("access", visible=visible)
    full = os.path.abspath(os.path.expanduser(db_path))
    if not os.path.exists(full):
        raise ComToolError("Database not found: " + full)
    try:
        current = str(app.CurrentDb().Name)
    except Exception:  # noqa: BLE001 - nothing open
        current = ""
    if os.path.abspath(current).lower() != full.lower():
        app.OpenCurrentDatabase(full)
    try:
        app.DoCmd.RunMacro(macro)
        kind = "macro"
    except Exception:  # noqa: BLE001 - fall back to a VBA procedure
        app.Run(macro)
        kind = "procedure"
    return {"database": full, "ran": macro, "kind": kind}


# --------------------------------------------------------------------------
# Publisher
# --------------------------------------------------------------------------


def _pubapp(visible: bool | None = None) -> Any:
    return get_app("publisher", visible=visible)


def _pub(handle: str | None) -> Any:
    app = _pubapp()
    if not handle:
        return app.ActiveDocument
    if handle.startswith("pbdoc:"):
        return get_handle(handle)
    for doc in app.Documents:
        if str(doc.Name).lower() == handle.strip().lower():
            return doc
    raise ComToolError("No open Publisher document matches %r" % handle)


@com_tool
def publisher_open(path: str, visible: bool = True) -> dict[str, Any]:
    """Open a .pub publication."""
    full = os.path.abspath(os.path.expanduser(path))
    if not os.path.exists(full):
        raise ComToolError("File not found: " + full)
    app = _pubapp(visible=visible)
    doc = app.Open(full)
    existing = find_handle_by("pbdoc", "name", str(doc.Name))
    hid = existing or put_handle("pbdoc", doc, name=str(doc.Name), path=full)
    return {"handle": hid, "name": str(doc.Name), "pages": int(doc.Pages.Count)}


@com_tool
def publisher_read_text(handle: str | None = None, max_chars: int = 20000) -> dict:
    """Extract the text of every story in a publication."""
    doc = _pub(handle)
    chunks = []
    for page in doc.Pages:
        for shape in page.Shapes:
            try:
                if shape.HasTextFrame:
                    chunks.append(str(shape.TextFrame.TextRange.Text))
            except Exception:  # noqa: BLE001
                continue
    text = "\n".join(c for c in chunks if c.strip())
    return {
        "name": str(doc.Name),
        "pages": int(doc.Pages.Count),
        "chars": len(text),
        "truncated": len(text) > max_chars,
        "text": text[:max_chars],
    }


@com_tool
def publisher_export_pdf(output_path: str, handle: str | None = None) -> dict:
    """Export a publication to PDF."""
    doc = _pub(handle)
    full = os.path.abspath(os.path.expanduser(output_path))
    os.makedirs(os.path.dirname(full) or ".", exist_ok=True)
    doc.ExportAsFixedFormat(1, full)  # pbFixedFormatTypePDF
    return {"pdf": full, "exists": os.path.exists(full)}


@com_tool
def publisher_close(handle: str | None = None, save: bool = False) -> dict:
    """Close a publication."""
    doc = _pub(handle)
    name = str(doc.Name)
    doc.Close()
    hid = handle if (handle or "").startswith("pbdoc:") else find_handle_by(
        "pbdoc", "name", name
    )
    if hid:
        drop_handle(hid)
    return {"closed": name, "saved": bool(save)}


TOOLS = [
    access_list_objects,
    access_query,
    access_execute,
    access_export,
    access_run_macro,
    publisher_open,
    publisher_read_text,
    publisher_export_pdf,
    publisher_close,
]
