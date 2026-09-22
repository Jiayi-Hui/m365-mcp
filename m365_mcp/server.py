"""m365-mcp - MCP server exposing Microsoft 365 desktop apps over COM.

Start with:
    python -m m365_mcp.server          # stdio transport

Tool families:
    m365_*      status, handles, conversion
    excel_*     workbooks, sheets, ranges, formats, charts, macros
    word_*      documents, text, tables, comments, revisions
    ppt_*       presentations, slides, shapes, notes, export
    outlook_*   mail, folders, search, calendar, contacts, tasks
    onenote_*   notebooks, pages, search, publish
    access_*    ADO queries against .accdb/.mdb, macros
    publisher_* publications
    com_*       generic object-model access (describe/get/set/call/eval)

Everything runs against the already-signed-in desktop applications on this
machine. No passwords, no Graph registration, no network.
"""

from __future__ import annotations

from typing import Any

from mcp.server.fastmcp import FastMCP

from . import (
    tools_common,
    tools_excel,
    tools_misc,
    tools_onenote,
    tools_outlook,
    tools_ppt,
    tools_word,
)

mcp = FastMCP("m365")

MODULES = (
    tools_common,
    tools_excel,
    tools_word,
    tools_ppt,
    tools_outlook,
    tools_onenote,
    tools_misc,
)

FAMILIES = {
    "m365_": "server + application status, handles, file conversion",
    "excel_": "workbooks, sheets, ranges, formulas, formatting, charts, macros",
    "word_": "documents, text, outline, tables, images, comments, revisions",
    "ppt_": "presentations, slides, shapes, tables, charts, notes, export",
    "outlook_": "mail, folders, search, drafts, calendar, contacts, tasks",
    "onenote_": "notebook hierarchy, page read/write, search, publish",
    "access_": "ADO queries and macros against .accdb/.mdb",
    "publisher_": "publications: read text, export PDF",
    "com_": "generic COM: describe / get / set / call / eval on any object",
    "office_": "cross-app file conversion",
}


def _registered_tools() -> list[str]:
    names = []
    for module in MODULES:
        names.extend(fn.__name__ for fn in module.TOOLS)
    return sorted(names)


def m365_help(family: str | None = None) -> dict[str, Any]:
    """Map of every tool this server exposes, grouped by application.

    Call this when you are not sure which tool fits, or to find out whether an
    Office feature is covered by a typed tool before falling back to com_eval.
    """
    names = _registered_tools()
    if family:
        prefix = family if family.endswith("_") else family + "_"
        picked = [n for n in names if n.startswith(prefix)]
        return {"ok": True, "family": prefix, "count": len(picked), "tools": picked}
    grouped = {
        prefix: {
            "description": desc,
            "tools": [n for n in names if n.startswith(prefix)],
        }
        for prefix, desc in FAMILIES.items()
    }
    return {
        "ok": True,
        "total_tools": len(names) + 1,
        "families": grouped,
        "notes": [
            "Handles (xlwb:1, wddoc:2, ...) address an open file across calls; "
            "most tools also accept a file name or path, or default to the "
            "active document.",
            "Tools that transmit or destroy something (outlook_send_draft, "
            "outlook_delete_message, access_execute, m365_quit_app) require "
            "confirm=true and the user's approval.",
            "com_describe + com_eval reach any part of the object model that "
            "has no typed tool.",
        ],
    }


def register(server: FastMCP | None = None) -> FastMCP:
    """Register every tool on a FastMCP instance (defaults to the module one)."""
    target = server or mcp
    target.tool()(m365_help)
    for module in MODULES:
        for fn in module.TOOLS:
            target.tool()(fn)
    return target


register()


def main() -> None:
    mcp.run()


if __name__ == "__main__":
    main()
