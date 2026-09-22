"""OneNote tools (COM, desktop OneNote 2016 / Application.15).

The OneNote object model is XML in / XML out: GetHierarchy, GetPageContent,
UpdatePageContent. These tools hide the XML for the common cases and still let
you fetch the raw XML when you need it.

The OneNote store app ("OneNote for Windows 10") exposes no COM - only the
desktop app does.
"""

from __future__ import annotations

import os
import xml.etree.ElementTree as ET
from typing import Any

from .comcore import ComToolError, com_tool, get_app

ONE_NS = "{http://schemas.microsoft.com/office/onenote/2013/onenote}"

# HierarchyScope
SCOPES = {"self": 0, "children": 1, "notebooks": 2, "sections": 3, "pages": 4}
# PublishFormat
PUBLISH = {"one": 0, "onea": 1, "docx": 2, "pdf": 3, "xps": 4, "mht": 5, "emf": 6}


def _one() -> Any:
    return get_app("onenote")


def _hierarchy_xml(start: str | None, scope: str) -> str:
    app = _one()
    code = SCOPES.get(scope.lower())
    if code is None:
        raise ComToolError("scope must be one of " + ", ".join(SCOPES))
    return app.GetHierarchy(start or "", code)


def _node_rows(xml_text: str) -> list[dict[str, Any]]:
    root = ET.fromstring(xml_text)
    rows: list[dict[str, Any]] = []

    def walk(node: ET.Element, depth: int, path: str) -> None:
        for child in node:
            tag = child.tag.replace(ONE_NS, "")
            if tag not in ("Notebook", "SectionGroup", "Section", "Page"):
                continue
            name = child.get("name") or child.get("title") or ""
            full = (path + "/" + name) if path else name
            rows.append(
                {
                    "type": tag.lower(),
                    "name": name,
                    "path": full,
                    "id": child.get("ID"),
                    "depth": depth,
                    "last_modified": child.get("lastModifiedTime"),
                    "page_level": child.get("pageLevel"),
                }
            )
            walk(child, depth + 1, full)

    walk(root, 0, "")
    return rows


@com_tool
def onenote_hierarchy(
    start_id: str | None = None, scope: str = "pages", max_results: int = 400
) -> dict[str, Any]:
    """List notebooks / section groups / sections / pages.

    scope: self | children | notebooks | sections | pages (pages = full tree)
    """
    xml_text = _hierarchy_xml(start_id, scope)
    rows = _node_rows(xml_text)
    return {
        "scope": scope,
        "count": len(rows),
        "truncated": len(rows) > max_results,
        "nodes": rows[:max_results],
    }


@com_tool
def onenote_find_pages(
    query: str, start_id: str | None = None, max_results: int = 50
) -> dict[str, Any]:
    """Full-text search across OneNote and return matching pages."""
    app = _one()
    xml_text = app.FindPages(start_id or "", query)
    rows = [r for r in _node_rows(xml_text) if r["type"] == "page"]
    return {"query": query, "count": len(rows), "pages": rows[:max_results]}


@com_tool
def onenote_get_page(
    page_id: str, as_text: bool = True, max_chars: int = 30000
) -> dict[str, Any]:
    """Read a page. as_text=True flattens the XML to plain text."""
    app = _one()
    xml_text = app.GetPageContent(page_id)
    if not as_text:
        return {
            "page_id": page_id,
            "xml": xml_text[:max_chars],
            "truncated": len(xml_text) > max_chars,
        }
    root = ET.fromstring(xml_text)
    title = None
    lines: list[str] = []
    for el in root.iter():
        tag = el.tag.replace(ONE_NS, "")
        if tag == "T" and el.text:
            text = ET.fromstring("<x>%s</x>" % el.text.replace("&nbsp;", " ")) \
                if "<" in el.text else None
            lines.append("".join(text.itertext()) if text is not None else el.text)
    if root.get("name"):
        title = root.get("name")
    body = "\n".join(l for l in lines if l and l.strip())
    return {
        "page_id": page_id,
        "title": title,
        "chars": len(body),
        "truncated": len(body) > max_chars,
        "text": body[:max_chars],
    }


@com_tool
def onenote_create_page(
    section_id: str, title: str, content: str | list | None = None
) -> dict[str, Any]:
    """Create a page in a section; `content` may be a string or a list of lines."""
    app = _one()
    page_id = app.CreateNewPage(section_id)
    lines = content if isinstance(content, list) else ([content] if content else [])
    outlines = "".join(
        "<one:OE><one:T><![CDATA[%s]]></one:T></one:OE>" % str(line) for line in lines
    )
    body = (
        "<one:Outline><one:OEChildren>%s</one:OEChildren></one:Outline>" % outlines
        if outlines
        else ""
    )
    xml_text = (
        '<?xml version="1.0"?>'
        '<one:Page xmlns:one="http://schemas.microsoft.com/office/onenote/2013/onenote"'
        ' ID="%s">'
        '<one:Title><one:OE><one:T><![CDATA[%s]]></one:T></one:OE></one:Title>'
        "%s</one:Page>" % (page_id, title, body)
    )
    app.UpdatePageContent(xml_text)
    return {"page_id": page_id, "title": title, "lines": len(lines)}


@com_tool
def onenote_append_to_page(page_id: str, text: str | list) -> dict[str, Any]:
    """Append one or more paragraphs to an existing page."""
    app = _one()
    lines = text if isinstance(text, list) else [text]
    outlines = "".join(
        "<one:OE><one:T><![CDATA[%s]]></one:T></one:OE>" % str(line) for line in lines
    )
    xml_text = (
        '<?xml version="1.0"?>'
        '<one:Page xmlns:one="http://schemas.microsoft.com/office/onenote/2013/onenote"'
        ' ID="%s">'
        "<one:Outline><one:OEChildren>%s</one:OEChildren></one:Outline>"
        "</one:Page>" % (page_id, outlines)
    )
    app.UpdatePageContent(xml_text)
    return {"page_id": page_id, "appended_lines": len(lines)}


@com_tool
def onenote_update_page_xml(page_xml: str) -> dict[str, Any]:
    """Escape hatch: push raw OneNote page XML through UpdatePageContent."""
    _one().UpdatePageContent(page_xml)
    return {"updated": True, "bytes": len(page_xml)}


@com_tool
def onenote_export(
    node_id: str, output_path: str, export_format: str = "pdf"
) -> dict[str, Any]:
    """Publish a page/section/notebook to a file.

    export_format: pdf | docx | one | onea | xps | mht | emf
    """
    app = _one()
    code = PUBLISH.get(export_format.lower())
    if code is None:
        raise ComToolError("export_format must be one of " + ", ".join(PUBLISH))
    full = os.path.abspath(os.path.expanduser(output_path))
    os.makedirs(os.path.dirname(full) or ".", exist_ok=True)
    if os.path.exists(full):
        os.remove(full)  # Publish refuses to overwrite
    app.Publish(node_id, full, code, "")
    return {"path": full, "exists": os.path.exists(full), "format": export_format}


@com_tool
def onenote_navigate(node_id: str) -> dict[str, Any]:
    """Bring a page/section into view in the OneNote window."""
    _one().NavigateTo(node_id, "", False)
    return {"navigated_to": node_id}


TOOLS = [
    onenote_hierarchy,
    onenote_find_pages,
    onenote_get_page,
    onenote_create_page,
    onenote_append_to_page,
    onenote_update_page_xml,
    onenote_export,
    onenote_navigate,
]

