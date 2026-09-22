"""Word tools (COM).

Documents, text, styles, find/replace, tables, images, comments, revisions,
table of contents, macros, PDF export.
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
    jsonable,
    list_handles,
    put_handle,
)

# --- Word enums ---
WD_FORMAT = {
    "docx": 16,
    "doc": 0,
    "pdf": 17,
    "txt": 2,
    "rtf": 6,
    "html": 8,
    "xml": 12,
    "docm": 13,
}
WD_STORY = 6
WD_COLLAPSE_END = 0
WD_REPLACE_ALL = 2
WD_FIND_STOP = 0
WD_EXPORT_PDF = 17
WD_ALIGN = {"left": 0, "center": 1, "right": 2, "justify": 3}


def _wordapp(visible: bool | None = None) -> Any:
    return get_app("word", visible=visible)


def _register_doc(doc: Any) -> str:
    existing = find_handle_by("wddoc", "name", doc.Name)
    if existing:
        return existing
    try:
        path = doc.FullName
    except Exception:  # noqa: BLE001
        path = doc.Name
    return put_handle("wddoc", doc, name=doc.Name, path=path)


def _doc(handle: str | None) -> Any:
    app = _wordapp()
    if not handle:
        if int(app.Documents.Count) == 0:
            raise ComToolError(
                "No document open",
                hint="Use word_open / word_new, or pass a handle.",
            )
        return app.ActiveDocument
    if handle.startswith("wddoc:"):
        return get_handle(handle)
    target = handle.strip().lower()
    for doc in app.Documents:
        if doc.Name.lower() == target:
            return doc
        try:
            if os.path.abspath(doc.FullName).lower() == os.path.abspath(handle).lower():
                return doc
        except Exception:  # noqa: BLE001
            continue
    raise ComToolError("No open document matches %r" % handle)


def _doc_range(doc: Any, start_par: int | None, end_par: int | None) -> Any:
    if not start_par and not end_par:
        return doc.Content
    total = int(doc.Paragraphs.Count)
    s = max(1, int(start_par or 1))
    e = min(total, int(end_par or s))
    if s > total:
        raise ComToolError("Document has only %d paragraphs" % total)
    return doc.Range(doc.Paragraphs(s).Range.Start, doc.Paragraphs(e).Range.End)


# --------------------------------------------------------------------------
# documents
# --------------------------------------------------------------------------


@com_tool
def word_list_documents() -> dict[str, Any]:
    """List documents open in Word."""
    app = _wordapp()
    rows = []
    for doc in app.Documents:
        rows.append(
            {
                "handle": _register_doc(doc),
                "name": doc.Name,
                "path": getattr(doc, "FullName", doc.Name),
                "saved": bool(doc.Saved),
                "paragraphs": int(doc.Paragraphs.Count),
                "words": int(doc.Words.Count),
            }
        )
    return {"count": len(rows), "documents": rows}


@com_tool
def word_open(
    path: str, read_only: bool = False, visible: bool = True
) -> dict[str, Any]:
    """Open a .docx/.doc/.rtf/.txt document (returns the existing handle if open)."""
    full = os.path.abspath(os.path.expanduser(path))
    if not os.path.exists(full):
        raise ComToolError("File not found: " + full)
    app = _wordapp(visible=visible)
    for doc in app.Documents:
        try:
            if os.path.abspath(doc.FullName).lower() == full.lower():
                return {"handle": _register_doc(doc), "name": doc.Name,
                        "already_open": True}
        except Exception:  # noqa: BLE001
            continue
    doc = app.Documents.Open(FileName=full, ReadOnly=bool(read_only))
    return {
        "handle": _register_doc(doc),
        "name": doc.Name,
        "path": full,
        "paragraphs": int(doc.Paragraphs.Count),
        "already_open": False,
    }


@com_tool
def word_new(visible: bool = True, template: str | None = None) -> dict[str, Any]:
    """Create a new document, optionally from a .dotx template."""
    app = _wordapp(visible=visible)
    doc = (
        app.Documents.Add(Template=os.path.abspath(template))
        if template
        else app.Documents.Add()
    )
    return {"handle": _register_doc(doc), "name": doc.Name}


@com_tool
def word_save(
    handle: str | None = None, path: str | None = None, file_format: str | None = None
) -> dict[str, Any]:
    """Save a document; with `path`, save-as (format inferred from extension)."""
    doc = _doc(handle)
    if not path:
        doc.Save()
        return {"saved": True, "path": getattr(doc, "FullName", doc.Name)}
    full = os.path.abspath(os.path.expanduser(path))
    os.makedirs(os.path.dirname(full) or ".", exist_ok=True)
    key = (file_format or os.path.splitext(full)[1].lstrip(".")).lower()
    fmt = WD_FORMAT.get(key)
    if fmt is None:
        raise ComToolError(
            "Unsupported format %r" % key, hint="One of " + ", ".join(WD_FORMAT)
        )
    doc.SaveAs2(FileName=full, FileFormat=fmt)
    return {"saved": True, "path": full, "format": key}


@com_tool
def word_close(handle: str | None = None, save: bool = False) -> dict[str, Any]:
    """Close a document (save=False discards changes)."""
    doc = _doc(handle)
    name = doc.Name
    doc.Close(SaveChanges=-1 if save else 0)
    hid = handle if (handle or "").startswith("wddoc:") else find_handle_by(
        "wddoc", "name", name
    )
    if hid:
        drop_handle(hid)
    return {"closed": name, "saved": bool(save)}


@com_tool
def word_export_pdf(output_path: str, handle: str | None = None) -> dict[str, Any]:
    """Export the document to PDF."""
    doc = _doc(handle)
    full = os.path.abspath(os.path.expanduser(output_path))
    os.makedirs(os.path.dirname(full) or ".", exist_ok=True)
    doc.ExportAsFixedFormat(OutputFileName=full, ExportFormat=WD_EXPORT_PDF)
    return {"pdf": full, "exists": os.path.exists(full)}


@com_tool
def word_document_info(handle: str | None = None) -> dict[str, Any]:
    """Statistics and metadata: pages, words, headings, tracked-changes state."""
    doc = _doc(handle)
    stats = {}
    for label, code in (("pages", 2), ("words", 0), ("characters", 3), ("paragraphs", 4)):
        try:
            stats[label] = int(doc.ComputeStatistics(code))
        except Exception:  # noqa: BLE001
            stats[label] = None
    props = {}
    for name in ("Title", "Author", "Last author", "Creation date", "Company"):
        try:
            props[name] = jsonable(doc.BuiltInDocumentProperties(name).Value)
        except Exception:  # noqa: BLE001
            continue
    return {
        "name": doc.Name,
        "path": getattr(doc, "FullName", doc.Name),
        "statistics": stats,
        "properties": props,
        "track_revisions": bool(doc.TrackRevisions),
        "revisions": int(doc.Revisions.Count),
        "comments": int(doc.Comments.Count),
        "tables": int(doc.Tables.Count),
        "inline_shapes": int(doc.InlineShapes.Count),
    }


# --------------------------------------------------------------------------
# text
# --------------------------------------------------------------------------


@com_tool
def word_read_text(
    handle: str | None = None,
    start_paragraph: int | None = None,
    end_paragraph: int | None = None,
    max_chars: int = 40000,
    with_paragraph_numbers: bool = False,
) -> dict[str, Any]:
    """Read document text, optionally a paragraph slice (1-based, inclusive)."""
    doc = _doc(handle)
    if with_paragraph_numbers:
        total = int(doc.Paragraphs.Count)
        s = max(1, int(start_paragraph or 1))
        e = min(total, int(end_paragraph or total))
        out = []
        used = 0
        for i in range(s, e + 1):
            text = str(doc.Paragraphs(i).Range.Text).rstrip("\r\a")
            used += len(text)
            if used > max_chars:
                return {
                    "name": doc.Name,
                    "paragraphs": out,
                    "truncated": True,
                    "total_paragraphs": total,
                }
            out.append({"index": i, "style": str(doc.Paragraphs(i).Style.NameLocal),
                        "text": text})
        return {"name": doc.Name, "paragraphs": out, "truncated": False,
                "total_paragraphs": total}
    rng = _doc_range(doc, start_paragraph, end_paragraph)
    text = str(rng.Text)
    return {
        "name": doc.Name,
        "chars": len(text),
        "truncated": len(text) > max_chars,
        "text": text[:max_chars],
        "total_paragraphs": int(doc.Paragraphs.Count),
    }


@com_tool
def word_get_outline(handle: str | None = None, max_level: int = 9) -> dict[str, Any]:
    """List the heading structure (paragraph index, level, text)."""
    doc = _doc(handle)
    rows = []
    for i, par in enumerate(doc.Paragraphs, start=1):
        try:
            level = int(par.OutlineLevel)
        except Exception:  # noqa: BLE001
            continue
        if level <= max_level and level < 10:  # 10 = wdOutlineLevelBodyText
            rows.append(
                {
                    "index": i,
                    "level": level,
                    "style": str(par.Style.NameLocal),
                    "text": str(par.Range.Text).rstrip("\r\a"),
                }
            )
    return {"name": doc.Name, "count": len(rows), "headings": rows}


@com_tool
def word_append_text(
    text: str,
    handle: str | None = None,
    style: str | None = None,
    new_paragraph: bool = True,
) -> dict[str, Any]:
    """Append text at the end of the document, optionally with a style name."""
    doc = _doc(handle)
    rng = doc.Content
    rng.Collapse(WD_COLLAPSE_END)
    rng.InsertAfter(("\r" if new_paragraph else "") + text)
    if style:
        try:
            rng.Style = style
        except Exception as exc:  # noqa: BLE001
            raise ComToolError("Unknown style %r" % style) from exc
    return {"appended_chars": len(text), "paragraphs": int(doc.Paragraphs.Count)}


@com_tool
def word_insert_paragraph(
    text: str,
    after_paragraph: int,
    handle: str | None = None,
    style: str | None = None,
) -> dict[str, Any]:
    """Insert a new paragraph after the given 1-based paragraph index (0 = top)."""
    doc = _doc(handle)
    total = int(doc.Paragraphs.Count)
    if after_paragraph <= 0:
        rng = doc.Paragraphs(1).Range
        rng.InsertParagraphBefore()
        target = doc.Paragraphs(1).Range
    else:
        idx = min(int(after_paragraph), total)
        rng = doc.Paragraphs(idx).Range
        rng.InsertParagraphAfter()
        target = doc.Paragraphs(idx + 1).Range
    target.Text = text
    if style:
        target.Style = style
    return {"inserted_after": after_paragraph, "paragraphs": int(doc.Paragraphs.Count)}


@com_tool
def word_replace_text(
    find_text: str,
    replace_text: str,
    handle: str | None = None,
    match_case: bool = False,
    whole_word: bool = False,
    use_wildcards: bool = False,
    replace_all: bool = True,
) -> dict[str, Any]:
    """Find & replace across the document body."""
    doc = _doc(handle)
    find = doc.Content.Find
    find.ClearFormatting()
    find.Replacement.ClearFormatting()
    found = find.Execute(
        FindText=find_text,
        MatchCase=bool(match_case),
        MatchWholeWord=bool(whole_word),
        MatchWildcards=bool(use_wildcards),
        Forward=True,
        Wrap=WD_FIND_STOP,
        ReplaceWith=replace_text,
        Replace=WD_REPLACE_ALL if replace_all else 1,
    )
    return {"found": bool(found), "find": find_text, "replace": replace_text}


@com_tool
def word_find(
    query: str,
    handle: str | None = None,
    match_case: bool = False,
    use_wildcards: bool = False,
    max_results: int = 30,
    context_chars: int = 120,
) -> dict[str, Any]:
    """Find occurrences and return their paragraph index plus surrounding text."""
    doc = _doc(handle)
    rng = doc.Content
    find = rng.Find
    find.ClearFormatting()
    hits = []
    while len(hits) < max_results:
        ok = find.Execute(
            FindText=query,
            MatchCase=bool(match_case),
            MatchWildcards=bool(use_wildcards),
            Forward=True,
            Wrap=WD_FIND_STOP,
        )
        if not ok:
            break
        start = int(rng.Start)
        ctx = doc.Range(
            max(0, start - context_chars), min(int(doc.Content.End), int(rng.End) + context_chars)
        )
        hits.append(
            {
                "start": start,
                "paragraph": int(
                    doc.Range(0, start).Paragraphs.Count
                ),
                "match": str(rng.Text),
                "context": str(ctx.Text).replace("\r", " ").strip(),
            }
        )
        rng.Collapse(WD_COLLAPSE_END)
        find = rng.Find
    return {"query": query, "count": len(hits), "hits": hits}


@com_tool
def word_format_paragraphs(
    start_paragraph: int,
    end_paragraph: int | None = None,
    handle: str | None = None,
    style: str | None = None,
    bold: bool | None = None,
    italic: bool | None = None,
    font_size: float | None = None,
    font_name: str | None = None,
    font_color: str | None = None,
    alignment: str | None = None,
    space_after: float | None = None,
    line_spacing: float | None = None,
) -> dict[str, Any]:
    """Format a paragraph range (1-based, inclusive). alignment: left|center|right|justify."""
    doc = _doc(handle)
    rng = _doc_range(doc, start_paragraph, end_paragraph or start_paragraph)
    applied = []
    if style:
        rng.Style = style
        applied.append("style")
    if bold is not None:
        rng.Font.Bold = bool(bold)
        applied.append("bold")
    if italic is not None:
        rng.Font.Italic = bool(italic)
        applied.append("italic")
    if font_size is not None:
        rng.Font.Size = float(font_size)
        applied.append("font_size")
    if font_name:
        rng.Font.Name = font_name
        applied.append("font_name")
    if font_color:
        h = font_color.lstrip("#")
        r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
        rng.Font.Color = b * 65536 + g * 256 + r
        applied.append("font_color")
    if alignment:
        rng.ParagraphFormat.Alignment = WD_ALIGN.get(alignment.lower(), 0)
        applied.append("alignment")
    if space_after is not None:
        rng.ParagraphFormat.SpaceAfter = float(space_after)
        applied.append("space_after")
    if line_spacing is not None:
        rng.ParagraphFormat.LineSpacingRule = 0
        rng.ParagraphFormat.LineSpacing = float(line_spacing)
        applied.append("line_spacing")
    return {"paragraphs": [start_paragraph, end_paragraph or start_paragraph],
            "applied": applied}


@com_tool
def word_list_styles(handle: str | None = None, in_use_only: bool = True) -> dict:
    """List style names available in the document."""
    doc = _doc(handle)
    names = []
    for st in doc.Styles:
        try:
            if in_use_only and not bool(st.InUse):
                continue
            names.append(str(st.NameLocal))
        except Exception:  # noqa: BLE001
            continue
    return {"count": len(names), "styles": names}


# --------------------------------------------------------------------------
# tables, images, structure
# --------------------------------------------------------------------------


@com_tool
def word_list_tables(handle: str | None = None) -> dict[str, Any]:
    """List tables with their dimensions and first row."""
    doc = _doc(handle)
    rows = []
    for i, tbl in enumerate(doc.Tables, start=1):
        try:
            header = [
                str(tbl.Cell(1, c + 1).Range.Text).rstrip("\r\a\x07")
                for c in range(int(tbl.Columns.Count))
            ]
        except Exception:  # noqa: BLE001 - merged cells
            header = []
        rows.append(
            {
                "index": i,
                "rows": int(tbl.Rows.Count),
                "columns": int(tbl.Columns.Count),
                "first_row": header,
            }
        )
    return {"count": len(rows), "tables": rows}


@com_tool
def word_read_table(
    table_index: int = 1, handle: str | None = None, max_rows: int = 200
) -> dict[str, Any]:
    """Read a table into a 2-D list (1-based table index)."""
    doc = _doc(handle)
    if table_index < 1 or table_index > int(doc.Tables.Count):
        raise ComToolError("Document has %d tables" % int(doc.Tables.Count))
    tbl = doc.Tables(table_index)
    n_rows, n_cols = int(tbl.Rows.Count), int(tbl.Columns.Count)
    data = []
    for r in range(1, min(n_rows, max_rows) + 1):
        row = []
        for c in range(1, n_cols + 1):
            try:
                row.append(str(tbl.Cell(r, c).Range.Text).rstrip("\r\a\x07"))
            except Exception:  # noqa: BLE001 - merged/missing cell
                row.append(None)
        data.append(row)
    return {
        "table": table_index,
        "rows": n_rows,
        "columns": n_cols,
        "truncated": n_rows > max_rows,
        "values": data,
    }


@com_tool
def word_insert_table(
    values: list,
    handle: str | None = None,
    after_paragraph: int | None = None,
    header_bold: bool = True,
    style: str | None = "Table Grid",
) -> dict[str, Any]:
    """Insert a table from a 2-D list at the end (or after a paragraph index)."""
    doc = _doc(handle)
    if not values or not isinstance(values, list):
        raise ComToolError("values must be a non-empty list of rows")
    grid = [r if isinstance(r, list) else [r] for r in values]
    n_rows, n_cols = len(grid), max(len(r) for r in grid)
    if after_paragraph:
        anchor = doc.Paragraphs(min(int(after_paragraph),
                                    int(doc.Paragraphs.Count))).Range
        anchor.InsertParagraphAfter()
        rng = doc.Paragraphs(min(int(after_paragraph) + 1,
                                 int(doc.Paragraphs.Count))).Range
    else:
        rng = doc.Content
        rng.Collapse(WD_COLLAPSE_END)
        rng.InsertParagraphAfter()
        rng = doc.Paragraphs(int(doc.Paragraphs.Count)).Range
    tbl = doc.Tables.Add(rng, n_rows, n_cols)
    if style:
        try:
            tbl.Style = style
        except Exception:  # noqa: BLE001 - style missing in this template
            pass
    for r, row in enumerate(grid, start=1):
        for c, val in enumerate(row, start=1):
            tbl.Cell(r, c).Range.Text = "" if val is None else str(val)
    if header_bold and n_rows > 1:
        tbl.Rows(1).Range.Font.Bold = True
    return {"table_index": int(doc.Tables.Count), "rows": n_rows, "columns": n_cols}


@com_tool
def word_insert_image(
    image_path: str,
    handle: str | None = None,
    after_paragraph: int | None = None,
    width_points: float | None = None,
) -> dict[str, Any]:
    """Insert a picture at the end, or after a paragraph index."""
    doc = _doc(handle)
    full = os.path.abspath(os.path.expanduser(image_path))
    if not os.path.exists(full):
        raise ComToolError("Image not found: " + full)
    if after_paragraph:
        rng = doc.Paragraphs(min(int(after_paragraph),
                                 int(doc.Paragraphs.Count))).Range
        rng.Collapse(WD_COLLAPSE_END)
    else:
        rng = doc.Content
        rng.Collapse(WD_COLLAPSE_END)
    shape = doc.InlineShapes.AddPicture(FileName=full, LinkToFile=False,
                                        SaveWithDocument=True, Range=rng)
    if width_points:
        ratio = float(width_points) / float(shape.Width)
        shape.Width = float(width_points)
        shape.Height = float(shape.Height) * ratio
    return {"image": full, "inline_shapes": int(doc.InlineShapes.Count)}


@com_tool
def word_insert_break(
    handle: str | None = None, kind: str = "page", after_paragraph: int | None = None
) -> dict[str, Any]:
    """Insert a break. kind: page | section_next | column | line."""
    doc = _doc(handle)
    codes = {"page": 7, "section_next": 2, "column": 8, "line": 6}
    if kind not in codes:
        raise ComToolError("kind must be one of " + ", ".join(codes))
    if after_paragraph:
        rng = doc.Paragraphs(min(int(after_paragraph),
                                 int(doc.Paragraphs.Count))).Range
        rng.Collapse(WD_COLLAPSE_END)
    else:
        rng = doc.Content
        rng.Collapse(WD_COLLAPSE_END)
    rng.InsertBreak(codes[kind])
    return {"inserted": kind}


@com_tool
def word_insert_toc(
    handle: str | None = None, at_paragraph: int = 1, max_level: int = 3
) -> dict[str, Any]:
    """Insert a table of contents built from heading styles."""
    doc = _doc(handle)
    rng = doc.Paragraphs(min(int(at_paragraph), int(doc.Paragraphs.Count))).Range
    toc = doc.TablesOfContents.Add(
        Range=rng, UseHeadingStyles=True, UpperHeadingLevel=1,
        LowerHeadingLevel=int(max_level)
    )
    toc.Update()
    return {"toc_count": int(doc.TablesOfContents.Count)}


# --------------------------------------------------------------------------
# review: comments and revisions
# --------------------------------------------------------------------------


@com_tool
def word_list_comments(handle: str | None = None, max_results: int = 100) -> dict:
    """List comments with author, anchor text and reply text."""
    doc = _doc(handle)
    rows = []
    for i, cm in enumerate(doc.Comments, start=1):
        if i > max_results:
            break
        try:
            rows.append(
                {
                    "index": i,
                    "author": str(cm.Author),
                    "date": jsonable(cm.Date),
                    "scope_text": str(cm.Scope.Text).rstrip("\r\a"),
                    "text": str(cm.Range.Text),
                    "done": bool(getattr(cm, "Done", False)),
                }
            )
        except Exception:  # noqa: BLE001
            continue
    return {"count": int(doc.Comments.Count), "comments": rows}


@com_tool
def word_add_comment(
    text: str,
    handle: str | None = None,
    paragraph: int | None = None,
    anchor_text: str | None = None,
) -> dict[str, Any]:
    """Add a comment on a paragraph, or on the first match of anchor_text."""
    doc = _doc(handle)
    if anchor_text:
        rng = doc.Content
        find = rng.Find
        find.ClearFormatting()
        if not find.Execute(FindText=anchor_text, Forward=True, Wrap=WD_FIND_STOP):
            raise ComToolError("anchor_text not found: %r" % anchor_text)
        target = rng
    else:
        target = doc.Paragraphs(int(paragraph or 1)).Range
    doc.Comments.Add(Range=target, Text=text)
    return {"comments": int(doc.Comments.Count)}


@com_tool
def word_revisions(
    action: str = "list", handle: str | None = None, max_results: int = 100
) -> dict[str, Any]:
    """Tracked changes. action: list | enable | disable | accept_all | reject_all."""
    doc = _doc(handle)
    act = action.lower()
    if act == "enable":
        doc.TrackRevisions = True
        return {"track_revisions": True}
    if act == "disable":
        doc.TrackRevisions = False
        return {"track_revisions": False}
    if act == "accept_all":
        n = int(doc.Revisions.Count)
        doc.Revisions.AcceptAll()
        return {"accepted": n}
    if act == "reject_all":
        n = int(doc.Revisions.Count)
        doc.Revisions.RejectAll()
        return {"rejected": n}
    kinds = {0: "no_revision", 1: "insert", 2: "delete", 3: "property",
             4: "paragraph_number", 5: "display_field", 6: "reconcile",
             7: "conflict", 8: "style", 9: "replace", 10: "paragraph_property",
             11: "table_property", 12: "section_property", 13: "style_definition",
             14: "moved_from", 15: "moved_to", 16: "cell_insertion",
             17: "cell_deletion", 18: "cell_merge"}
    rows = []
    for i, rev in enumerate(doc.Revisions, start=1):
        if i > max_results:
            break
        try:
            rows.append(
                {
                    "index": i,
                    "type": kinds.get(int(rev.Type), int(rev.Type)),
                    "author": str(rev.Author),
                    "date": jsonable(rev.Date),
                    "text": str(rev.Range.Text)[:300],
                }
            )
        except Exception:  # noqa: BLE001
            continue
    return {
        "track_revisions": bool(doc.TrackRevisions),
        "count": int(doc.Revisions.Count),
        "revisions": rows,
    }


@com_tool
def word_run_macro(
    macro: str, args: list | None = None, handle: str | None = None
) -> dict[str, Any]:
    """Run a VBA macro in Word."""
    app = _wordapp()
    if handle:
        _doc(handle).Activate()
    result = app.Run(macro, *(args or []))
    return {"macro": macro, "result": jsonable(result)}


@com_tool
def word_list_open_handles() -> dict[str, Any]:
    """Diagnostics: document handles this server has handed out."""
    return {"handles": list_handles("wddoc")}


TOOLS = [
    word_list_documents,
    word_open,
    word_new,
    word_save,
    word_close,
    word_export_pdf,
    word_document_info,
    word_read_text,
    word_get_outline,
    word_append_text,
    word_insert_paragraph,
    word_replace_text,
    word_find,
    word_format_paragraphs,
    word_list_styles,
    word_list_tables,
    word_read_table,
    word_insert_table,
    word_insert_image,
    word_insert_break,
    word_insert_toc,
    word_list_comments,
    word_add_comment,
    word_revisions,
    word_run_macro,
    word_list_open_handles,
]
