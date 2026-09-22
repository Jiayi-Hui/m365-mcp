"""PowerPoint tools (COM).

Presentations, slides, shapes, text, notes, tables, images, charts, export.
Note: PowerPoint refuses Visible=False - it always automates a visible window.
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

# --- PowerPoint enums ---
PP_SAVE = {"pptx": 24, "ppt": 1, "pdf": 32, "png": 18, "jpg": 17, "gif": 16,
           "pptm": 25, "potx": 26, "xml": 11, "mp4": 39}
PP_LAYOUT = {
    "title": 1,
    "text": 2,
    "two_column": 3,
    "table": 4,
    "title_only": 11,
    "blank": 12,
    "object": 16,
    "chart": 8,
    "picture": 36,
}
MSO_TRUE = -1
MSO_FALSE = 0


def _pptapp() -> Any:
    return get_app("powerpoint")


def _register_pres(pres: Any) -> str:
    existing = find_handle_by("pppres", "name", pres.Name)
    if existing:
        return existing
    try:
        path = pres.FullName
    except Exception:  # noqa: BLE001
        path = pres.Name
    return put_handle("pppres", pres, name=pres.Name, path=path)


def _pres(handle: str | None) -> Any:
    app = _pptapp()
    if not handle:
        if int(app.Presentations.Count) == 0:
            raise ComToolError(
                "No presentation open", hint="Use ppt_open / ppt_new, or pass a handle."
            )
        return app.ActivePresentation
    if handle.startswith("pppres:"):
        return get_handle(handle)
    target = handle.strip().lower()
    for pres in app.Presentations:
        if pres.Name.lower() == target:
            return pres
        try:
            if os.path.abspath(pres.FullName).lower() == os.path.abspath(handle).lower():
                return pres
        except Exception:  # noqa: BLE001
            continue
    raise ComToolError("No open presentation matches %r" % handle)


def _slide(pres: Any, index: int) -> Any:
    total = int(pres.Slides.Count)
    if index < 1 or index > total:
        raise ComToolError("Slide %d out of range (1..%d)" % (index, total))
    return pres.Slides(index)


def _shape(slide: Any, shape: str | int) -> Any:
    if isinstance(shape, int) or (isinstance(shape, str) and shape.isdigit()):
        idx = int(shape)
        if idx < 1 or idx > int(slide.Shapes.Count):
            raise ComToolError("Shape index %d out of range" % idx)
        return slide.Shapes(idx)
    try:
        return slide.Shapes(shape)
    except Exception as exc:  # noqa: BLE001
        names = [s.Name for s in slide.Shapes]
        raise ComToolError(
            "No shape %r on this slide" % shape, hint="Shapes: " + ", ".join(names)
        ) from exc


def _shape_text(shape: Any) -> str | None:
    try:
        if shape.HasTextFrame and shape.TextFrame.HasText:
            return str(shape.TextFrame.TextRange.Text)
    except Exception:  # noqa: BLE001
        pass
    return None


# --------------------------------------------------------------------------
# presentations
# --------------------------------------------------------------------------


@com_tool
def ppt_list_presentations() -> dict[str, Any]:
    """List open presentations."""
    app = _pptapp()
    rows = []
    for pres in app.Presentations:
        rows.append(
            {
                "handle": _register_pres(pres),
                "name": pres.Name,
                "path": getattr(pres, "FullName", pres.Name),
                "slides": int(pres.Slides.Count),
                "saved": bool(pres.Saved),
            }
        )
    return {"count": len(rows), "presentations": rows}


@com_tool
def ppt_open(path: str, read_only: bool = False) -> dict[str, Any]:
    """Open a .pptx/.ppt/.potx presentation."""
    full = os.path.abspath(os.path.expanduser(path))
    if not os.path.exists(full):
        raise ComToolError("File not found: " + full)
    app = _pptapp()
    try:
        app.Visible = MSO_TRUE
    except Exception:  # noqa: BLE001
        pass
    for pres in app.Presentations:
        try:
            if os.path.abspath(pres.FullName).lower() == full.lower():
                return {"handle": _register_pres(pres), "name": pres.Name,
                        "already_open": True}
        except Exception:  # noqa: BLE001
            continue
    pres = app.Presentations.Open(
        FileName=full, ReadOnly=MSO_TRUE if read_only else MSO_FALSE,
        WithWindow=MSO_TRUE
    )
    return {
        "handle": _register_pres(pres),
        "name": pres.Name,
        "path": full,
        "slides": int(pres.Slides.Count),
        "already_open": False,
    }


@com_tool
def ppt_new(template: str | None = None) -> dict[str, Any]:
    """Create a new presentation, optionally from a .potx/.pptx template."""
    app = _pptapp()
    try:
        app.Visible = MSO_TRUE
    except Exception:  # noqa: BLE001
        pass
    pres = app.Presentations.Add(WithWindow=MSO_TRUE)
    if template:
        tpl = os.path.abspath(os.path.expanduser(template))
        if not os.path.exists(tpl):
            raise ComToolError("Template not found: " + tpl)
        pres.ApplyTemplate(tpl)
    return {"handle": _register_pres(pres), "name": pres.Name}


@com_tool
def ppt_save(
    handle: str | None = None, path: str | None = None, file_format: str | None = None
) -> dict[str, Any]:
    """Save a presentation; with `path`, save-as (format inferred from extension)."""
    pres = _pres(handle)
    if not path:
        pres.Save()
        return {"saved": True, "path": getattr(pres, "FullName", pres.Name)}
    full = os.path.abspath(os.path.expanduser(path))
    os.makedirs(os.path.dirname(full) or ".", exist_ok=True)
    key = (file_format or os.path.splitext(full)[1].lstrip(".")).lower()
    fmt = PP_SAVE.get(key)
    if fmt is None:
        raise ComToolError(
            "Unsupported format %r" % key, hint="One of " + ", ".join(PP_SAVE)
        )
    pres.SaveAs(FileName=full, FileFormat=fmt)
    return {"saved": True, "path": full, "format": key}


@com_tool
def ppt_close(handle: str | None = None, save: bool = False) -> dict[str, Any]:
    """Close a presentation (save=False discards changes)."""
    pres = _pres(handle)
    name = pres.Name
    if save:
        pres.Save()
    else:
        pres.Saved = MSO_TRUE  # suppress the save prompt
    pres.Close()
    hid = handle if (handle or "").startswith("pppres:") else find_handle_by(
        "pppres", "name", name
    )
    if hid:
        drop_handle(hid)
    return {"closed": name, "saved": bool(save)}


@com_tool
def ppt_export_pdf(output_path: str, handle: str | None = None) -> dict[str, Any]:
    """Export the presentation to PDF."""
    pres = _pres(handle)
    full = os.path.abspath(os.path.expanduser(output_path))
    os.makedirs(os.path.dirname(full) or ".", exist_ok=True)
    pres.SaveAs(FileName=full, FileFormat=PP_SAVE["pdf"])
    return {"pdf": full, "exists": os.path.exists(full)}


@com_tool
def ppt_export_images(
    output_dir: str,
    handle: str | None = None,
    image_format: str = "PNG",
    width: int = 1920,
    height: int = 1080,
    slide: int | None = None,
) -> dict[str, Any]:
    """Export slides as images (whole deck, or one slide)."""
    pres = _pres(handle)
    out = os.path.abspath(os.path.expanduser(output_dir))
    os.makedirs(out, exist_ok=True)
    if slide:
        target = os.path.join(out, "slide%d.%s" % (slide, image_format.lower()))
        _slide(pres, slide).Export(target, image_format, width, height)
        return {"files": [target], "count": 1}
    pres.Export(out, image_format, width, height)
    files = sorted(
        os.path.join(out, f)
        for f in os.listdir(out)
        if f.lower().endswith("." + image_format.lower())
    )
    return {"directory": out, "count": len(files), "files": files[:100]}


# --------------------------------------------------------------------------
# slides
# --------------------------------------------------------------------------


@com_tool
def ppt_list_slides(
    handle: str | None = None, include_text: bool = True, max_chars: int = 400
) -> dict[str, Any]:
    """List slides with title, layout and (optionally) the text of every shape."""
    pres = _pres(handle)
    rows = []
    for i in range(1, int(pres.Slides.Count) + 1):
        sld = pres.Slides(i)
        title = None
        try:
            if sld.Shapes.HasTitle:
                title = str(sld.Shapes.Title.TextFrame.TextRange.Text)
        except Exception:  # noqa: BLE001
            pass
        row: dict[str, Any] = {
            "index": i,
            "title": title,
            "layout": int(sld.Layout),
            "shapes": int(sld.Shapes.Count),
        }
        if include_text:
            texts = []
            for shp in sld.Shapes:
                t = _shape_text(shp)
                if t:
                    texts.append(t.strip())
            joined = " | ".join(texts)
            row["text"] = joined[:max_chars]
            row["text_truncated"] = len(joined) > max_chars
        try:
            notes = sld.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text
            row["notes"] = str(notes)[:max_chars] or None
        except Exception:  # noqa: BLE001
            row["notes"] = None
        rows.append(row)
    return {"presentation": pres.Name, "count": len(rows), "slides": rows}


@com_tool
def ppt_add_slide(
    handle: str | None = None,
    layout: str = "title_only",
    index: int | None = None,
    title: str | None = None,
    body: str | list | None = None,
) -> dict[str, Any]:
    """Add a slide. layout: title|text|two_column|table|title_only|blank|object|chart|picture.

    `body` may be a string or a list of bullet lines.
    """
    pres = _pres(handle)
    code = PP_LAYOUT.get(layout.lower())
    if code is None:
        raise ComToolError(
            "Unknown layout %r" % layout, hint="One of " + ", ".join(sorted(PP_LAYOUT))
        )
    pos = int(index) if index else int(pres.Slides.Count) + 1
    sld = pres.Slides.Add(pos, code)
    if title is not None:
        try:
            sld.Shapes.Title.TextFrame.TextRange.Text = title
        except Exception as exc:  # noqa: BLE001
            raise ComToolError(
                "This layout has no title placeholder", hint="Use layout=title_only"
            ) from exc
    if body is not None:
        text = "\r".join(body) if isinstance(body, list) else str(body)
        placed = False
        for shp in sld.Shapes:
            try:
                if shp.HasTextFrame and shp.Name != sld.Shapes.Title.Name:
                    shp.TextFrame.TextRange.Text = text
                    placed = True
                    break
            except Exception:  # noqa: BLE001
                continue
        if not placed:
            box = sld.Shapes.AddTextbox(1, 60, 140, 600, 300)
            box.TextFrame.TextRange.Text = text
    return {"slide": pos, "slides": int(pres.Slides.Count)}


@com_tool
def ppt_delete_slide(index: int, handle: str | None = None) -> dict[str, Any]:
    """Delete a slide by 1-based index."""
    pres = _pres(handle)
    _slide(pres, index).Delete()
    return {"deleted": index, "slides": int(pres.Slides.Count)}


@com_tool
def ppt_duplicate_slide(index: int, handle: str | None = None) -> dict[str, Any]:
    """Duplicate a slide; the copy lands right after the original."""
    pres = _pres(handle)
    _slide(pres, index).Duplicate()
    return {"duplicated": index, "slides": int(pres.Slides.Count)}


@com_tool
def ppt_move_slide(index: int, to_index: int, handle: str | None = None) -> dict:
    """Move a slide to another position."""
    pres = _pres(handle)
    _slide(pres, index).MoveTo(int(to_index))
    return {"moved": index, "to": to_index}


@com_tool
def ppt_copy_slide_from(
    source_path: str,
    source_index: int,
    handle: str | None = None,
    at_index: int | None = None,
) -> dict[str, Any]:
    """Insert a slide from another presentation file (keeps its design)."""
    pres = _pres(handle)
    full = os.path.abspath(os.path.expanduser(source_path))
    if not os.path.exists(full):
        raise ComToolError("File not found: " + full)
    pos = int(at_index) if at_index else int(pres.Slides.Count) + 1
    pres.Slides.InsertFromFile(full, pos - 1, int(source_index), int(source_index))
    return {"inserted_at": pos, "slides": int(pres.Slides.Count)}


@com_tool
def ppt_set_notes(index: int, text: str, handle: str | None = None) -> dict[str, Any]:
    """Set the speaker notes of a slide."""
    pres = _pres(handle)
    sld = _slide(pres, index)
    sld.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text = text
    return {"slide": index, "notes_chars": len(text)}


# --------------------------------------------------------------------------
# shapes
# --------------------------------------------------------------------------


@com_tool
def ppt_list_shapes(slide: int, handle: str | None = None) -> dict[str, Any]:
    """List the shapes on a slide with position, size and text."""
    pres = _pres(handle)
    sld = _slide(pres, slide)
    rows = []
    for i, shp in enumerate(sld.Shapes, start=1):
        rows.append(
            {
                "index": i,
                "name": shp.Name,
                "type": int(shp.Type),
                "left": round(float(shp.Left), 1),
                "top": round(float(shp.Top), 1),
                "width": round(float(shp.Width), 1),
                "height": round(float(shp.Height), 1),
                "text": _shape_text(shp),
            }
        )
    return {"slide": slide, "count": len(rows), "shapes": rows}


@com_tool
def ppt_set_text(
    slide: int, shape: str, text: str, handle: str | None = None
) -> dict[str, Any]:
    """Replace the text of a shape (by name or 1-based index)."""
    pres = _pres(handle)
    shp = _shape(_slide(pres, slide), shape)
    if not shp.HasTextFrame:
        raise ComToolError("Shape %r holds no text frame" % shape)
    shp.TextFrame.TextRange.Text = text
    return {"slide": slide, "shape": shp.Name, "chars": len(text)}


@com_tool
def ppt_add_textbox(
    slide: int,
    text: str,
    handle: str | None = None,
    left: float = 60.0,
    top: float = 60.0,
    width: float = 600.0,
    height: float = 100.0,
    font_size: float | None = None,
    bold: bool | None = None,
    font_color: str | None = None,
) -> dict[str, Any]:
    """Add a textbox at an absolute position (points; a 16:9 slide is 960x540)."""
    pres = _pres(handle)
    sld = _slide(pres, slide)
    box = sld.Shapes.AddTextbox(1, left, top, width, height)
    tr = box.TextFrame.TextRange
    tr.Text = text
    if font_size:
        tr.Font.Size = float(font_size)
    if bold is not None:
        tr.Font.Bold = MSO_TRUE if bold else MSO_FALSE
    if font_color:
        h = font_color.lstrip("#")
        r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
        tr.Font.Color.RGB = b * 65536 + g * 256 + r
    return {"slide": slide, "shape": box.Name}


@com_tool
def ppt_add_image(
    slide: int,
    image_path: str,
    handle: str | None = None,
    left: float = 60.0,
    top: float = 60.0,
    width: float | None = None,
    height: float | None = None,
) -> dict[str, Any]:
    """Add a picture to a slide (omit width/height to keep the native size)."""
    pres = _pres(handle)
    sld = _slide(pres, slide)
    full = os.path.abspath(os.path.expanduser(image_path))
    if not os.path.exists(full):
        raise ComToolError("Image not found: " + full)
    shp = sld.Shapes.AddPicture(
        FileName=full,
        LinkToFile=MSO_FALSE,
        SaveWithDocument=MSO_TRUE,
        Left=left,
        Top=top,
        Width=-1 if width is None else width,
        Height=-1 if height is None else height,
    )
    return {"slide": slide, "shape": shp.Name,
            "width": round(float(shp.Width), 1), "height": round(float(shp.Height), 1)}


@com_tool
def ppt_add_table(
    slide: int,
    values: list,
    handle: str | None = None,
    left: float = 60.0,
    top: float = 120.0,
    width: float = 800.0,
    height: float = 300.0,
) -> dict[str, Any]:
    """Add a table from a 2-D list."""
    pres = _pres(handle)
    sld = _slide(pres, slide)
    grid = [r if isinstance(r, list) else [r] for r in values]
    if not grid:
        raise ComToolError("values must be a non-empty list of rows")
    n_rows, n_cols = len(grid), max(len(r) for r in grid)
    shp = sld.Shapes.AddTable(n_rows, n_cols, left, top, width, height)
    tbl = shp.Table
    for r, row in enumerate(grid, start=1):
        for c in range(1, n_cols + 1):
            val = row[c - 1] if c - 1 < len(row) else None
            tbl.Cell(r, c).Shape.TextFrame.TextRange.Text = "" if val is None else str(val)
    return {"slide": slide, "shape": shp.Name, "rows": n_rows, "columns": n_cols}


@com_tool
def ppt_add_chart(
    slide: int,
    categories: list,
    series: dict,
    handle: str | None = None,
    chart_type: str = "column",
    title: str | None = None,
    left: float = 60.0,
    top: float = 120.0,
    width: float = 640.0,
    height: float = 360.0,
) -> dict[str, Any]:
    """Add a native chart. `series` maps series name -> list of values.

    chart_type: column | bar | line | pie | scatter | area
    """
    pres = _pres(handle)
    sld = _slide(pres, slide)
    kinds = {"column": 51, "bar": 57, "line": 4, "pie": 5, "scatter": -4169, "area": 1}
    kind = kinds.get(chart_type.lower())
    if kind is None:
        raise ComToolError("chart_type must be one of " + ", ".join(kinds))
    shp = sld.Shapes.AddChart2(-1, kind, left, top, width, height)
    chart = shp.Chart
    wb = chart.ChartData
    wb.Activate()
    ws = wb.Workbook.Worksheets(1)
    ws.Cells.Clear()
    ws.Cells(1, 1).Value = ""
    for c, name in enumerate(series.keys(), start=2):
        ws.Cells(1, c).Value = name
    for r, cat in enumerate(categories, start=2):
        ws.Cells(r, 1).Value = cat
    for c, values in enumerate(series.values(), start=2):
        for r, val in enumerate(values, start=2):
            ws.Cells(r, c).Value = val
    n_rows, n_cols = len(categories) + 1, len(series) + 1
    block = ws.Range(ws.Cells(1, 1), ws.Cells(n_rows, n_cols))
    try:
        address = str(block.Address(True, True))
    except TypeError:  # late-bound: Address is a plain string property
        address = str(block.Address)
    chart.SetSourceData("'%s'!%s" % (ws.Name, address))
    try:
        wb.Workbook.Close()
    except Exception:  # noqa: BLE001
        pass
    if title:
        chart.HasTitle = True
        chart.ChartTitle.Text = title
    return {"slide": slide, "shape": shp.Name, "series": list(series.keys())}


@com_tool
def ppt_format_shape(
    slide: int,
    shape: str,
    handle: str | None = None,
    left: float | None = None,
    top: float | None = None,
    width: float | None = None,
    height: float | None = None,
    fill_color: str | None = None,
    line_color: str | None = None,
    font_size: float | None = None,
    font_color: str | None = None,
    bold: bool | None = None,
    rotation: float | None = None,
) -> dict[str, Any]:
    """Move, resize, recolour a shape."""
    pres = _pres(handle)
    shp = _shape(_slide(pres, slide), shape)

    def _rgb(hex_color: str) -> int:
        h = hex_color.lstrip("#")
        r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
        return b * 65536 + g * 256 + r

    applied = []
    for attr, val in (("Left", left), ("Top", top), ("Width", width),
                      ("Height", height), ("Rotation", rotation)):
        if val is not None:
            setattr(shp, attr, float(val))
            applied.append(attr.lower())
    if fill_color:
        shp.Fill.Visible = MSO_TRUE
        shp.Fill.ForeColor.RGB = _rgb(fill_color)
        applied.append("fill")
    if line_color:
        shp.Line.Visible = MSO_TRUE
        shp.Line.ForeColor.RGB = _rgb(line_color)
        applied.append("line")
    if shp.HasTextFrame:
        tr = shp.TextFrame.TextRange
        if font_size:
            tr.Font.Size = float(font_size)
            applied.append("font_size")
        if font_color:
            tr.Font.Color.RGB = _rgb(font_color)
            applied.append("font_color")
        if bold is not None:
            tr.Font.Bold = MSO_TRUE if bold else MSO_FALSE
            applied.append("bold")
    return {"slide": slide, "shape": shp.Name, "applied": applied}


@com_tool
def ppt_delete_shape(slide: int, shape: str, handle: str | None = None) -> dict:
    """Delete a shape from a slide."""
    pres = _pres(handle)
    shp = _shape(_slide(pres, slide), shape)
    name = shp.Name
    shp.Delete()
    return {"slide": slide, "deleted": name}


@com_tool
def ppt_replace_text(
    find_text: str,
    replace_text: str,
    handle: str | None = None,
    slide: int | None = None,
) -> dict[str, Any]:
    """Replace text across the deck (or one slide) - useful for templating."""
    pres = _pres(handle)
    slides = [_slide(pres, slide)] if slide else list(pres.Slides)
    count = 0
    for sld in slides:
        for shp in sld.Shapes:
            try:
                if not shp.HasTextFrame or not shp.TextFrame.HasText:
                    continue
                tr = shp.TextFrame.TextRange
                found = tr.Replace(find_text, replace_text)
                while found is not None:
                    count += 1
                    found = tr.Replace(find_text, replace_text)
            except Exception:  # noqa: BLE001
                continue
    return {"replacements": count, "find": find_text}


@com_tool
def ppt_run_macro(
    macro: str, args: list | None = None, handle: str | None = None
) -> dict[str, Any]:
    """Run a VBA macro in PowerPoint."""
    app = _pptapp()
    result = app.Run(macro, *(args or []))
    return {"macro": macro, "result": jsonable(result)}


@com_tool
def ppt_list_open_handles() -> dict[str, Any]:
    """Diagnostics: presentation handles this server has handed out."""
    return {"handles": list_handles("pppres")}


TOOLS = [
    ppt_list_presentations,
    ppt_open,
    ppt_new,
    ppt_save,
    ppt_close,
    ppt_export_pdf,
    ppt_export_images,
    ppt_list_slides,
    ppt_add_slide,
    ppt_delete_slide,
    ppt_duplicate_slide,
    ppt_move_slide,
    ppt_copy_slide_from,
    ppt_set_notes,
    ppt_list_shapes,
    ppt_set_text,
    ppt_add_textbox,
    ppt_add_image,
    ppt_add_table,
    ppt_add_chart,
    ppt_format_shape,
    ppt_delete_shape,
    ppt_replace_text,
    ppt_run_macro,
    ppt_list_open_handles,
]
