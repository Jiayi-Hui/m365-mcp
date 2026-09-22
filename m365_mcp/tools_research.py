"""Research primitives: read notes of any format, and map a financial model.

These exist to support the "notes -> model -> memo" workflows. They are inputs
to judgement, not judgement: `excel_model_map` reports what the workbook IS
(which cells are assumptions, which are formulas, what each row and column
means), so an agent can locate the cell a note refers to. It never decides what
a number should be.

Design notes drawn from a real sell-side model (Huali 300979.SZ, 25 sheets,
9,506 defined names):

  * Colour is the convention that matters. In that model - and in the sell-side
    house style generally - green constants are reported history, blue
    constants are the analyst's assumptions, red/orange are manual overrides.
    Black constants are usually labels or long-settled history. We report the
    distribution and a conventional reading; we do not silently enforce it.
  * Defined names are mostly garbage: 5,510 of 9,506 were #REF!, and most of
    the rest were FactSet audit-link string arrays. Anything that does not
    resolve to a real range is dropped.
  * Sell-side models keep dated snapshot sheets (e.g. "241104") that are pure
    values with no formulas. They are previous published versions, not part of
    the live model, and must not be confused with it.
"""

from __future__ import annotations

import datetime as _dt
import json
import os
import re
from typing import Any

from .comcore import ComToolError, com_tool, jsonable
from .tools_excel import _addr, _sheet, _wb, excel_open

# --- colour conventions (COM stores colour as BGR) -------------------------

COLOUR_MEANING = {
    "#0000FF": "assumption - analyst input (blue)",
    "#000080": "assumption - analyst input (dark blue)",
    "#008000": "reported actual, hardcoded (green)",
    "#C00000": "manual override / flagged (dark red)",
    "#FF0000": "manual override / flagged (red)",
    "#E26B0A": "manual override / flagged (orange)",
    "#7030A0": "link to another workbook (purple)",
    "#000000": "plain constant - label or settled history (black)",
    "#FFFFFF": "hidden helper value (white on white)",
}

WRITABLE_CLASSES = {"assumption", "override"}


def _hex_from_bgr(colour: int) -> str:
    """COM font colour is BGR-packed; return #RRGGBB."""
    c = int(colour)
    return "#%02X%02X%02X" % (c & 0xFF, (c >> 8) & 0xFF, (c >> 16) & 0xFF)


def _classify(hex_colour: str) -> str:
    if hex_colour in ("#0000FF", "#000080", "#1F497D"):
        return "assumption"
    if hex_colour == "#008000":
        return "actual"
    if hex_colour in ("#C00000", "#FF0000", "#E26B0A", "#FFC000"):
        return "override"
    if hex_colour == "#7030A0":
        return "external_link"
    if hex_colour == "#FFFFFF":
        return "hidden_helper"
    if hex_colour == "#000000":
        return "plain"
    return "other"


# --- period / label heuristics --------------------------------------------

_PERIOD_RE = re.compile(
    r"^\s*(?:"
    r"(?:19|20)\d{2}\s*[EA]?"                  # 2026, 2026E
    r"|[1-4]Q\s*(?:FY)?\d{2,4}\s*[EA]?"        # 1Q26E, 1QFY26
    r"|(?:FY)?\d{2,4}\s*[1-4]Q\s*[EA]?"        # FY261Q
    r"|[1-2]H\s*(?:FY)?\d{2,4}\s*[EA]?"        # 1H26E
    r"|(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*[-\s]?\d{2,4}"
    r")\s*$",
    re.I,
)

_SNAPSHOT_RE = re.compile(r"^(?:20)?\d{6}$|^\d{2}[HQ][12]$|^\d{6}[a-z]?$", re.I)


def _looks_like_period(value: Any) -> bool:
    if value is None:
        return False
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return 1990 <= float(value) <= 2100 and float(value) == int(value)
    if isinstance(value, _dt.datetime):
        return True
    return bool(_PERIOD_RE.match(str(value)))


def _col_letter(index: int) -> str:
    """1 -> A, 27 -> AA. Cheaper than asking Excel for the address."""
    letters = ""
    while index > 0:
        index, rem = divmod(index - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


def _block_values(ws: Any, top: int, left: int, rows: int, cols: int) -> list[list]:
    """Read a rectangular block in ONE COM round trip, always as a 2-D list.

    Range.Value returns a scalar for a single cell and a flat tuple for a single
    row/column, so normalise before anything downstream indexes into it.
    """
    rows, cols = max(1, int(rows)), max(1, int(cols))
    block = ws.Range(ws.Cells(top, left), ws.Cells(top + rows - 1, left + cols - 1))
    raw = block.Value
    if rows == 1 and cols == 1:
        return [[raw]]
    if not isinstance(raw, tuple):
        return [[raw]]
    if raw and isinstance(raw[0], tuple):
        return [list(r) for r in raw]
    return [list(raw)] if rows == 1 else [[v] for v in raw]


def _period_text(value: Any) -> str | None:
    if value is None:
        return None
    if isinstance(value, _dt.datetime):
        return value.strftime("%Y-%m")
    if isinstance(value, float) and value == int(value):
        return str(int(value))
    return str(value).strip() or None


# --------------------------------------------------------------------------
# notes / documents in
# --------------------------------------------------------------------------


@com_tool
def pdf_read_text(
    path: str,
    pages: str | None = None,
    max_chars: int = 60000,
    with_page_numbers: bool = True,
) -> dict[str, Any]:
    """Extract text from a PDF (research note, filing, transcript).

    Args:
        pages: "1-5" or "3" or "2,7,9"; omit for the whole document
        with_page_numbers: prefix each page with [p12] so quotes stay locatable
    """
    full = os.path.abspath(os.path.expanduser(path))
    if not os.path.exists(full):
        raise ComToolError("File not found: " + full)
    try:
        import fitz  # PyMuPDF
    except ImportError as exc:  # pragma: no cover - dependency is in requirements
        raise ComToolError(
            "PyMuPDF is not installed", hint="pip install pymupdf"
        ) from exc

    doc = fitz.open(full)
    try:
        total = doc.page_count
        wanted: list[int] = []
        if pages:
            for part in str(pages).split(","):
                part = part.strip()
                if "-" in part:
                    a, b = part.split("-", 1)
                    wanted.extend(range(int(a) - 1, min(int(b), total)))
                elif part:
                    wanted.append(int(part) - 1)
        else:
            wanted = list(range(total))
        wanted = [i for i in wanted if 0 <= i < total]

        chunks = []
        used = 0
        truncated = False
        for i in wanted:
            text = doc.load_page(i).get_text()
            if with_page_numbers:
                text = "[p%d]\n%s" % (i + 1, text)
            if used + len(text) > max_chars:
                chunks.append(text[: max(0, max_chars - used)])
                truncated = True
                break
            chunks.append(text)
            used += len(text)
        body = "\n".join(chunks)
    finally:
        doc.close()
    return {
        "path": full,
        "total_pages": total,
        "pages_read": len(wanted),
        "chars": len(body),
        "truncated": truncated,
        "text": body,
    }


@com_tool
def notes_read(
    path: str, max_chars: int = 60000, numbered: bool = True
) -> dict[str, Any]:
    """Read notes of any supported format into numbered lines.

    Handles .md / .txt / .csv directly, .docx / .doc / .rtf through Word, and
    .pdf through PyMuPDF. Line numbers are the point: a proposal that changes a
    model must be able to cite the line it came from.
    """
    full = os.path.abspath(os.path.expanduser(path))
    if not os.path.exists(full):
        raise ComToolError("File not found: " + full)
    ext = os.path.splitext(full)[1].lower()

    if ext in (".md", ".markdown", ".txt", ".csv", ".json", ".yaml", ".yml"):
        with open(full, "r", encoding="utf-8", errors="replace") as fh:
            body = fh.read()
        source = "text"
    elif ext == ".pdf":
        result = pdf_read_text(path=full, max_chars=max_chars)
        if not result.get("ok", True):
            return result
        body = result["text"]
        source = "pdf"
    elif ext in (".docx", ".doc", ".rtf", ".odt", ".htm", ".html"):
        from .tools_word import word_open, word_close, word_read_text

        opened = word_open(path=full, read_only=True, visible=False)
        if not opened.get("ok", True):
            return opened
        handle = opened["handle"]
        try:
            read = word_read_text(handle=handle, max_chars=max_chars)
            body = read.get("text", "") if read.get("ok", True) else ""
        finally:
            word_close(handle=handle, save=False)
        body = body.replace("\r", "\n")
        source = "word"
    else:
        raise ComToolError(
            "Unsupported notes format %r" % ext,
            hint="Supported: .md .txt .csv .json .yaml .pdf .docx .doc .rtf .html",
        )

    truncated = len(body) > max_chars
    body = body[:max_chars]
    lines = body.splitlines()
    out: dict[str, Any] = {
        "path": full,
        "source": source,
        "chars": len(body),
        "lines": len(lines),
        "truncated": truncated,
    }
    if numbered:
        out["numbered_text"] = "\n".join(
            "%4d| %s" % (i, line) for i, line in enumerate(lines, start=1)
        )
    else:
        out["text"] = body
    return out


# --------------------------------------------------------------------------
# model map
# --------------------------------------------------------------------------


def _sheet_role(name: str, formulas: int, constants: int, visible: int,
                has_charts: bool) -> str:
    lowered = name.lower()
    if visible == 2 or lowered in ("mw-cache", "_cache"):
        return "cache"
    if has_charts and formulas < 200:
        return "charts"
    if _SNAPSHOT_RE.match(name.strip()) and formulas == 0:
        return "snapshot"
    if formulas == 0 and constants > 20:
        return "data"
    if formulas > constants:
        return "model"
    if formulas > 0:
        return "mixed"
    return "other"


@com_tool(timeout=900)
def excel_model_map(
    path: str | None = None,
    handle: str | None = None,
    sheets: str | None = None,
    max_inputs_per_sheet: int = 400,
    include_plain_constants: bool = False,
    include_data_sheets: bool = False,
    max_cells_scan: int = 60000,
    max_names_scan: int = 400,
) -> dict[str, Any]:
    """Map a financial model: sheet roles, axes, and every assumption cell.

    This is the structural survey the other model tools build on. For each
    sheet it reports the role (model / data / snapshot / charts / cache), the
    label column and period header row it inferred, and the constants grouped
    by font colour - which in sell-side models is what separates an analyst
    assumption (blue) from reported history (green) from a manual override
    (red/orange).

    Each listed cell carries its row label and period header, so "L20" arrives
    as "Net financial expenses x 2026E" - the coordinates a note can be matched
    against.

    Args:
        path: workbook to open read-only (or pass `handle` for an open one)
        sheets: comma-separated subset; omit to map everything
        max_inputs_per_sheet: cap on listed cells per sheet
        include_plain_constants: also list black constants (usually noise)
        include_data_sheets: also scan pure-data sheets (slow, rarely useful)
        max_cells_scan: skip the cell-level scan on sheets bigger than this
        max_names_scan: defined names to inspect (sell-side models carry
            thousands of dead ones; each costs a COM round trip)
    """
    if not path and not handle:
        raise ComToolError("Pass path= or handle=")
    if path and not handle:
        opened = excel_open(path=path, read_only=True, visible=False,
                            update_links=False)
        if not opened.get("ok", True):
            return opened
        handle = opened["handle"]

    wanted = None
    if sheets:
        wanted = {s.strip().lower() for s in sheets.split(",") if s.strip()}

    def work() -> dict[str, Any]:
        wb = _wb(handle)
        sheet_rows: list[dict[str, Any]] = []
        colour_totals: dict[str, int] = {}
        warnings: list[str] = []

        for ws in wb.Worksheets:
            name = str(ws.Name)
            if wanted and name.lower() not in wanted:
                continue
            used = ws.UsedRange
            n_rows, n_cols = int(used.Rows.Count), int(used.Columns.Count)
            try:
                visible = int(ws.Visible)
            except Exception:  # noqa: BLE001
                visible = -1
            try:
                has_charts = int(ws.ChartObjects().Count) > 0
            except Exception:  # noqa: BLE001
                has_charts = False

            n_formula = 0
            try:
                n_formula = int(used.SpecialCells(-4123).Count)  # xlCellTypeFormulas
            except Exception:  # noqa: BLE001 - none present
                pass
            n_const = 0
            try:
                n_const = int(used.SpecialCells(2).Count)  # xlCellTypeConstants
            except Exception:  # noqa: BLE001
                pass

            row: dict[str, Any] = {
                "name": name,
                "index": int(ws.Index),
                "visible": {0: "hidden", 2: "very_hidden"}.get(visible, "visible"),
                "used_range": _addr(used),
                "rows": n_rows,
                "columns": n_cols,
                "formulas": n_formula,
                "constants": n_const,
                "role": _sheet_role(name, n_formula, n_const, visible, has_charts),
            }

            if n_rows * n_cols > max_cells_scan:
                row["scan"] = "skipped: %d cells over max_cells_scan" % (
                    n_rows * n_cols)
                sheet_rows.append(row)
                continue
            if row["role"] in ("cache", "charts", "snapshot"):
                sheet_rows.append(row)
                continue
            if row["role"] == "data" and not include_data_sheets:
                row["scan"] = "skipped: pure-data sheet (no formulas, no assumptions)"
                sheet_rows.append(row)
                continue

            # --- axes: the period header row and the label column ---
            # Read the top band and the left band in ONE round trip each. Cell
            # by cell this was thousands of COM calls per sheet.
            top = int(used.Row)
            left = int(used.Column)
            band_rows = min(12, n_rows)
            band_cols = min(60, n_cols)
            band = _block_values(
                ws, top, left, band_rows, band_cols
            )  # band[r][c], 0-based within the band

            header_row = None
            periods: list[dict[str, Any]] = []
            best_hits = 0
            for ri, values in enumerate(band):
                hits = [
                    (left + ci, _period_text(v))
                    for ci, v in enumerate(values)
                    if _looks_like_period(v)
                ]
                if len(hits) > best_hits:
                    best_hits = len(hits)
                    header_row = top + ri
                    periods = [
                        {"col_index": c, "period": p} for c, p in hits
                    ]
            if best_hits < 2:
                header_row, periods = None, []

            label_col = None
            label_by_row: dict[int, str] = {}
            if header_row:
                left_band = _block_values(
                    ws, header_row + 1, left,
                    min(n_rows, top + n_rows - header_row - 1), min(6, n_cols)
                )
                best_text = 0
                for ci in range(len(left_band[0]) if left_band else 0):
                    filled = sum(
                        1 for r in left_band
                        if isinstance(r[ci], str) and r[ci].strip()
                    )
                    if filled > best_text:
                        best_text, label_col = filled, left + ci
                if label_col is not None:
                    ci = label_col - left
                    for ri, values in enumerate(left_band):
                        text = values[ci]
                        if not (isinstance(text, str) and text.strip()):
                            # fall back to any label further left on that row
                            for back in range(ci - 1, -1, -1):
                                alt = values[back]
                                if isinstance(alt, str) and alt.strip():
                                    text = alt
                                    break
                        if isinstance(text, str) and text.strip():
                            label_by_row[header_row + 1 + ri] = text.strip()

            row["header_row"] = header_row
            row["label_column"] = label_col
            row["periods"] = [p["period"] for p in periods][:60]

            # --- constants by colour ---
            period_by_col = {p["col_index"]: p["period"] for p in periods}
            cells: list[dict[str, Any]] = []
            colours: dict[str, int] = {}
            try:
                consts = used.SpecialCells(2)
            except Exception:  # noqa: BLE001
                consts = None
            if consts is not None:
                for area in consts.Areas:
                    a_count = int(area.Count)
                    if a_count > 8000:
                        warnings.append(
                            "%s: skipped a %d-cell constant block" % (name, a_count)
                        )
                        continue
                    a_top, a_left = int(area.Row), int(area.Column)
                    a_rows, a_cols = int(area.Rows.Count), int(area.Columns.Count)
                    values = _block_values(ws, a_top, a_left, a_rows, a_cols)

                    # One probe first: a block of uniform font colour answers for
                    # all of its cells at once. Excel returns None when mixed.
                    block_colour = None
                    try:
                        block_colour = area.Font.Color
                    except Exception:  # noqa: BLE001
                        block_colour = None

                    uniform_hex = (
                        _hex_from_bgr(block_colour) if block_colour is not None else None
                    )
                    if uniform_hex is not None and _classify(uniform_hex) == "plain" \
                            and not include_plain_constants:
                        numeric = sum(
                            1 for r in values for v in r
                            if v is not None and not isinstance(v, str)
                        )
                        colours[uniform_hex] = colours.get(uniform_hex, 0) + numeric
                        colour_totals[uniform_hex] = (
                            colour_totals.get(uniform_hex, 0) + numeric)
                        continue  # nothing listable in this block

                    for ri in range(a_rows):
                        for ci in range(a_cols):
                            value = values[ri][ci]
                            if value is None or isinstance(value, str):
                                continue
                            r, c = a_top + ri, a_left + ci
                            hexc = uniform_hex
                            if hexc is None:
                                hexc = _hex_from_bgr(ws.Cells(r, c).Font.Color)
                            klass = _classify(hexc)
                            colours[hexc] = colours.get(hexc, 0) + 1
                            colour_totals[hexc] = colour_totals.get(hexc, 0) + 1
                            if klass == "plain" and not include_plain_constants:
                                continue
                            if len(cells) >= max_inputs_per_sheet:
                                continue
                            cells.append(
                                {
                                    "cell": "%s%d" % (_col_letter(c), r),
                                    "value": jsonable(value),
                                    "colour": hexc,
                                    "class": klass,
                                    "writable": klass in WRITABLE_CLASSES,
                                    "label": label_by_row.get(r),
                                    "period": period_by_col.get(c),
                                }
                            )
            row["colour_census"] = dict(
                sorted(colours.items(), key=lambda kv: -kv[1]))
            row["listed_cells"] = len(cells)
            row["cells_omitted"] = max(
                0, sum(colours.values()) - len(cells)) if colours else 0
            row["cells"] = cells
            sheet_rows.append(row)

        # --- defined names, minus the archaeology ---
        names_total = int(wb.Names.Count)
        names: list[dict[str, Any]] = []
        junk = {"ref_error": 0, "hidden": 0, "not_a_range": 0, "vendor": 0}
        scanned = 0
        for nm in wb.Names:
            scanned += 1
            if scanned > max_names_scan:
                break
            try:
                refers = str(nm.RefersTo)
                nm_name = str(nm.Name)
                visible = bool(nm.Visible)
            except Exception:  # noqa: BLE001
                junk["ref_error"] += 1
                continue
            if "#REF" in refers:
                junk["ref_error"] += 1
            elif not visible:
                junk["hidden"] += 1
            elif nm_name.startswith(("__FDS", "_xl", "_FilterDatabase")) or \
                    "FDSAUDITLINK" in nm_name:
                junk["vendor"] += 1
            elif not re.search(r"![$]?[A-Z]{1,3}[$]?\d+", refers):
                junk["not_a_range"] += 1
            elif len(names) < 200:
                names.append({"name": nm_name, "refers_to": refers})

        links: list[str] = []
        try:
            sources = wb.LinkSources(1)  # xlExcelLinks
            links = [str(s) for s in (sources or [])][:20]
        except Exception:  # noqa: BLE001
            pass

        legend = {
            hexc: COLOUR_MEANING.get(hexc, "unmapped colour")
            for hexc in sorted(colour_totals, key=lambda k: -colour_totals[k])
        }
        return {
            "workbook": str(wb.Name),
            "path": str(getattr(wb, "FullName", wb.Name)),
            "handle": handle,
            "sheet_count": len(sheet_rows),
            "sheets": sheet_rows,
            "colour_legend": legend,
            "colour_totals": dict(
                sorted(colour_totals.items(), key=lambda kv: -kv[1])),
            "named_ranges": names,
            "named_ranges_total": names_total,
            "named_ranges_scanned": scanned,
            "named_ranges_dropped": junk,
            "external_links": links,
            "warnings": warnings,
            "reading": (
                "class=assumption cells are the analyst's inputs and the normal "
                "target for a note-driven change; class=actual are reported "
                "figures; class=override are manual overrides worth flagging. "
                "Colour convention is inferred, not guaranteed - confirm against "
                "a cell you recognise before relying on it."
            ),
        }

    return work()


@com_tool
def excel_propose_changes(
    changes: list,
    handle: str | None = None,
    changeset_path: str | None = None,
    note_source: str | None = None,
    allow_formula_cells: bool = False,
) -> dict[str, Any]:
    """Validate a set of proposed model edits and write them to a changeset file.

    The agent decides WHAT to change and why; this decides whether each edit is
    safe to make, and records it so a human can review it outside the chat.

    Each change is a dict:
        {"sheet": "Breakdown", "cell": "L14", "new_value": 19500,
         "reason": "1H26 Sportswear revenue 9,934mn, -12.4% YoY",
         "source": "2026H1 report p15", "confidence": "high"}

    Every proposal is checked against the live workbook and comes back with a
    verdict:
      ok            - a writable assumption/override cell
      blocked       - the target holds a FORMULA; writing would silently break
                      the model's structure. Change its driver instead, or pass
                      allow_formula_cells=True deliberately.
      warn          - writable but not a recognised assumption colour, or the
                      value is unchanged, or the cell is empty
    Nothing is written to the workbook here - use excel_apply_changeset.
    """
    if not isinstance(changes, list) or not changes:
        raise ComToolError("changes must be a non-empty list")
    wb = _wb(handle)
    reviewed: list[dict[str, Any]] = []
    counts = {"ok": 0, "warn": 0, "blocked": 0}

    for item in changes:
        if not isinstance(item, dict):
            raise ComToolError("each change must be an object, got %r" % type(item))
        cell_ref = item.get("cell")
        if not cell_ref:
            raise ComToolError("each change needs a 'cell'")
        entry: dict[str, Any] = {
            "sheet": item.get("sheet"),
            "cell": cell_ref,
            "new_value": item.get("new_value"),
            "reason": item.get("reason"),
            "source": item.get("source"),
            "confidence": item.get("confidence"),
        }
        try:
            ws = _sheet(wb, item.get("sheet"))
            rng = ws.Range(cell_ref)
        except Exception as exc:  # noqa: BLE001
            entry.update(verdict="blocked", issue="cannot resolve cell: %s" % exc)
            counts["blocked"] += 1
            reviewed.append(entry)
            continue

        has_formula = bool(rng.HasFormula)
        current = jsonable(rng.Value)
        hexc = _hex_from_bgr(rng.Font.Color)
        klass = _classify(hexc)
        entry.update(
            sheet=str(ws.Name),
            current_value=current,
            colour=hexc,
            cell_class=klass,
            has_formula=has_formula,
            current_formula=jsonable(rng.Formula) if has_formula else None,
        )
        try:
            entry["dependents"] = int(rng.Dependents.Count)
        except Exception:  # noqa: BLE001 - no dependents
            entry["dependents"] = 0

        issues = []
        if has_formula and not allow_formula_cells:
            entry["verdict"] = "blocked"
            issues.append(
                "target is a formula (%s); overwriting it replaces model "
                "structure with a hardcoded number" % entry["current_formula"]
            )
        else:
            entry["verdict"] = "ok"
            if klass not in WRITABLE_CLASSES:
                entry["verdict"] = "warn"
                issues.append(
                    "cell is classed '%s' (%s), not a recognised assumption "
                    "colour" % (klass, hexc)
                )
            if current is None:
                entry["verdict"] = "warn"
                issues.append("cell is currently empty")
            elif current == item.get("new_value"):
                entry["verdict"] = "warn"
                issues.append("new value equals the current value")
            if not item.get("reason") or not item.get("source"):
                entry["verdict"] = "warn"
                issues.append(
                    "no reason/source cited - a model edit should point at the "
                    "line of notes it came from"
                )
        entry["issues"] = issues
        counts[entry["verdict"]] += 1
        reviewed.append(entry)

    target = changeset_path
    if not target:
        base = str(getattr(wb, "FullName", "")) or str(wb.Name)
        root = os.path.splitext(base)[0] if os.path.dirname(base) else os.path.abspath(
            os.path.splitext(str(wb.Name))[0])
        target = root + ".changeset.json"
    payload = {
        "workbook": str(getattr(wb, "FullName", wb.Name)),
        "created": _dt.datetime.now().isoformat(timespec="seconds"),
        "note_source": note_source,
        "counts": counts,
        "changes": reviewed,
        "applied": False,
    }
    os.makedirs(os.path.dirname(target) or ".", exist_ok=True)
    with open(target, "w", encoding="utf-8") as fh:
        json.dump(payload, fh, ensure_ascii=False, indent=1)

    return {
        "changeset": target,
        "counts": counts,
        "changes": reviewed,
        "next_step": (
            "Show these to the user. Blocked rows must be re-pointed at the "
            "driver cell. Then call excel_apply_changeset with confirm=true."
        ),
    }


@com_tool(timeout=600)
def excel_apply_changeset(
    changeset_path: str,
    confirm: bool = False,
    handle: str | None = None,
    snapshot: bool = True,
    annotate: bool = True,
    include_warnings: bool = True,
) -> dict[str, Any]:
    """Apply a reviewed changeset to the workbook. Requires confirm=True.

    Takes a snapshot first, writes only rows the review passed, and leaves the
    reason and source in a cell comment so the edit stays auditable months later.
    Rows marked `blocked` are never written.
    """
    full = os.path.abspath(os.path.expanduser(changeset_path))
    if not os.path.exists(full):
        raise ComToolError("Changeset not found: " + full)
    with open(full, "r", encoding="utf-8") as fh:
        payload = json.load(fh)

    rows = payload.get("changes", [])
    writable = [
        r for r in rows
        if r.get("verdict") == "ok" or (include_warnings and r.get("verdict") == "warn")
    ]
    if not confirm:
        return {
            "ok": False,
            "error": "confirm=False - nothing was written",
            "changeset": full,
            "would_write": len(writable),
            "blocked": sum(1 for r in rows if r.get("verdict") == "blocked"),
            "preview": [
                {"cell": "%s!%s" % (r.get("sheet"), r.get("cell")),
                 "from": r.get("current_value"), "to": r.get("new_value"),
                 "verdict": r.get("verdict"), "reason": r.get("reason")}
                for r in writable[:20]
            ],
            "hint": "Show the preview to the user; call again with confirm=true "
                    "once they approve.",
        }

    wb = _wb(handle or payload.get("workbook"))
    snap = None
    if snapshot:
        snap = excel_snapshot(handle=handle or payload.get("workbook"),
                             label="pre-changeset")
        if not snap.get("ok", True):
            return snap

    written, failed = [], []
    for r in writable:
        try:
            ws = _sheet(wb, r.get("sheet"))
            rng = ws.Range(r["cell"])
            before = jsonable(rng.Value)
            rng.Value = r.get("new_value")
            if annotate:
                text = "m365-mcp %s\n%s\nsource: %s\nwas: %s" % (
                    _dt.datetime.now().strftime("%Y-%m-%d"),
                    r.get("reason") or "(no reason given)",
                    r.get("source") or "(no source given)",
                    before,
                )
                try:
                    if rng.Comment is not None:
                        rng.Comment.Delete()
                except Exception:  # noqa: BLE001 - no existing comment
                    pass
                try:
                    rng.AddComment(text[:900])
                except Exception:  # noqa: BLE001 - comments can be disabled
                    pass
            written.append({"cell": "%s!%s" % (ws.Name, r["cell"]),
                            "from": before, "to": r.get("new_value")})
        except Exception as exc:  # noqa: BLE001
            failed.append({"cell": "%s!%s" % (r.get("sheet"), r.get("cell")),
                           "error": str(exc)[:200]})

    payload["applied"] = True
    payload["applied_at"] = _dt.datetime.now().isoformat(timespec="seconds")
    payload["snapshot"] = snap.get("snapshot") if snap else None
    payload["written"] = written
    payload["failed"] = failed
    with open(full, "w", encoding="utf-8") as fh:
        json.dump(payload, fh, ensure_ascii=False, indent=1)

    return {
        "changeset": full,
        "snapshot": snap.get("snapshot") if snap else None,
        "written": len(written),
        "failed": failed,
        "cells": written,
        "note": "Workbook is edited but NOT saved - review it, then excel_save.",
    }


@com_tool
def excel_trace_precedents(
    cell: str,
    handle: str | None = None,
    sheet: str | None = None,
    depth: int = 3,
    max_nodes: int = 120,
) -> dict[str, Any]:
    """Walk the precedent tree of a cell: what drives this number, recursively.

    Answers "which assumptions does this EPS actually depend on". Cross-sheet
    precedents are reported but not expanded (Excel exposes them only through
    the audit arrows, which need the sheet active).
    """
    wb = _wb(handle)
    ws = _sheet(wb, sheet)
    root = ws.Range(cell)
    seen: set[str] = set()
    nodes: list[dict[str, Any]] = []
    leaves: list[dict[str, Any]] = []

    def visit(rng: Any, level: int) -> None:
        if len(nodes) >= max_nodes:
            return
        key = "%s!%s" % (rng.Worksheet.Name, _addr(rng))
        if key in seen:
            return
        seen.add(key)
        has_formula = bool(rng.HasFormula)
        entry = {
            "cell": key,
            "level": level,
            "value": jsonable(rng.Value),
            "formula": jsonable(rng.Formula) if has_formula else None,
            "is_input": not has_formula,
        }
        if not has_formula:
            hexc = _hex_from_bgr(rng.Font.Color)
            entry["colour"] = hexc
            entry["class"] = _classify(hexc)
            leaves.append(entry)
        nodes.append(entry)
        if not has_formula or level >= depth:
            return
        try:
            precedents = rng.Precedents
        except Exception:  # noqa: BLE001 - no precedents on this sheet
            return
        for area in precedents.Areas:
            if int(area.Count) > 200:
                continue
            for child in area:
                visit(child, level + 1)

    def work() -> dict[str, Any]:
        visit(root, 0)
        return {
            "root": "%s!%s" % (ws.Name, _addr(root)),
            "nodes": len(nodes),
            "truncated": len(nodes) >= max_nodes,
            "tree": nodes,
            "input_leaves": leaves,
            "note": "Precedents on other sheets are not expanded by Excel's "
                    "object model unless that sheet is active.",
        }

    return work()


@com_tool
def excel_snapshot(
    handle: str | None = None,
    path: str | None = None,
    output_dir: str | None = None,
    label: str | None = None,
) -> dict[str, Any]:
    """Save a timestamped copy of a workbook before changing it.

    This is what makes an edit safe to approve: the copy is a plain file next to
    the original (or in output_dir), so recovery never depends on this server.
    """
    wb = _wb(handle) if (handle or not path) else None
    if wb is None:
        opened = excel_open(path=path, read_only=False, visible=False,
                            update_links=False)
        if not opened.get("ok", True):
            return opened
        wb = _wb(opened["handle"])
    source = str(getattr(wb, "FullName", ""))
    if not source or not os.path.dirname(source):
        raise ComToolError(
            "Workbook has never been saved, so there is nothing to snapshot",
            hint="Save it once with excel_save(path=...) first.",
        )
    base, ext = os.path.splitext(os.path.basename(source))
    stamp = _dt.datetime.now().strftime("%Y%m%d-%H%M%S")
    suffix = ("-" + re.sub(r"[^\w.-]", "_", label)) if label else ""
    target_dir = os.path.abspath(os.path.expanduser(output_dir)) \
        if output_dir else os.path.dirname(source)
    os.makedirs(target_dir, exist_ok=True)
    target = os.path.join(target_dir, "%s.snapshot-%s%s%s" % (base, stamp, suffix, ext))

    def work() -> dict[str, Any]:
        wb.SaveCopyAs(target)
        return {
            "source": source,
            "snapshot": target,
            "exists": os.path.exists(target),
            "bytes": os.path.getsize(target) if os.path.exists(target) else 0,
        }

    return work()


TOOLS = [
    notes_read,
    pdf_read_text,
    excel_model_map,
    excel_trace_precedents,
    excel_snapshot,
    excel_propose_changes,
    excel_apply_changeset,
]
