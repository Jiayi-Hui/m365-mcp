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


def _as_text(value: Any) -> Any:
    """Make a string safe to put in a cell as literal text.

    Excel parses anything starting with = + - @ as a formula, so a note that
    quotes a formula ("=Y31/1.0614, back-solved from 1H25") raises 0x800A03EC
    on assignment. A leading apostrophe forces text and is not displayed.
    """
    if isinstance(value, str) and value[:1] in ("=", "+", "-", "@"):
        return "'" + value
    return value


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


def _block_values_attr(rng: Any, rows: int, cols: int, attr: str) -> list[list]:
    """Read Formula / Text / Value for a whole range in ONE COM round trip."""
    try:
        raw = getattr(rng, attr)
    except Exception:  # noqa: BLE001
        return [[None] * cols for _ in range(rows)]
    if rows == 1 and cols == 1:
        return [[raw]]
    if not isinstance(raw, tuple):
        return [[raw] * cols for _ in range(rows)]
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


# --- external data add-ins -------------------------------------------------

# A formula that calls one of these is LIVE data, not a number someone typed.
# It must never be overwritten by hand: the add-in owns that cell.
ADDIN_PATTERNS = {
    "wind": (
        re.compile(r"windfunc|wind\.net|\bWSD\(|\bWSS\(|\bWSET\(|\bEDB\(|"
                   r"\bS_STM\d|\bW_[A-Z]", re.I),
        ("wind",),
    ),
    "bloomberg": (
        re.compile(r"bloombergui|\bBDP\(|\bBDH\(|\bBDS\(|\bBQL\(|\bBSRCH\(", re.I),
        ("bloomberg", "blp"),
    ),
    "factset": (re.compile(r"factset|\bFDS[A-Z_(]|\bFQL\(", re.I), ("factset", "fds")),
    "capiq": (re.compile(r"capiq|\bCIQ[A-Z_(]", re.I), ("capiq", "capital iq")),
}

_ERROR_TEXTS = ("#NAME?", "#REF!", "#DIV/0!", "#VALUE!", "#N/A", "#NULL!",
                "#NUM!", "#SPILL!", "#CALC!")

# An Excel error cell read over COM arrives as a large negative integer
# (#REF! is -2146826265, #NAME? -2146826259, ...). Treated as a number it
# poisons every sum, median and ratio downstream.
_COM_ERROR_FLOOR = -2146800000


def _is_com_error(value: Any) -> bool:
    return (isinstance(value, int) and not isinstance(value, bool)
            and value < _COM_ERROR_FLOOR)


def _loaded_addins(app: Any) -> dict[str, list[str]]:
    """Which data add-ins are actually live in THIS Excel instance."""
    found: dict[str, list[str]] = {}
    try:
        for addin in app.AddIns:
            try:
                if not bool(addin.Installed):
                    continue
                name = str(addin.Name).lower()
            except Exception:  # noqa: BLE001
                continue
            for vendor, (_rx, keys) in ADDIN_PATTERNS.items():
                if any(k in name for k in keys):
                    found.setdefault(vendor, []).append(str(addin.Name))
    except Exception:  # noqa: BLE001
        pass
    try:
        for com in app.COMAddIns:
            try:
                if not bool(com.Connect):
                    continue
                desc = str(com.Description).lower()
            except Exception:  # noqa: BLE001
                continue
            for vendor, (_rx, keys) in ADDIN_PATTERNS.items():
                if any(k in desc for k in keys):
                    found.setdefault(vendor, []).append(str(com.Description))
    except Exception:  # noqa: BLE001
        pass
    return found


def _classify_formula_vendor(formula: str) -> str | None:
    for vendor, (rx, _keys) in ADDIN_PATTERNS.items():
        if rx.search(formula):
            return vendor
    return None


_RATIO_WORDS = ("%", "margin", "ratio", "yoy", "growth", "per ", "asp", "eps",
                "yield", "rate", "利润率", "毛利率", "增速", "同比", "占比")


def _is_ratio_row(label: str) -> bool:
    """A rate does not add up across periods; only a flow does."""
    low = label.lower()
    return any(w in low for w in _RATIO_WORDS)


def _looks_averaged(total: float, fy: float, parts: int = 2) -> bool:
    """total ~= fy * parts means a rate was summed by mistake, not a flow."""
    if abs(fy) < 1e-9:
        return False
    return abs(total / fy - parts) < 0.25


# --- period arithmetic (for the consistency checks) ------------------------

_FY_RE = re.compile(r"^(?:FY)?\s*((?:19|20)?\d{2})\s*[EAF]?$", re.I)
_H_RE = re.compile(r"^([12])\s*H\s*(?:FY)?\s*((?:19|20)?\d{2})\s*[EAF]?$", re.I)
_Q_RE = re.compile(r"^([1-4])\s*Q\s*(?:FY)?\s*((?:19|20)?\d{2})\s*[EAF]?$", re.I)


def _year4(text: str) -> int:
    n = int(text)
    if n < 100:
        return 2000 + n if n < 80 else 1900 + n
    return n


def _parse_period(text: Any) -> tuple[str, int, int] | None:
    """'2025' -> ('FY',2025,0); '1H25' -> ('H',2025,1); '3Q26E' -> ('Q',2026,3)."""
    if text is None:
        return None
    s = str(text).strip()
    if not s:
        return None
    m = _H_RE.match(s)
    if m:
        return ("H", _year4(m.group(2)), int(m.group(1)))
    m = _Q_RE.match(s)
    if m:
        return ("Q", _year4(m.group(2)), int(m.group(1)))
    m = _FY_RE.match(s)
    if m:
        year = _year4(m.group(1))
        if 1990 <= year <= 2100:
            return ("FY", year, 0)
    return None


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


@com_tool(timeout=600)
def excel_data_sources(
    path: str | None = None,
    handle: str | None = None,
    max_samples: int = 5,
) -> dict[str, Any]:
    """Which live data add-ins does this workbook depend on, and are they loaded?

    A cell whose formula calls Wind, Bloomberg, FactSet or CapIQ is owned by
    that add-in: it refreshes itself and must never be typed over. Worse, if the
    add-in is NOT loaded in this Excel, every one of those cells reads #NAME? -
    and saving the file in that state writes the errors in permanently, losing
    the cached values.

    Use this before editing an unfamiliar model. `excel_apply_changeset` calls
    it automatically and refuses to write when the verdict is unsafe.
    """
    if not path and not handle:
        raise ComToolError("Pass path= or handle=")
    if path and not handle:
        opened = excel_open(path=path, read_only=True, visible=False,
                            update_links=False)
        if not opened.get("ok", True):
            return opened
        handle = opened["handle"]

    wb = _wb(handle)
    app = wb.Application
    loaded = _loaded_addins(app)

    by_vendor: dict[str, dict[str, Any]] = {}
    by_sheet: dict[str, int] = {}
    errors: dict[str, int] = {}
    error_from_addin = 0
    error_other: list[dict[str, Any]] = []

    for ws in wb.Worksheets:
        sheet_hits = 0
        try:
            formulas = ws.UsedRange.SpecialCells(-4123)  # xlCellTypeFormulas
        except Exception:  # noqa: BLE001 - no formulas here
            formulas = None
        if formulas is not None:
            for area in formulas.Areas:
                a_rows, a_cols = int(area.Rows.Count), int(area.Columns.Count)
                if a_rows * a_cols > 200000:
                    continue
                a_top, a_left = int(area.Row), int(area.Column)
                # One round trip for the whole block. Cell by cell this was
                # ~23k COM calls on a single valuation sheet.
                fx = _block_values_attr(area, a_rows, a_cols, "Formula")
                vals = _block_values_attr(area, a_rows, a_cols, "Value")
                for ri in range(len(fx)):
                    for ci in range(len(fx[ri])):
                        formula = fx[ri][ci]
                        if not isinstance(formula, str) or not formula:
                            continue
                        vendor = _classify_formula_vendor(formula)
                        if not vendor:
                            continue
                        sheet_hits += 1
                        rec = by_vendor.setdefault(
                            vendor, {"cells": 0, "errored": 0, "samples": []})
                        rec["cells"] += 1
                        raw = (vals[ri][ci] if ri < len(vals) and ci < len(vals[ri])
                               else None)
                        errored = _is_com_error(raw)
                        if errored:
                            rec["errored"] += 1
                            error_from_addin += 1
                        if len(rec["samples"]) < max_samples:
                            rec["samples"].append({
                                "cell": "%s!%s%d" % (ws.Name,
                                                     _col_letter(a_left + ci),
                                                     a_top + ri),
                                "formula": formula[:120],
                                "errored": errored,
                            })
        if sheet_hits:
            by_sheet[str(ws.Name)] = sheet_hits

        # error census, so add-in-caused errors can be told apart from real breaks
        try:
            errs = ws.UsedRange.SpecialCells(-4123, 16)
        except Exception:  # noqa: BLE001
            continue
        for area in errs.Areas:
            a_rows, a_cols = int(area.Rows.Count), int(area.Columns.Count)
            a_top, a_left = int(area.Row), int(area.Column)
            if a_rows * a_cols > 200000:
                errors["(large block)"] = errors.get(
                    "(large block)", 0) + int(area.Count)
                continue
            fx = _block_values_attr(area, a_rows, a_cols, "Formula")
            tx = _block_values_attr(area, a_rows, a_cols, "Text")
            for ri in range(len(fx)):
                for ci in range(len(fx[ri])):
                    text = tx[ri][ci] if ri < len(tx) and ci < len(tx[ri]) else None
                    text = str(text) if text is not None else ""
                    if text not in _ERROR_TEXTS:
                        continue
                    errors[text] = errors.get(text, 0) + 1
                    formula = fx[ri][ci] if isinstance(fx[ri][ci], str) else ""
                    if not _classify_formula_vendor(formula) and len(error_other) < 20:
                        error_other.append({
                            "cell": "%s!%s%d" % (ws.Name, _col_letter(a_left + ci),
                                                 a_top + ri),
                            "error": text,
                            "formula": formula[:90],
                        })

    missing = [v for v in by_vendor if v not in loaded]
    verdict = "safe_to_edit"
    reasons = []
    if missing:
        verdict = "read_only_recommended"
        for vendor in missing:
            reasons.append(
                "%s formulas present (%d cells, %d currently #ERROR) but the %s "
                "add-in is not loaded in this Excel - saving would bake the "
                "errors in" % (vendor, by_vendor[vendor]["cells"],
                               by_vendor[vendor]["errored"], vendor)
            )
    return {
        "workbook": str(getattr(wb, "FullName", wb.Name)),
        "handle": handle,
        "verdict": verdict,
        "reasons": reasons,
        "addins_loaded": loaded,
        "addins_required": sorted(by_vendor),
        "addins_missing": missing,
        "formula_cells_by_vendor": {
            v: {"cells": d["cells"], "errored": d["errored"], "samples": d["samples"]}
            for v, d in by_vendor.items()
        },
        "sheets_with_external_formulas": by_sheet,
        "error_cells_by_type": errors,
        "error_cells_from_addins": error_from_addin,
        "errors_not_explained_by_addins": error_other,
        "reading": (
            "Cells owned by an add-in refresh themselves - never write to them. "
            "Errors listed under errors_not_explained_by_addins are real breaks "
            "(usually #REF! from a deleted range) and will not fix themselves "
            "by loading an add-in."
        ),
    }


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
    force: bool = False,
) -> dict[str, Any]:
    """Apply a reviewed changeset to the workbook. Requires confirm=True.

    Takes a snapshot first, writes only rows the review passed, and leaves the
    reason and source in a cell comment so the edit stays auditable months later.
    Rows marked `blocked` are never written.

    It also refuses outright if the workbook depends on a data add-in (Wind,
    Bloomberg, ...) that is not loaded in this Excel - in that state the add-in
    cells read #NAME? and saving destroys their cached values. `force=true`
    overrides that, and is only safe if you will not save.
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

    # Never write into a workbook whose live-data add-ins are missing: the
    # add-in cells currently read #NAME?, and saving would bake that in.
    sources = excel_data_sources(handle=handle or payload.get("workbook"))
    if sources.get("verdict") == "read_only_recommended" and not force:
        return {
            "ok": False,
            "error": "workbook depends on add-ins that are not loaded: %s" % (
                ", ".join(sources.get("addins_missing", []))),
            "reasons": sources.get("reasons", []),
            "hint": "Open this model on a machine where those add-ins are "
                    "loaded and logged in, or pass force=true if you are "
                    "certain you will not save the file.",
        }

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


@com_tool(timeout=600)
def excel_proposal_sheet(
    proposals: list,
    handle: str | None = None,
    sheet_name: str = "AI_Proposals",
    visible: bool = True,
    replace: bool = True,
) -> dict[str, Any]:
    """Build an approval sheet inside the workbook: one row per proposed edit,
    with a tick box the analyst fills in without leaving Excel.

    A comment tells you something; it gives you nothing to approve. This lays
    the proposals out as a table - target cell (hyperlinked), what the cell
    means, current value, proposed value, the basis, the derivation, the source,
    and how many cells depend on it - with an APPROVE column to mark `Y`.
    `excel_apply_approved` then applies only the ticked rows.

    Each proposal:
        {"sheet": "Breakdown", "cell": "Y15", "proposed": 261.65,
         "basis": "arithmetic",          # arithmetic | rule | judgement
         "label": "Outdoor footwear x 1H25",
         "derivation": "194.79 / (1 - 25.55%) = 261.64; segments must sum to 12,649.04",
         "source": "2026H1 report p15"}

    `basis` is the honest part. `arithmetic` has one correct answer and can be
    approved on sight; `rule` carries an assumption that must be agreed with
    first; `judgement` should not carry a number at all - leave `proposed` empty
    and let the row state the reconciliation instead.

    Nothing is written to the model and the file is not saved.
    """
    if not isinstance(proposals, list) or not proposals:
        raise ComToolError("proposals must be a non-empty list")
    wb = _wb(handle)
    app = wb.Application
    try:
        app.Visible = bool(visible)
    except Exception:  # noqa: BLE001
        pass

    existing = None
    for ws in wb.Worksheets:
        if str(ws.Name).lower() == sheet_name.lower():
            existing = ws
            break
    if existing is not None:
        if not replace:
            raise ComToolError(
                "Sheet %r already exists" % sheet_name,
                hint="Pass replace=true to rebuild it.")
        existing.Delete()
    ws = wb.Worksheets.Add(Before=wb.Worksheets(1))
    ws.Name = sheet_name

    headers = ["#", "APPROVE", "Target", "What it is", "Current", "Proposed",
               "Basis", "Derivation / reconciliation", "Source", "Dependents",
               "Status"]
    for c, text in enumerate(headers, start=1):
        cell = ws.Cells(1, c)
        cell.Value = text
        cell.Font.Bold = True
        cell.Interior.Color = 0x703030          # BGR: dark slate
        cell.Font.Color = 0xFFFFFF

    basis_colour = {"arithmetic": 0xD9F2D9,     # green - safe to approve
                    "rule": 0x9CDCFF,           # amber - agree the rule first
                    "judgement": 0xF0F0F0}      # grey  - no number offered
    rows = []
    for i, item in enumerate(proposals, start=1):
        if not isinstance(item, dict) or not item.get("cell"):
            continue
        r = i + 1
        target_sheet = str(item.get("sheet") or "")
        target = "%s!%s" % (target_sheet, item["cell"]) if target_sheet \
            else str(item["cell"])
        current = None
        dependents = None
        has_formula = None
        try:
            tws = _sheet(wb, target_sheet or None)
            rng = tws.Range(item["cell"])
            current = jsonable(rng.Value)
            has_formula = bool(rng.HasFormula)
            if has_formula:
                current = jsonable(rng.Formula)
            try:
                dependents = int(rng.Dependents.Count)
            except Exception:  # noqa: BLE001
                dependents = 0
        except Exception as exc:  # noqa: BLE001
            current = "?? %s" % str(exc)[:40]

        basis = str(item.get("basis", "judgement")).lower()
        ws.Cells(r, 1).Value = i
        ws.Cells(r, 2).Value = ""
        ws.Cells(r, 3).Value = target
        try:
            ws.Hyperlinks.Add(Anchor=ws.Cells(r, 3), Address="",
                              SubAddress=target, TextToDisplay=target)
        except Exception:  # noqa: BLE001 - sheet names with spaces etc.
            pass
        ws.Cells(r, 4).Value = _as_text(item.get("label"))
        ws.Cells(r, 5).Value = _as_text(current)
        proposed = item.get("proposed")
        ws.Cells(r, 6).Value = "" if proposed is None else proposed
        ws.Cells(r, 7).Value = basis
        ws.Cells(r, 8).Value = _as_text(item.get("derivation"))
        ws.Cells(r, 9).Value = _as_text(item.get("source"))
        ws.Cells(r, 10).Value = dependents
        ws.Cells(r, 11).Value = "pending" if proposed is not None else "info only"
        colour = basis_colour.get(basis, 0xF0F0F0)
        ws.Range(ws.Cells(r, 1), ws.Cells(r, 11)).Interior.Color = colour
        if has_formula:
            ws.Cells(r, 5).Font.Italic = True
        rows.append({"row": r, "target": target, "basis": basis,
                     "proposed": jsonable(proposed), "dependents": dependents,
                     "is_formula": has_formula})

    # tick box: a two-value dropdown beats free text
    if rows:
        box = ws.Range(ws.Cells(2, 2), ws.Cells(len(rows) + 1, 2))
        try:
            box.Validation.Delete()
            box.Validation.Add(Type=3, AlertStyle=1, Operator=1, Formula1="Y,N")
            box.Validation.InCellDropdown = True
        except Exception:  # noqa: BLE001
            pass
        box.HorizontalAlignment = -4108
        box.Font.Bold = True

    # Cosmetics. None of this is worth failing the whole call for, and Excel
    # rejects some of it depending on the sheet's state (AutoFilter in
    # particular), so each step stands alone.
    cosmetic_errors = []
    for c, width in ((1, 5), (2, 10), (3, 20), (4, 34), (5, 16), (6, 14),
                     (7, 12), (8, 62), (9, 30), (10, 11), (11, 12)):
        try:
            ws.Columns(c).ColumnWidth = width
        except Exception as exc:  # noqa: BLE001
            cosmetic_errors.append("width col %d: %s" % (c, str(exc)[:60]))
    for step, action in (
        ("wrap", lambda: setattr(ws.Range("H:H"), "WrapText", True)),
        ("autofilter", lambda: ws.Range(
            ws.Cells(1, 1), ws.Cells(max(2, len(rows) + 1), 11)).AutoFilter()),
        ("activate", lambda: ws.Activate()),
        ("freeze", lambda: (setattr(app.ActiveWindow, "SplitRow", 1),
                            setattr(app.ActiveWindow, "FreezePanes", True))),
    ):
        try:
            action()
        except Exception as exc:  # noqa: BLE001
            cosmetic_errors.append("%s: %s" % (step, str(exc)[:60]))

    counts: dict[str, int] = {}
    for row in rows:
        counts[row["basis"]] = counts.get(row["basis"], 0) + 1
    return {
        "sheet": sheet_name,
        "rows": rows,
        "by_basis": counts,
        "cosmetic_warnings": cosmetic_errors,
        "saved": False,
        "how_to_use": (
            "Open the %s tab, put Y in the APPROVE column for the rows you "
            "accept, then ask to apply the approved proposals. Green rows are "
            "arithmetic - one correct answer, safe to approve on sight. Amber "
            "rows carry an assumption stated in the Derivation column: agree "
            "with that first. Grey rows offer no number on purpose." % sheet_name
        ),
    }


@com_tool(timeout=600)
def excel_apply_approved(
    handle: str | None = None,
    sheet_name: str = "AI_Proposals",
    confirm: bool = False,
    snapshot: bool = True,
    annotate: bool = True,
    force: bool = False,
) -> dict[str, Any]:
    """Apply the proposal rows the analyst ticked `Y`, and nothing else.

    Runs the same guards as excel_apply_changeset: a formula target is refused,
    a missing data add-in refuses the whole run, a snapshot is taken first, and
    each edited cell keeps a comment with the derivation, the source and its
    previous value. The Status column is written back so the sheet becomes the
    record of what was done.
    """
    wb = _wb(handle)
    ws = None
    for candidate in wb.Worksheets:
        if str(candidate.Name).lower() == sheet_name.lower():
            ws = candidate
            break
    if ws is None:
        raise ComToolError(
            "No %r sheet in this workbook" % sheet_name,
            hint="Build one with excel_proposal_sheet first.")

    last = int(ws.UsedRange.Rows.Count)
    approved, skipped = [], []
    for r in range(2, last + 1):
        target = ws.Cells(r, 3).Value
        if not target:
            continue
        mark = str(ws.Cells(r, 2).Value or "").strip().upper()
        proposed = ws.Cells(r, 6).Value
        if mark != "Y":
            skipped.append({"row": r, "target": str(target),
                            "reason": "not approved" if mark != "N" else "marked N"})
            continue
        if proposed is None or proposed == "":
            skipped.append({"row": r, "target": str(target),
                            "reason": "no proposed value on this row"})
            continue
        approved.append({
            "row": r, "target": str(target), "proposed": proposed,
            "basis": str(ws.Cells(r, 7).Value or ""),
            "derivation": str(ws.Cells(r, 8).Value or ""),
            "source": str(ws.Cells(r, 9).Value or ""),
        })

    if not confirm:
        return {
            "ok": False,
            "error": "confirm=False - nothing was written",
            "approved_rows": len(approved),
            "skipped_rows": len(skipped),
            "preview": [
                {"target": a["target"], "to": jsonable(a["proposed"]),
                 "basis": a["basis"]} for a in approved[:20]],
            "hint": "Show the preview, then call again with confirm=true.",
        }
    if not approved:
        return {"written": 0, "approved_rows": 0, "skipped": skipped,
                "note": "No row was marked Y."}

    sources = excel_data_sources(handle=handle)
    if sources.get("verdict") == "read_only_recommended" and not force:
        return {
            "ok": False,
            "error": "workbook depends on add-ins that are not loaded: %s"
                     % ", ".join(sources.get("addins_missing", [])),
            "reasons": sources.get("reasons", []),
            "hint": "Applying is fine in memory, but this workbook must not be "
                    "saved in this state. Load the add-in, or pass force=true "
                    "if you will not save.",
        }

    snap = None
    if snapshot:
        snap = excel_snapshot(handle=handle, label="pre-approved")
        if not snap.get("ok", True):
            return snap

    written, failed = [], []
    for item in approved:
        target = item["target"]
        try:
            sheet_ref, _, cell_ref = target.rpartition("!")
            tws = _sheet(wb, sheet_ref or None)
            rng = tws.Range(cell_ref)
            if bool(rng.HasFormula):
                failed.append({"target": target,
                               "error": "target holds a formula: %s"
                                        % str(rng.Formula)[:70]})
                ws.Cells(item["row"], 11).Value = "refused: formula"
                continue
            before = jsonable(rng.Value)
            rng.Value = item["proposed"]
            if annotate:
                text = "[AI] %s applied\n%s\nsource: %s\nwas: %s" % (
                    _dt.datetime.now().strftime("%Y-%m-%d"),
                    item["derivation"] or "(no derivation recorded)",
                    item["source"] or "(no source recorded)", before)
                try:
                    if rng.Comment is not None:
                        rng.Comment.Delete()
                except Exception:  # noqa: BLE001
                    pass
                try:
                    rng.AddComment(text[:1200])
                except Exception:  # noqa: BLE001
                    pass
            ws.Cells(item["row"], 11).Value = "applied %s" % \
                _dt.datetime.now().strftime("%H:%M")
            written.append({"target": target, "from": before,
                            "to": jsonable(item["proposed"])})
        except Exception as exc:  # noqa: BLE001
            failed.append({"target": target, "error": str(exc)[:160]})
            ws.Cells(item["row"], 11).Value = "failed"

    return {
        "written": len(written),
        "cells": written,
        "failed": failed,
        "skipped": skipped,
        "snapshot": snap.get("snapshot") if snap else None,
        "note": "Workbook edited in memory and NOT saved. Recalculate and check "
                "the outputs before deciding whether to save.",
    }


@com_tool(timeout=600)
def excel_annotate(
    annotations: list,
    handle: str | None = None,
    path: str | None = None,
    highlight: bool = True,
    visible: bool = True,
    prefix: str = "[AI]",
    replace_existing: bool = True,
    on_existing: str = "skip",
    save: bool = False,
) -> dict[str, Any]:
    """Put the agent's findings ON the cells, in the analyst's own Excel window.

    This is the delivery step for model work. A model is maintained by looking
    at the grid, so a finding that lives in a chat log or a separate document is
    a finding the analyst has to re-locate by hand. Here it arrives attached to
    the cell it is about: hover, read the evidence, decide, edit in place.

    Each annotation is a dict:
        {"sheet": "Breakdown", "cell": "Y15",
         "text": "1H25 should be ~261.65, not 2,616.49 (10x)...",
         "severity": "high"}

    **It writes comments and cell colour, never values, and by default it does
    not save.** That matters on a machine whose data add-in is missing: the
    workbook stays open with the annotations visible, while the file on disk is
    untouched, so the add-in's cached values cannot be destroyed. The analyst
    reads the notes in the live window and decides what to do.

    `on_existing` governs cells that already carry somebody else's note - and in
    a maintained model there are many: sourcing ("ML forecast", "Adidas annual
    report"), call notes dated by month, fiscal-calendar definitions. Those took
    real work and their Author is part of the record, so the default `skip`
    leaves them completely untouched (the cell is still highlighted, and the
    collision is reported back). `append` adds below the original text, losing
    the original Author; nothing ever overwrites another person's note.
    """
    if not isinstance(annotations, list) or not annotations:
        raise ComToolError("annotations must be a non-empty list")
    if path and not handle:
        opened = excel_open(path=path, read_only=False, visible=visible,
                            update_links=False)
        if not opened.get("ok", True):
            return opened
        handle = opened["handle"]

    wb = _wb(handle)
    try:
        wb.Application.Visible = bool(visible)
    except Exception:  # noqa: BLE001
        pass

    severity_colour = {          # Interior colour is BGR
        "high": 0x9CA0FF,        # soft red
        "medium": 0x9CDCFF,      # soft amber
        "low": 0xD9F2D9,         # soft green
        "info": 0xF2E6D9,
    }
    if on_existing not in ("skip", "append"):
        raise ComToolError("on_existing must be 'skip' or 'append'")
    written, failed, skipped = [], [], []
    for item in annotations:
        if not isinstance(item, dict) or not item.get("cell"):
            failed.append({"item": jsonable(item), "error": "needs a 'cell'"})
            continue
        try:
            ws = _sheet(wb, item.get("sheet"))
            rng = ws.Range(item["cell"])
            severity = str(item.get("severity", "info")).lower()
            body = "%s %s\n%s" % (
                prefix, _dt.datetime.now().strftime("%Y-%m-%d"),
                str(item.get("text", "")).strip())
            if item.get("source"):
                body += "\nsource: %s" % item["source"]
            existing = None
            try:
                existing = rng.Comment
            except Exception:  # noqa: BLE001
                existing = None
            if existing is not None:
                old = str(existing.Text())
                author = ""
                try:
                    author = str(existing.Author or "")
                except Exception:  # noqa: BLE001
                    pass
                mine = prefix in old
                if mine and replace_existing:
                    existing.Delete()
                elif mine:
                    body = old + "\n---\n" + body
                    existing.Delete()
                else:
                    # An analyst's own note. These carry sourcing, call notes and
                    # definitions that took real work; rewriting the comment
                    # would also drop its Author. Leave it completely alone.
                    if on_existing == "skip":
                        skipped.append({
                            "cell": "%s!%s" % (ws.Name, item["cell"]),
                            "existing_author": author,
                            "existing_text": old.strip()[:160],
                            "reason": "cell already has a note from someone else; "
                                      "not overwritten",
                        })
                        if highlight and severity in severity_colour:
                            rng.Interior.Color = severity_colour[severity]
                        continue
                    body = old.rstrip() + "\n---\n" + body
                    existing.Delete()
            rng.AddComment(body[:1800])
            try:
                rng.Comment.Shape.TextFrame.AutoSize = True
            except Exception:  # noqa: BLE001
                pass
            if highlight and severity in severity_colour:
                rng.Interior.Color = severity_colour[severity]
            written.append({
                "cell": "%s!%s" % (ws.Name, item["cell"]),
                "severity": severity,
                "current_value": jsonable(rng.Value),
                "has_formula": bool(rng.HasFormula),
            })
        except Exception as exc:  # noqa: BLE001
            failed.append({"cell": item.get("cell"), "error": str(exc)[:160]})

    saved = False
    if save:
        sources = excel_data_sources(handle=handle)
        if sources.get("verdict") == "read_only_recommended":
            return {
                "annotated": written,
                "failed": failed,
                "skipped_existing_notes": skipped,
                "saved": False,
                "ok": False,
                "error": "annotations applied in memory but NOT saved: %s"
                         % ", ".join(sources.get("addins_missing", [])),
                "hint": "Saving now would bake the add-in's #NAME? cells into "
                        "the file. Leave the window open and review there.",
            }
        wb.Save()
        saved = True

    return {
        "workbook": str(getattr(wb, "FullName", wb.Name)),
        "handle": handle,
        "annotated": written,
        "failed": failed,
        "skipped_existing_notes": skipped,
        "saved": saved,
        "note": "Comments and highlighting are in the open workbook; the file "
                "on disk is unchanged. Review in Excel, then save yourself if "
                "you want to keep them. Cells already carrying someone else's "
                "note were highlighted but not written to.",
    }


@com_tool(timeout=300)
def excel_clear_annotations(
    handle: str | None = None,
    sheets: str | None = None,
    prefix: str = "[AI]",
    clear_highlight: bool = True,
) -> dict[str, Any]:
    """Remove annotations this server added, leaving anyone else's comments."""
    wb = _wb(handle)
    wanted = None
    if sheets:
        wanted = {s.strip().lower() for s in sheets.split(",") if s.strip()}
    removed = []
    for ws in wb.Worksheets:
        if wanted and str(ws.Name).lower() not in wanted:
            continue
        try:
            comments = ws.Comments
            count = int(comments.Count)
        except Exception:  # noqa: BLE001
            continue
        for i in range(count, 0, -1):
            try:
                comment = comments(i)
                if prefix not in str(comment.Text()):
                    continue
                cell = comment.Parent
                addr = "%s!%s" % (ws.Name, str(cell.Address).replace("$", ""))
                comment.Delete()
                if clear_highlight:
                    cell.Interior.ColorIndex = -4142  # xlColorIndexNone
                removed.append(addr)
            except Exception:  # noqa: BLE001
                continue
    return {"removed": removed, "count": len(removed),
            "note": "Workbook not saved; close without saving to revert fully."}


@com_tool(timeout=900)
def excel_model_check(
    path: str | None = None,
    handle: str | None = None,
    sheets: str | None = None,
    tolerance: float = 0.005,
    jump_factor: float = 10.0,
    max_findings: int = 60,
) -> dict[str, Any]:
    """Health check a model before trusting it with new data.

    Four checks, each aimed at a failure that produces no error message:

    1. **Error cells**, split into add-in-caused (#NAME? because Wind/Bloomberg
       is not loaded - fixable by loading it) and real breaks (#REF! from a
       deleted range - not fixable that way).
    2. **Period arithmetic**: for every row, does FY equal 1H+2H, and does it
       equal the four quarters? A mismatch means one of the three is wrong.
    3. **Subtotal rows**: where a row equals the sum of the rows beneath it in
       most columns, report the columns where it does not. This is what catches
       a mistyped component that nothing else checks.
    4. **Magnitude jumps**: a value far ABOVE its own row's median - the
       classic misplaced decimal point. Only the upside is flagged; a small
       period is ordinary. This is a weak heuristic compared to check 3: a
       10x error inside a volatile row may sit under the threshold, which is
       exactly why the subtotal check exists.

    Iterative calculation is also reported: it usually means a circular
    reference someone chose to live with.
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

    wb = _wb(handle)
    app = wb.Application
    findings: list[dict[str, Any]] = []
    checked_sheets: list[str] = []

    quota = {"period_sum": max(8, max_findings // 4),
             "subtotal": max_findings,
             "magnitude": max(10, max_findings // 3)}
    used_quota: dict[str, int] = {}

    def add(kind: str, severity: str, sheet: str, detail: dict[str, Any]) -> None:
        # Per-check quotas. period_sum is by far the chattiest check; without
        # this it fills the whole budget and the subtotal/magnitude findings -
        # the ones that catch silent data errors - are never recorded at all.
        if used_quota.get(kind, 0) >= quota.get(kind, max_findings):
            return
        if len(findings) >= max_findings * 3:
            return
        used_quota[kind] = used_quota.get(kind, 0) + 1
        findings.append({"check": kind, "severity": severity,
                         "sheet": sheet, **detail})

    for ws in wb.Worksheets:
        name = str(ws.Name)
        if wanted and name.lower() not in wanted:
            continue
        used = ws.UsedRange
        n_rows, n_cols = int(used.Rows.Count), int(used.Columns.Count)
        if n_rows * n_cols > 120000 or n_rows < 3:
            continue

        top, left = int(used.Row), int(used.Column)
        band = _block_values(ws, top, left, min(12, n_rows), min(80, n_cols))

        # locate the period header row
        header_row, periods = None, {}
        best = 0
        for ri, row_values in enumerate(band):
            parsed = {}
            for ci, value in enumerate(row_values):
                p = _parse_period(_period_text(value))
                if p:
                    parsed[left + ci] = p
            if len(parsed) > best:
                best, header_row, periods = len(parsed), top + ri, parsed
        if best < 3:
            continue
        checked_sheets.append(name)

        body_top = header_row + 1
        body_rows = min(n_rows - (header_row - top) - 1, 220)
        if body_rows < 2:
            continue
        cols = sorted(periods)
        grid = _block_values(ws, body_top, left, body_rows, n_cols)
        labels = [
            next((str(v).strip() for v in row[:4]
                  if isinstance(v, str) and str(v).strip()), "")
            for row in grid
        ]

        def val(ri: int, col: int):
            ci = col - left
            if 0 <= ci < len(grid[ri]):
                v = grid[ri][ci]
                if isinstance(v, (int, float)) and not isinstance(v, bool):
                    if _is_com_error(v):
                        return None
                    return float(v)
            return None

        # --- check 2: FY vs halves vs quarters -----------------------------
        years: dict[int, dict[str, Any]] = {}
        for col, (kind, year, idx) in periods.items():
            slot = years.setdefault(year, {"FY": None, "H": {}, "Q": {}})
            if kind == "FY":
                slot["FY"] = col
            elif kind == "H":
                slot["H"][idx] = col
            else:
                slot["Q"][idx] = col

        for ri in range(len(grid)):
            label = labels[ri]
            if not label or _is_ratio_row(label) or label.strip().lower() in (
                    "check", "checks", "diff", "difference", "校验", "检查"):
                continue
            for year, slot in sorted(years.items()):
                fy_col = slot["FY"]
                if fy_col is None:
                    continue
                fy = val(ri, fy_col)
                # Below this the row is a checksum/rounding artefact, not a
                # flow worth reconciling (models here are in millions).
                if fy is None or abs(fy) < 1.0:
                    continue
                if len(slot["H"]) == 2:
                    parts = [val(ri, slot["H"][i]) for i in (1, 2)]
                    if all(p is not None for p in parts):
                        total = sum(parts)
                        if _looks_averaged(total, fy):
                            continue  # a rate, not a flow
                        if abs(total - fy) > abs(fy) * tolerance:
                            add("period_sum", "high", name, {
                                "row": body_top + ri, "label": label,
                                "period": "FY%d" % year,
                                "detail": "1H+2H = %.2f but FY = %.2f (gap %.2f)" % (
                                    total, fy, total - fy),
                                "cells": ["%s%d" % (_col_letter(slot["H"][i]),
                                                    body_top + ri) for i in (1, 2)],
                            })
                if len(slot["Q"]) == 4:
                    parts = [val(ri, slot["Q"][i]) for i in (1, 2, 3, 4)]
                    if all(p is not None for p in parts):
                        total = sum(parts)
                        if _looks_averaged(total, fy, 4):
                            continue
                        if abs(total - fy) > abs(fy) * tolerance:
                            add("period_sum", "high", name, {
                                "row": body_top + ri, "label": label,
                                "period": "FY%d" % year,
                                "detail": "Q1..Q4 = %.2f but FY = %.2f (gap %.2f)" % (
                                    total, fy, total - fy),
                            })

        # --- check 3: subtotal rows vs their components --------------------
        for ri in range(len(grid)):
            if not labels[ri]:
                continue
            for width in (2, 3, 4, 5):
                comp = range(ri + 1, ri + 1 + width)
                if comp.stop > len(grid):
                    continue
                if not all(labels[c] for c in comp):
                    continue
                agree, disagree = [], []
                for col in cols:
                    total = val(ri, col)
                    parts = [val(c, col) for c in comp]
                    if total is None or any(p is None for p in parts):
                        continue
                    if abs(total) < 1e-9:
                        continue
                    s = sum(parts)
                    if abs(s - total) <= abs(total) * tolerance:
                        agree.append(col)
                    else:
                        disagree.append((col, s, total))
                # a real subtotal relationship holds nearly everywhere
                if len(agree) >= 4 and disagree and len(agree) >= 3 * len(disagree):
                    for col, s, total in disagree[:3]:
                        kind, year, idx = periods[col]
                        period = {"FY": "FY%d", "H": "%dH" % idx + "%d",
                                  "Q": "%dQ" % idx + "%d"}.get(kind, "%d") % year
                        add("subtotal", "high", name, {
                            "row": body_top + ri, "label": labels[ri],
                            "period": period,
                            "components": [labels[c] for c in comp],
                            "detail": "components sum to %.2f but the total row "
                                      "says %.2f (gap %.2f); the same "
                                      "relationship holds in %d other columns" % (
                                          s, total, s - total, len(agree)),
                            "component_cells": ["%s%d" % (_col_letter(col),
                                                          body_top + c)
                                                for c in comp],
                        })
                    break

        # --- check 4: magnitude jumps against the row's own median ---------
        for ri in range(len(grid)):
            if not labels[ri]:
                continue
            series = [(col, val(ri, col)) for col in cols]
            numbers = [abs(v) for _c, v in series if v is not None and abs(v) > 1e-9]
            if len(numbers) < 5:
                continue
            numbers.sort()
            median = numbers[len(numbers) // 2]
            if median <= 0:
                continue
            for col, v in series:
                if v is None or abs(v) < 1e-9:
                    continue
                ratio = abs(v) / median
                if ratio >= jump_factor:
                    cell = ws.Cells(body_top + ri, col)
                    add("magnitude", "medium" if ratio < 20 else "high", name, {
                        "row": body_top + ri, "label": labels[ri],
                        "cell": "%s%d" % (_col_letter(col), body_top + ri),
                        "period": _period_text(ws.Cells(header_row, col).Value),
                        "value": v,
                        "row_median": median,
                        "ratio": round(ratio, 1),
                        "hand_entered": not bool(cell.HasFormula),
                        "detail": "%.4g is %.1fx the row median %.4g%s" % (
                            v, ratio, median,
                            " (hand-entered constant)" if not cell.HasFormula else ""),
                    })

    sources = excel_data_sources(handle=handle)
    summary = {
        "period_sum": sum(1 for f in findings if f["check"] == "period_sum"),
        "subtotal": sum(1 for f in findings if f["check"] == "subtotal"),
        "magnitude": sum(1 for f in findings if f["check"] == "magnitude"),
    }
    return {
        "workbook": str(getattr(wb, "FullName", wb.Name)),
        "handle": handle,
        "sheets_checked": checked_sheets,
        "iterative_calculation": bool(app.Iteration),
        "iteration_note": (
            "Iterative calculation is ON - the model contains a circular "
            "reference that someone chose to tolerate. Results depend on the "
            "iteration settings." if bool(app.Iteration) else None
        ),
        "error_cells_by_type": sources.get("error_cells_by_type", {}),
        "error_cells_from_addins": sources.get("error_cells_from_addins", 0),
        "errors_not_explained_by_addins":
            sources.get("errors_not_explained_by_addins", []),
        "addins_missing": sources.get("addins_missing", []),
        "findings_summary": summary,
        "findings": findings,
        "truncated": len(findings) >= max_findings,
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
    excel_data_sources,
    excel_model_check,
    excel_annotate,
    excel_clear_annotations,
    excel_proposal_sheet,
    excel_apply_approved,
]
