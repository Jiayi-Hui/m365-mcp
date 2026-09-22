"""End-to-end smoke test: drives every tool family against the real apps.

Run:  python tests/smoke_test.py [--skip ppt,outlook,onenote,access]

It creates its files under a temp directory and closes what it opens. Outlook
coverage is READ-ONLY plus one draft that is deleted again - nothing is sent.
"""

from __future__ import annotations

import argparse
import json
import os
import sys
import tempfile
import traceback

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from m365_mcp import (  # noqa: E402
    tools_common as common,
    tools_research as rs,
    tools_excel as xl,
    tools_misc as misc,
    tools_onenote as one,
    tools_outlook as ol,
    tools_ppt as pp,
    tools_word as wd,
)

RESULTS: list[tuple[str, bool, str]] = []
TMP = os.path.join(tempfile.gettempdir(), "m365_mcp_smoke")
os.makedirs(TMP, exist_ok=True)
# Apps that were already running before the test: we must leave those alone.
PRE_RUNNING: set[str] = set()


def check(label: str, result, *, expect_keys: tuple[str, ...] = ()) -> dict:
    ok = isinstance(result, dict) and result.get("ok") is not False
    detail = ""
    if not ok:
        detail = str(result.get("error") if isinstance(result, dict) else result)[:200]
    elif expect_keys:
        missing = [k for k in expect_keys if k not in result]
        if missing:
            ok = False
            detail = "missing keys: " + ", ".join(missing)
    RESULTS.append((label, ok, detail))
    status = "PASS" if ok else "FAIL"
    print("[%s] %s %s" % (status, label, ("- " + detail) if detail else ""))
    return result if isinstance(result, dict) else {}


def _norm(value):
    """Excel hands numbers back as floats; compare on value, not on type."""
    if isinstance(value, list):
        return [_norm(v) for v in value]
    if isinstance(value, bool) or value is None:
        return value
    if isinstance(value, (int, float)):
        return float(value)
    return value


def assert_equal(label: str, actual, expected) -> bool:
    """Round-trip assertion. Key presence is not evidence that a tool works -
    excel_write_range once wrote to the wrong cell and every key was still there.
    """
    ok = _norm(actual) == _norm(expected)
    detail = "" if ok else "got %r, want %r" % (actual, expected)
    RESULTS.append((label, ok, detail))
    print("[%s] %s %s" % ("PASS" if ok else "FAIL", label,
                          ("- " + detail[:200]) if detail else ""))
    return ok


def section(name: str) -> None:
    print("\n=== %s ===" % name)


def test_common() -> None:
    section("common")
    status = check("m365_status", common.m365_status(), expect_keys=("applications",))
    for row in status.get("applications", []):
        print("    %-12s installed=%-5s running=%s" % (
            row["app"], row["installed"], row["running"]))
        if row["running"]:
            PRE_RUNNING.add(row["app"])
    check("m365_list_handles", common.m365_list_handles())


def test_excel() -> None:
    section("excel")
    created = check("excel_new", xl.excel_new(visible=False), expect_keys=("handle",))
    h = created.get("handle")
    if not h:
        return
    try:
        grid = [["name", "qty", "price"], ["alpha", 3, 1.5], ["beta", 7, 2.25]]
        wrote = check("excel_write_range", xl.excel_write_range(
            values=grid, handle=h, start_cell="A1"))
        assert_equal("excel_write_range lands on A1:C3", wrote.get("address"), "A1:C3")
        check("excel_write_range(formula)", xl.excel_write_range(
            values=[["=B2*C2"], ["=B3*C3"]], handle=h, start_cell="D2",
            as_formula=True))
        check("excel_calculate", xl.excel_calculate(handle=h))

        # Round-trip: the written block must come back cell for cell.
        back = check("excel_read_range", xl.excel_read_range(
            handle=h, range_a1="A1:C3", mode="values"), expect_keys=("values",))
        assert_equal("excel round-trip values", back.get("values"), grid)
        assert_equal("excel round-trip address", back.get("address"), "A1:C3")
        computed = xl.excel_read_range(handle=h, range_a1="D2:D3", mode="values")
        assert_equal("excel formulas evaluate", computed.get("values"),
                     [[4.5], [15.75]])
        read = check("excel_read_range(used)", xl.excel_read_range(
            handle=h, mode="values"), expect_keys=("values",))
        print("    values:", json.dumps(read.get("values"), default=str)[:160])
        formulas = check("excel_read_range(formulas)",
                         xl.excel_read_range(handle=h, range_a1="D2:D3",
                                             mode="formulas"))
        print("    formulas:", formulas.get("values"))
        check("excel_format_range", xl.excel_format_range(
            range_a1="A1:D1", handle=h, bold=True, fill_color="#DDEBF7",
            borders=True))
        check("excel_autofit", xl.excel_autofit(handle=h))
        appended = check("excel_append_rows", xl.excel_append_rows(
            values=[["gamma", 5, 3.0]], handle=h))
        assert_equal("excel_append_rows lands below the last row",
                     appended.get("address"), "A4:C4")
        assert_equal("excel_append_rows content",
                     xl.excel_read_range(handle=h, range_a1="A4:C4").get("values"),
                     [["gamma", 5, 3.0]])
        check("excel_add_sheet", xl.excel_add_sheet(name="Notes", handle=h))
        check("excel_list_sheets", xl.excel_list_sheets(handle=h),
              expect_keys=("sheets",))
        check("excel_rename_sheet", xl.excel_rename_sheet(
            sheet="Notes", new_name="Meta", handle=h))
        check("excel_set_name", xl.excel_set_name(
            name="data_block", refers_to="=Sheet1!$A$1:$D$4", handle=h))
        check("excel_list_names", xl.excel_list_names(handle=h))
        check("excel_find", xl.excel_find(what="beta", handle=h, sheet="Sheet1"))
        check("excel_replace", xl.excel_replace(
            find_text="gamma", replace_text="delta", handle=h, sheet="Sheet1"))
        check("excel_cell_info", xl.excel_cell_info(cell="D2", handle=h,
                                                    sheet="Sheet1"),
              expect_keys=("formula",))
        check("excel_add_chart", xl.excel_add_chart(
            data_range="A1:B4", handle=h, sheet="Sheet1", chart_type="column",
            title="Qty"))
        check("excel_list_charts", xl.excel_list_charts(handle=h))
        check("excel_insert_delete", xl.excel_insert_delete(
            action="insert_rows", target="5:5", handle=h, sheet="Sheet1"))
        check("excel_clear_range", xl.excel_clear_range(
            range_a1="A5:D5", handle=h, sheet="Sheet1", what="contents"))
        path = os.path.join(TMP, "smoke.xlsx")
        check("excel_save(as)", xl.excel_save(handle=h, path=path))
        check("excel_export_pdf", xl.excel_export_pdf(
            output_path=os.path.join(TMP, "smoke.pdf"), handle=h, sheet="Sheet1"),
            expect_keys=("pdf",))
        check("com_describe(excel)", common.com_describe(
            app="excel", path="ActiveSheet"), expect_keys=("properties",))
        check("com_get(excel)", common.com_get(
            path="ActiveSheet.UsedRange.Address", app="excel"))
        hinted = common.com_get(path="ActiveSheet.UsedRange.Address()", app="excel")
        RESULTS.append(("com_get hints on parameterised property",
                        hinted.get("ok") is False and bool(hinted.get("hint")), ""))
        print("[%s] com_get hints on parameterised property"
              % ("PASS" if hinted.get("hint") else "FAIL"))
        check("com_eval(excel)", common.com_eval(
            expression="[ws.Name for ws in wb.Worksheets]", handle=h))
        check("com_call(excel)", common.com_call(
            path="Worksheets(1).Range('A1').Select", handle=h))
    finally:
        check("excel_close", xl.excel_close(handle=h, save=False))


def test_word() -> None:
    section("word")
    created = check("word_new", wd.word_new(visible=False), expect_keys=("handle",))
    h = created.get("handle")
    if not h:
        return
    try:
        check("word_append_text(h1)", wd.word_append_text(
            text="Smoke Test Report", handle=h, style="Heading 1"))
        check("word_append_text(body)", wd.word_append_text(
            text="This paragraph was written over COM.", handle=h))
        check("word_append_text(h2)", wd.word_append_text(
            text="Findings", handle=h, style="Heading 2"))
        check("word_insert_table", wd.word_insert_table(
            values=[["item", "value"], ["latency", "12ms"], ["rows", "3"]],
            handle=h))
        check("word_list_tables", wd.word_list_tables(handle=h))
        tbl = check("word_read_table", wd.word_read_table(table_index=1, handle=h),
                    expect_keys=("values",))
        assert_equal("word table round-trip", tbl.get("values"),
                     [["item", "value"], ["latency", "12ms"], ["rows", "3"]])
        check("word_get_outline", wd.word_get_outline(handle=h),
              expect_keys=("headings",))
        check("word_replace_text", wd.word_replace_text(
            find_text="over COM", replace_text="through the MCP server", handle=h))
        check("word_find", wd.word_find(query="MCP", handle=h))
        check("word_format_paragraphs", wd.word_format_paragraphs(
            start_paragraph=2, handle=h, italic=True))
        check("word_add_comment", wd.word_add_comment(
            text="reviewed by smoke test", handle=h, paragraph=2))
        check("word_list_comments", wd.word_list_comments(handle=h))
        check("word_revisions(enable)", wd.word_revisions(action="enable", handle=h))
        check("word_append_text(tracked)", wd.word_append_text(
            text="tracked insertion", handle=h))
        check("word_revisions(list)", wd.word_revisions(action="list", handle=h))
        check("word_revisions(accept)", wd.word_revisions(
            action="accept_all", handle=h))
        check("word_revisions(disable)", wd.word_revisions(
            action="disable", handle=h))
        check("word_insert_break", wd.word_insert_break(handle=h, kind="page"))
        check("word_document_info", wd.word_document_info(handle=h),
              expect_keys=("statistics",))
        check("word_list_styles", wd.word_list_styles(handle=h))
        path = os.path.join(TMP, "smoke.docx")
        check("word_save(as)", wd.word_save(handle=h, path=path))
        check("word_export_pdf", wd.word_export_pdf(
            output_path=os.path.join(TMP, "smoke_doc.pdf"), handle=h))
        check("office_convert(docx->pdf)", common.office_convert(
            input_path=path, output_path=os.path.join(TMP, "converted.pdf")),
            expect_keys=("output",))
    finally:
        check("word_close", wd.word_close(handle=h, save=False))


def test_ppt() -> None:
    section("powerpoint")
    created = check("ppt_new", pp.ppt_new(), expect_keys=("handle",))
    h = created.get("handle")
    if not h:
        return
    try:
        check("ppt_add_slide(title)", pp.ppt_add_slide(
            handle=h, layout="title_only", title="M365 MCP smoke test"))
        check("ppt_add_slide(text)", pp.ppt_add_slide(
            handle=h, layout="text", title="Agenda",
            body=["COM transport", "Tool families", "Escape hatch"]))
        check("ppt_add_textbox", pp.ppt_add_textbox(
            slide=1, text="generated by smoke test", handle=h, top=300,
            font_size=14))
        check("ppt_add_table", pp.ppt_add_table(
            slide=2, values=[["family", "tools"], ["excel", "30"], ["word", "26"]],
            handle=h, top=340, height=120))
        check("ppt_add_chart", pp.ppt_add_chart(
            slide=1, categories=["a", "b", "c"], series={"count": [3, 5, 2]},
            handle=h, chart_type="column", title="demo", left=420, top=60,
            width=420, height=260))
        check("ppt_set_notes", pp.ppt_set_notes(
            index=1, text="speaker notes from COM", handle=h))
        listed = check("ppt_list_slides", pp.ppt_list_slides(handle=h),
                       expect_keys=("slides",))
        assert_equal("ppt slide titles",
                     [s.get("title") for s in listed.get("slides", [])],
                     ["M365 MCP smoke test", "Agenda"])
        assert_equal("ppt speaker notes",
                     (listed.get("slides") or [{}])[0].get("notes"),
                     "speaker notes from COM")
        check("ppt_list_shapes", pp.ppt_list_shapes(slide=2, handle=h))
        check("ppt_duplicate_slide", pp.ppt_duplicate_slide(index=2, handle=h))
        check("ppt_move_slide", pp.ppt_move_slide(index=3, to_index=1, handle=h))
        check("ppt_replace_text", pp.ppt_replace_text(
            find_text="Agenda", replace_text="Plan", handle=h))
        check("ppt_delete_slide", pp.ppt_delete_slide(index=1, handle=h))
        path = os.path.join(TMP, "smoke.pptx")
        check("ppt_save(as)", pp.ppt_save(handle=h, path=path))
        check("ppt_export_pdf", pp.ppt_export_pdf(
            output_path=os.path.join(TMP, "smoke_deck.pdf"), handle=h))
        check("ppt_export_images", pp.ppt_export_images(
            output_dir=os.path.join(TMP, "slides"), handle=h, slide=1))
    finally:
        check("ppt_close", pp.ppt_close(handle=h, save=False))


def test_outlook() -> None:
    section("outlook (read-only + one deleted draft)")
    check("outlook_accounts", ol.outlook_accounts(), expect_keys=("accounts",))
    folders = check("outlook_list_folders", ol.outlook_list_folders(max_depth=2),
                    expect_keys=("folders",))
    print("    folders:", len(folders.get("folders", [])))
    check("outlook_folder_stats", ol.outlook_folder_stats(folder="inbox"))
    msgs = check("outlook_list_messages", ol.outlook_list_messages(
        folder="inbox", limit=3), expect_keys=("messages",))
    rows = msgs.get("messages", [])
    if rows:
        print("    newest:", str(rows[0].get("subject"))[:80])
        check("outlook_get_message", ol.outlook_get_message(
            entry_id=rows[0]["entry_id"], max_chars=500), expect_keys=("body",))
    check("outlook_search", ol.outlook_search(
        query="the", folder="inbox", limit=3, fields="subject"))
    check("outlook_calendar_list", ol.outlook_calendar_list(limit=5),
          expect_keys=("appointments",))
    check("outlook_contacts_list", ol.outlook_contacts_list(limit=5))
    check("outlook_tasks_list", ol.outlook_tasks_list(limit=5))
    check("outlook_categories", ol.outlook_categories(action="list"))
    draft = check("outlook_create_draft", ol.outlook_create_draft(
        to="", subject="[m365-mcp smoke test] draft", body="not sent"),
        expect_keys=("entry_id",))
    if draft.get("entry_id"):
        fetched = ol.outlook_get_message(entry_id=draft["entry_id"], max_chars=200)
        assert_equal("outlook draft round-trip subject", fetched.get("subject"),
                     "[m365-mcp smoke test] draft")
        unsent = ol.outlook_send_draft(entry_id=draft["entry_id"])
        RESULTS.append(("outlook_send_draft refuses without confirm",
                        unsent.get("ok") is False, ""))
        print("[%s] outlook_send_draft refuses without confirm"
              % ("PASS" if unsent.get("ok") is False else "FAIL"))
        check("outlook_delete_draft", ol.outlook_delete_draft(
            entry_id=draft["entry_id"]))


def test_onenote() -> None:
    section("onenote")
    tree = check("onenote_hierarchy", one.onenote_hierarchy(scope="sections"),
                 expect_keys=("nodes",))
    nodes = tree.get("nodes", [])
    print("    nodes:", len(nodes))
    check("onenote_find_pages", one.onenote_find_pages(query="meeting"))
    pages = [n for n in nodes if n["type"] == "page"]
    if pages:
        check("onenote_get_page", one.onenote_get_page(
            page_id=pages[0]["id"], max_chars=300))


def test_research() -> None:
    section("research primitives")

    # --- notes in, from each supported format ---
    md_path = os.path.join(TMP, "notes.md")
    with open(md_path, "w", encoding="utf-8") as fh:
        fh.write("# 3Q25 call\n\n- FY26 gross margin guided to 25.5%, from 26%\n"
                 "- Shipment volume 59mn pairs\n")
    notes = check("notes_read(md)", rs.notes_read(path=md_path),
                  expect_keys=("numbered_text",))
    assert_equal("notes_read numbers the lines", notes.get("lines"), 4)
    RESULTS.append(("notes_read keeps the citable line prefix",
                    "   3| - FY26 gross margin" in notes.get("numbered_text", ""), ""))
    print("[%s] notes_read keeps the citable line prefix"
          % ("PASS" if "   3| - FY26" in notes.get("numbered_text", "") else "FAIL"))

    for name, path in (("pdf", os.path.join(TMP, "smoke.pdf")),
                       ("docx", os.path.join(TMP, "smoke.docx"))):
        if os.path.exists(path):
            got = check("notes_read(%s)" % name, rs.notes_read(path=path,
                                                               max_chars=2000))
            assert_equal("notes_read(%s) routes correctly" % name, got.get("source"),
                         "pdf" if name == "pdf" else "word")
    if os.path.exists(os.path.join(TMP, "smoke.pdf")):
        check("pdf_read_text", rs.pdf_read_text(
            path=os.path.join(TMP, "smoke.pdf"), pages="1"),
            expect_keys=("total_pages", "text"))

    # --- a miniature model with the sell-side colour convention ---
    created = check("excel_new(model)", xl.excel_new(visible=False),
                    expect_keys=("handle",))
    h = created.get("handle")
    if not h:
        return
    model_path = os.path.join(TMP, "mini_model.xlsx")
    try:
        xl.excel_rename_sheet(sheet=xl.excel_list_sheets(handle=h)["sheets"][0]["name"],
                              new_name="Model", handle=h)
        xl.excel_write_range(handle=h, sheet="Model", start_cell="A1",
                             values=[["Item", 2024, 2025, "2026E"]])
        xl.excel_write_range(handle=h, sheet="Model", start_cell="A2",
                             values=[["Revenue", 1000, 1200, None],
                                     ["Gross margin", 0.25, 0.26, None],
                                     ["Gross profit", None, None, None]])
        # green = reported actuals, blue = the analyst's assumptions
        xl.excel_format_range(range_a1="B2:C3", handle=h, sheet="Model",
                              font_color="#008000")
        xl.excel_write_range(handle=h, sheet="Model", start_cell="D2",
                             values=[[1400], [0.255]])
        xl.excel_format_range(range_a1="D2:D3", handle=h, sheet="Model",
                              font_color="#0000FF")
        xl.excel_write_range(handle=h, sheet="Model", start_cell="B4",
                             values=[["=B2*B3", "=C2*C3", "=D2*D3"]], as_formula=True)
        xl.excel_calculate(handle=h)
        xl.excel_save(handle=h, path=model_path)

        mapped = check("excel_model_map", rs.excel_model_map(handle=h),
                       expect_keys=("sheets", "colour_legend"))
        sheet = (mapped.get("sheets") or [{}])[0]
        assert_equal("model_map finds the period header row",
                     sheet.get("header_row"), 1)
        assert_equal("model_map reads the period axis",
                     sheet.get("periods"), ["2024", "2025", "2026E"])
        by_cell = {c["cell"]: c for c in sheet.get("cells", [])}
        assert_equal("model_map classes blue as an assumption",
                     by_cell.get("D2", {}).get("class"), "assumption")
        assert_equal("model_map classes green as a reported actual",
                     by_cell.get("B2", {}).get("class"), "actual")
        assert_equal("model_map gives a cell its row label",
                     by_cell.get("D2", {}).get("label"), "Revenue")
        assert_equal("model_map gives a cell its period",
                     by_cell.get("D2", {}).get("period"), "2026E")
        assert_equal("model_map marks assumptions writable",
                     by_cell.get("D2", {}).get("writable"), True)
        assert_equal("model_map does not mark actuals writable",
                     by_cell.get("B2", {}).get("writable"), False)

        traced = check("excel_trace_precedents", rs.excel_trace_precedents(
            cell="D4", handle=h, sheet="Model"), expect_keys=("tree",))
        leaves = {n["cell"].split("!")[-1] for n in traced.get("input_leaves", [])}
        RESULTS.append(("trace_precedents reaches the assumption cells",
                        {"D2", "D3"} <= leaves, "got %s" % sorted(leaves)))
        print("[%s] trace_precedents reaches the assumption cells"
              % ("PASS" if {"D2", "D3"} <= leaves else "FAIL"))

        snap = check("excel_snapshot", rs.excel_snapshot(handle=h, label="test"),
                     expect_keys=("snapshot",))
        RESULTS.append(("snapshot file exists", bool(snap.get("exists")), ""))
        print("[%s] snapshot file exists" % ("PASS" if snap.get("exists") else "FAIL"))

        # --- propose: the verdicts are the whole point ---
        cs_path = os.path.join(TMP, "mini.changeset.json")
        proposed = check("excel_propose_changes", rs.excel_propose_changes(
            handle=h, changeset_path=cs_path, note_source="unit test",
            changes=[
                {"sheet": "Model", "cell": "D2", "new_value": 1500,
                 "reason": "guidance cut", "source": "notes line 3"},
                {"sheet": "Model", "cell": "D4", "new_value": 999,
                 "reason": "should be refused", "source": "notes line 4"},
                {"sheet": "Model", "cell": "B2", "new_value": 1001,
                 "reason": "restating history", "source": "notes line 5"},
                {"sheet": "Model", "cell": "D3", "new_value": 0.255,
                 "reason": "unchanged", "source": "notes line 6"},
            ]), expect_keys=("counts", "changes"))
        by_cell = {c["cell"]: c for c in proposed.get("changes", [])}
        assert_equal("propose accepts a blue assumption cell",
                     by_cell.get("D2", {}).get("verdict"), "ok")
        assert_equal("propose BLOCKS a formula cell",
                     by_cell.get("D4", {}).get("verdict"), "blocked")
        assert_equal("propose warns on a non-assumption colour",
                     by_cell.get("B2", {}).get("verdict"), "warn")
        assert_equal("propose warns when the value is unchanged",
                     by_cell.get("D3", {}).get("verdict"), "warn")
        RESULTS.append(("propose records the formula it refused to overwrite",
                        "=D2*D3" in str(by_cell.get("D4", {}).get("current_formula")),
                        ""))
        print("[%s] propose records the formula it refused to overwrite"
              % ("PASS" if "=D2*D3" in str(by_cell.get("D4", {}).get("current_formula"))
                 else "FAIL"))
        RESULTS.append(("changeset file written", os.path.exists(cs_path), ""))
        print("[%s] changeset file written" % ("PASS" if os.path.exists(cs_path)
                                               else "FAIL"))

        # --- apply: refuses without confirm, then writes + annotates ---
        dry = rs.excel_apply_changeset(changeset_path=cs_path, handle=h)
        assert_equal("apply refuses without confirm", dry.get("ok"), False)
        assert_equal("apply counts blocked rows", dry.get("blocked"), 1)

        before = xl.excel_read_range(handle=h, sheet="Model",
                                     range_a1="D2").get("values")
        applied = check("excel_apply_changeset(confirm)", rs.excel_apply_changeset(
            changeset_path=cs_path, handle=h, confirm=True, snapshot=True),
            expect_keys=("written", "cells"))
        after = xl.excel_read_range(handle=h, sheet="Model",
                                    range_a1="D2").get("values")
        assert_equal("apply wrote the assumption cell", after, [[1500]])
        RESULTS.append(("apply changed the value", before != after, ""))
        print("[%s] apply changed the value" % ("PASS" if before != after else "FAIL"))
        formula_after = xl.excel_read_range(handle=h, sheet="Model", range_a1="D4",
                                            mode="formulas").get("values")
        assert_equal("apply left the blocked formula intact",
                     formula_after, [["=D2*D3"]])
        info = xl.excel_cell_info(cell="D2", handle=h, sheet="Model")
        RESULTS.append(("apply left an audit comment",
                        "guidance cut" in str(info.get("comment", "")), ""))
        print("[%s] apply left an audit comment"
              % ("PASS" if "guidance cut" in str(info.get("comment", "")) else "FAIL"))

        # --- data sources: clean model needs no add-in ---
        clean = check("excel_data_sources(clean)", rs.excel_data_sources(handle=h),
                      expect_keys=("verdict",))
        assert_equal("clean workbook is safe to edit", clean.get("verdict"),
                     "safe_to_edit")

        # --- model_check must find a broken subtotal on its own ---
        xl.excel_add_sheet(name="Seg", handle=h)
        xl.excel_write_range(handle=h, sheet="Seg", start_cell="A1", values=[
            ["Item", "1H23", "2H23", "1H24", "2H24", "1H25"],
            ["Total", 100.0, 110.0, 120.0, 130.0, 140.0],
            ["Alpha", 60.0, 66.0, 72.0, 78.0, 84.0],
            ["Beta", 25.0, 27.5, 30.0, 32.5, 35.0],
            # last column is 10x what it should be (2.1 -> 21.0): the total row
            # and the other five columns all disagree with it
            ["Gamma", 15.0, 16.5, 18.0, 19.5, 210.0],
        ])
        checked = check("excel_model_check", rs.excel_model_check(
            handle=h, sheets="Seg"), expect_keys=("findings",))
        subs = [f for f in checked.get("findings", []) if f["check"] == "subtotal"]
        RESULTS.append(("model_check finds the broken subtotal column",
                        any(f.get("period") == "1H2025" for f in subs),
                        "found %s" % [f.get("period") for f in subs]))
        print("[%s] model_check finds the broken subtotal column"
              % ("PASS" if any(f.get("period") == "1H2025" for f in subs) else "FAIL"))

        # --- add-in dependency is detected and blocks writes ---
        xl.excel_write_range(handle=h, sheet="Seg", start_cell="H1",
                             values=[['=WSD("600519.SH","close")']],
                             as_formula=True)
        dep = check("excel_data_sources(wind)", rs.excel_data_sources(handle=h))
        assert_equal("wind dependency detected", dep.get("addins_required"), ["wind"])
        assert_equal("verdict turns read-only", dep.get("verdict"),
                     "read_only_recommended")
        guarded = rs.excel_apply_changeset(changeset_path=cs_path, handle=h,
                                           confirm=True)
        assert_equal("apply refuses while an add-in is missing",
                     guarded.get("ok"), False)
        RESULTS.append(("refusal names the missing add-in",
                        "wind" in str(guarded.get("error", "")), ""))
        print("[%s] refusal names the missing add-in"
              % ("PASS" if "wind" in str(guarded.get("error", "")) else "FAIL"))
    finally:
        check("excel_close(model)", xl.excel_close(handle=h, save=False))


def test_access() -> None:
    section("access / publisher (availability only)")
    result = misc.access_list_objects(db_path=os.path.join(TMP, "nonexistent.accdb"))
    RESULTS.append(("access_list_objects reports missing file",
                    result.get("ok") is False, ""))
    print("[%s] access_list_objects reports missing file"
          % ("PASS" if result.get("ok") is False else "FAIL"))
    guard = misc.access_execute(db_path="x.accdb", sql="DELETE FROM t")
    RESULTS.append(("access_execute refuses without confirm",
                    guard.get("ok") is False, ""))
    print("[%s] access_execute refuses without confirm"
          % ("PASS" if guard.get("ok") is False else "FAIL"))


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--skip", default="", help="comma list of sections to skip")
    args = parser.parse_args()
    skip = {s.strip().lower() for s in args.skip.split(",") if s.strip()}

    sections = [
        ("common", test_common),
        ("excel", test_excel),
        ("word", test_word),
        ("ppt", test_ppt),
        ("outlook", test_outlook),
        ("onenote", test_onenote),
        ("research", test_research),
        ("access", test_access),
    ]
    for name, fn in sections:
        if name in skip:
            print("\n=== %s (skipped) ===" % name)
            continue
        try:
            fn()
        except Exception:  # noqa: BLE001
            RESULTS.append((name + " (crashed)", False, traceback.format_exc(limit=3)))
            print("[FAIL] %s crashed:\n%s" % (name, traceback.format_exc(limit=3)))

    # Leave the desktop as we found it: quit only the apps we started ourselves.
    section("cleanup")
    for app in ("excel", "word", "powerpoint"):
        if app in PRE_RUNNING:
            print("    %s was already running - left alone" % app)
            continue
        running = {row["app"] for row in common.m365_status().get("applications", [])
                   if row.get("running")}
        if app in running:
            check("m365_quit_app(%s)" % app,
                  common.m365_quit_app(app=app, confirm=True))

    passed = sum(1 for _, ok, _ in RESULTS if ok)
    print("\n%d/%d checks passed. Artifacts in %s" % (passed, len(RESULTS), TMP))
    for label, ok, detail in RESULTS:
        if not ok:
            print("  FAILED: %s %s" % (label, detail[:300]))
    return 0 if passed == len(RESULTS) else 1


if __name__ == "__main__":
    sys.exit(main())
