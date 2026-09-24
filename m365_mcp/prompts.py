"""MCP prompts: the packaged workflows.

MCP servers can expose named workflow templates, not just tools. In Claude Code
these appear as slash commands (`/mcp__m365__update_model_from_notes`), which is
what lets a multi-step procedure ship *with* the server instead of living in a
separate skill directory that has to be synced.

The reasoning stays in the model. These only carry the procedure and the rules
it must not break - the same text works pasted into any agent if a harness turns
out not to support prompts.

Deliberately generic: no house-specific report structure lives here, because a
report outline belongs to the desk that wrote it, not to a server that gets
installed on other people's machines. Pass one in as a template path instead.
"""

from __future__ import annotations

# The rules every model-touching workflow shares. Stated once, referenced below.
_GROUND_RULES = """\
## Rules that override anything else in this procedure

1. **Every annotation must cite its source** - the line number from the notes,
   or the page of the filing. If you cannot point at where a claim came from,
   do not write it. No exceptions, and no inferring a number "from context".
2. **Never change a cell's value.** Use `excel_annotate` only. The analyst
   makes the edit; you make it possible to make it well.
3. **Never save the workbook.** The file on disk must stay byte-identical. If
   the model depends on a data add-in that is not loaded, saving would replace
   thousands of live cells with `#NAME?` permanently.
4. **Do not decide what an assumption should become.** Reconciliation is yours
   - "1H actual is 41% of the full-year estimate, against 50.7% last year" is a
   fact. "So cut FY26E revenue to 22,000" is the analyst's variant view. Give
   the arithmetic, stop before the conclusion.
5. **A formula cell is never the target.** If a figure lands on a formula, run
   `excel_trace_precedents` and annotate the input cells that actually drive it,
   naming them in the note.
"""

_SEVERITY = """\
## Severity

- `high` - an error in the model itself (wrong value, broken formula chain).
  These come first: they contaminate every conclusion drawn from the model.
- `medium` - new information that is materially at odds with a current
  assumption, and needs the analyst's judgement.
- `low` - worth knowing, not urgent. Definitional notes, mild divergences.
"""


def register(mcp) -> None:
    """Attach the prompts to a FastMCP instance."""

    @mcp.prompt()
    def update_model_from_notes(model_path: str, notes_path: str) -> str:
        """Locate what a set of notes means for a financial model, cell by cell,
        and annotate those cells in place. Writes no values and saves nothing."""
        return f"""\
Work out what the notes at `{notes_path}` mean for the model at `{model_path}`,
locate each implication on the specific cell it affects, and leave your findings
on those cells for the analyst to act on in Excel.

{_GROUND_RULES}

## Procedure

1. `notes_read(path="{notes_path}")` - the text comes back with line numbers.
   Those numbers are what you cite.
2. `excel_data_sources(path="{model_path}")` - find out which live data add-ins
   the model needs and whether they are loaded. Annotating is safe either way
   because you will not save; note the verdict in your final report.
3. `excel_model_map(handle=...)` - the structure. Every constant arrives with
   its row label, its period, its colour class (`assumption` / `actual` /
   `override`) and whether it is writable. This is the coordinate system you
   match the notes against.
4. `excel_model_check(handle=...)` - the model's own health. **Annotate what it
   finds before anything else**: a 10x typo or an off-by-one formula chain
   poisons any reconciliation you are about to do.
5. Go through the notes line by line. For every statement carrying a number or a
   guidance change:
   - find the row label x period it belongs to;
   - if that cell holds a formula, `excel_trace_precedents` to the real inputs;
   - if there is no cell for it at all, record it as a **model gap** - often the
     most valuable finding, because it means the model has no place for
     something that now exists (a period column that was never added, a segment
     that is not broken out).
6. `excel_annotate(handle=..., annotations=[...])` - one call, all findings.
7. Report back: how many annotations, split into model errors / assumptions
   needing judgement / gaps; plus anything in the notes you could not place.

## What goes in an annotation

- One sentence on what this cell's situation is.
- The evidence: the quoted line or page, and the arithmetic you did.
- For a formula cell: name the input cells that actually drive it.
- For a reconciliation: both numbers and the comparable prior-period ratio, so
  the analyst can see the gap rather than take your word for it.

{_SEVERITY}

## Finally

Tell the analyst the workbook is open with the annotations in it and the file on
disk is untouched, and that `excel_clear_annotations` removes them again.
"""

    @mcp.prompt()
    def check_model(model_path: str) -> str:
        """Health-check a model before trusting it with new data, and annotate
        whatever is wrong on the cells themselves."""
        return f"""\
Health-check the model at `{model_path}` and mark what you find on the cells.

{_GROUND_RULES}

## Procedure

1. `excel_data_sources(path="{model_path}")` - which add-ins it depends on,
   which are loaded, and which error cells that explains. Errors the add-ins do
   NOT explain are real breaks and matter more.
2. `excel_model_check(handle=...)` - four checks:
   - error cells, split by cause;
   - period arithmetic (FY vs 1H+2H vs the four quarters);
   - **subtotal rows against their components** - the strongest of the four,
     because it uses a redundancy the model already contains;
   - magnitude outliers against each row's own median.
   Also reports whether iterative calculation is on, which means a circular
   reference someone decided to tolerate.
3. For anything that looks real, confirm it before annotating: read the actual
   formulas with `excel_read_range(mode="formulas")`, and check whether the
   same relationship holds in the neighbouring periods. A finding you have not
   confirmed gets `low` severity and says so.
4. `excel_annotate(...)` - put each confirmed finding on its cell, with the
   arithmetic that demonstrates it.

## What makes a finding worth reporting

A structural contradiction beats a statistical oddity. "These three components
sum to X but the total row says Y, and the same relationship holds in 19 other
columns" is conclusive. "This value is 12x the row median" is a question, not an
answer - volatile rows are ordinary. Rank accordingly, and say which kind each
finding is.

{_SEVERITY}
"""

    @mcp.prompt()
    def model_change_brief(model_path: str, notes_path: str,
                           template_path: str = "") -> str:
        """Produce a written change brief from notes + model, as a document
        rather than cell annotations - for circulating or filing."""
        template_line = (
            f"Follow the outline at `{template_path}`."
            if template_path else
            "No outline was given, so use the three-part structure below."
        )
        return f"""\
Produce a written brief on what the notes at `{notes_path}` imply for the model
at `{model_path}`. {template_line}

This is the document form of the same work `update_model_from_notes` does on the
cells. Use it when the output has to be circulated or filed; use the annotation
workflow when the analyst will be working in Excel.

{_GROUND_RULES}

(Rules 2 and 3 still hold: this reads the model, it does not modify it.)

## Procedure

Same investigation as `update_model_from_notes` - `notes_read`,
`excel_data_sources`, `excel_model_map`, `excel_model_check`,
`excel_trace_precedents` where a figure lands on a formula - then write the
findings to a markdown file instead of annotating cells.

## Structure, when no outline is supplied

**A. Errors to fix first.** Anything wrong with the model itself, with the
arithmetic that proves it. These come first because they contaminate section B.

**B. Reconciliation.** Per affected assumption: the cell, its row label and
period, the current value, the fact from the notes, and the comparison that
makes the gap legible (usually the same ratio in the prior period). **No
recommended values.**

**C. Gaps.** What the notes contain that the model has nowhere to put.

Every row in every table cites its source line or page.

## After writing it

Offer to render it to .docx with `word_apply_markdown` if one exists, otherwise
`office_convert`. Do not do it unasked - the markdown is the reviewable artefact.
"""
