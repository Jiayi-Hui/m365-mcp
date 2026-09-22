# m365-mcp

**Lets an agent work inside the Excel, Word, PowerPoint and Outlook that are
already on your machine.**

The usual way to get AI near your own files is to export them, upload them, or
file a ticket for Graph API permissions — and what comes back is a copy, in a
chat window, disconnected from the workbook you actually have to hand in. This
server removes that detour. It drives the desktop applications through the same
object model a VBA macro uses, so the agent edits *your* file, reads *your*
already-signed-in mailbox, and saves the PDF where you asked. No password, no
Azure app registration, no network call.

141 tools: Excel, Word, PowerPoint, Outlook, OneNote, Access, Publisher, a
generic layer that reaches the rest of the object model when no dedicated tool
exists, and a small set of research primitives for working with financial
models and analyst notes.

---

## The mental model

Five things to internalise; the rest follows.

**1. You talk to an agent; the agent picks the tool.** You never type tool
names. "把那个季度模型里的收入读出来" is a complete request — the agent works
out that it needs `excel_open` then `excel_read_range`. The tool list exists for
the agent, not for you.

**2. It edits the original, not a copy.** This is the whole point, and also the
thing to respect: when you ask it to write a column, the column appears in your
file. There is no automatic backup and no undo stack the agent can reach. Read
operations are free; say "另存一份再改" when you want the original preserved.

**3. The trust boundary is this machine.** No credential ever enters this
server. It can read your mail because *your* Outlook is signed in on *your*
desktop, and it can reach exactly as much as you can, this session. Nothing
leaves the machine unless a tool you asked for sends it.

**4. Four ways to point at a file.** A handle the agent got when it opened the
file (`xlwb:1`), the file name (`Q3_model.xlsx`), the full path, or nothing at
all — which means "whatever is in front of me right now". Already have the
workbook open? Just name it; the agent attaches to that window instead of
opening a second copy.

**5. Sending mail is always two steps.** Everything that composes — new mail,
reply, forward — only ever writes a draft and hands back its id. A separate tool
sends, and it refuses unless you have said yes. Same shape for deleting mail,
writing to an Access database, and quitting an app.

---

## How to ask for things

Say it normally. These are patterns that work, and what the agent runs:

| What you want | Say something like | Runs |
| --- | --- | --- |
| See what's open right now | 「现在 Excel 和 Outlook 里都开着什么」 | `m365_status` |
| Read a sheet | 「把 Q3_model.xlsx 的 Summary 页读出来」 | `excel_open` → `excel_read_range` |
| Write results back | 「把算好的环比写到 F 列，加个千分位格式」 | `excel_write_range` |
| Append to a log sheet | 「把这三行追加到台账最下面」 | `excel_append_rows` |
| Understand one cell | 「D7 这个数是怎么算出来的」 | `excel_cell_info`（值+公式+格式+引用单元格） |
| Chart it | 「按月份画条折线图放右边」 | `excel_add_chart` |
| Export | 「这个表导成 PDF 发我桌面」 | `excel_export_pdf` |
| Draft a document | 「按这个大纲写份 Word 报告，一级标题用 Heading 1」 | `word_new` → `word_append_text` |
| Fill a template | 「把模板里的 {{客户名}} 全换成宁德时代」 | `word_replace_text` |
| Review someone's draft | 「这份稿子里有哪些批注和修订，谁提的」 | `word_list_comments` / `word_revisions` |
| Read a deck | 「这份 PPT 每页讲什么，把备注也带上」 | `ppt_list_slides` |
| Build slides | 「按这几个要点做三页，标题页 + 两页正文」 | `ppt_add_slide` |
| Re-skin a deck | 「把全篇的 2025 改成 2026」 | `ppt_replace_text` |
| Slides as images | 「把第 3 页导成 PNG」 | `ppt_export_images` |
| Find mail | 「上周老板发过来的邮件有哪些」 | `outlook_list_messages` |
| Search bodies | 「收件箱里提到 pricing pressure 的邮件」 | `outlook_search` |
| Read one in full | 「把那封读全文，附件也列一下」 | `outlook_get_message` |
| Save attachments | 「这封的附件都存到 D:\\downloads」 | `outlook_save_attachments` |
| Write a reply | 「帮我起个草稿回他，说这周五之前给数」 | `outlook_reply_draft`（只存草稿） |
| Actually send it | 「可以，发出去」 | `outlook_send_draft --confirm` |
| Tidy the inbox | 「把这些归到 Research 文件夹，标上已读」 | `outlook_move_message` / `outlook_update_message` |
| Check the calendar | 「我明天有什么会」 | `outlook_calendar_list` |
| Look someone up | 「这个人的邮箱和部门是什么」 | `outlook_resolve_recipient`（查 GAL） |
| Search OneNote | 「OneNote 里搜一下上次那个定价讨论」 | `onenote_find_pages` |
| Query an Access db | 「这个 accdb 里 orders 表今年有多少条」 | `access_query`（走 ADO，不开 Access 界面） |
| Convert anything | 「这几个 docx 都转成 PDF」 | `office_convert` |
| Something with no tool | 「给这张表加个数据透视」 | `com_describe` → `com_eval` |
| Read notes of any format | 「把这份纪要读进来」 | `notes_read`（md/docx/pdf 统一入口，带行号） |
| Understand a model's shape | 「这个模型里哪些格子是假设、哪些是公式」 | `excel_model_map` |
| Trace a number | 「这个 EPS 是由哪几个假设推出来的」 | `excel_trace_precedents` |
| Back up before editing | 「改之前先存个快照」 | `excel_snapshot` |
| Propose model edits from notes | 「按这份纪要看看模型哪些假设该动」 | `excel_propose_changes`（只校验+落盘，不写） |
| Apply them once approved | 「可以，改吧」 | `excel_apply_changeset --confirm` |
| Check a model before trusting it | 「这个模型有没有问题」 | `excel_model_check` |
| See what live data it depends on | 「这模型依赖 Wind 吗，现在能用吗」 | `excel_data_sources` |

### What makes a good request here

- **Name the file the way you see it.** If the workbook is open on your screen,
  its name is enough (`「预算表.xlsx 里」`). If it isn't, give the full path —
  the agent can't guess which `report.xlsx` you mean.
- **Say "另存" when you mean it.** Write tools modify the file in place. "把结果
  写回去" edits the original; "另存成 v2 再写" gets you a copy first. Neither is
  the default the agent should guess.
- **Re-reading is free, re-writing is not.** Asking for the same range twice
  costs nothing. Asking it to "再写一遍" appends or overwrites for real.
- **Mail is drafted, then sent, and you decide in between.** 「起个草稿」 and
  「发出去」 are two separate turns on purpose. If you say 「回封邮件给他」 you
  will get a draft and a preview, not a sent message.
- **Big ranges come back truncated.** Reads cap at 5,000 cells so a 200k-row
  sheet doesn't drown the conversation; the reply says `truncated: true` and how
  many rows exist. Ask for a specific range, or say 「全量读出来」 and the agent
  raises the cap.
- **Ask for the analysis, not the extraction.** The requests that earn their
  keep chain the file operation into what comes next:
  - 「把这个模型里近 8 个季度的收入和毛利读出来，算出环比，然后写回一个新 sheet 并配张折线图」
  - 「收件箱里找出这个月所有券商的业绩前瞻，按公司归类，做成一份 Word 摘要」
  - 「把这份 PPT 每页的要点提出来，跟上个季度那版对比，告诉我哪些口径变了」
- **If no tool fits, ask anyway.** Pivot tables, mail merge, conditional
  formatting, slide masters — none have dedicated tools, and the agent can still
  do them: `com_describe` shows it what the live object supports, `com_eval`
  drives it. Tell it what you want, not how.

---

## What you get

Structured JSON, not prose the agent has to re-parse. A range read:

```json
{
  "workbook": "Book1", "sheet": "Sheet1", "address": "A1:C3",
  "rows": 3, "columns": 3, "truncated": false, "total_rows": 4,
  "values": [["name", "qty", "price"],
             ["alpha", 3.0, 1.5],
             ["beta", 7.0, 2.25]],
  "ok": true
}
```

Read the used range instead and the computed column comes along, formulas
already evaluated:

```json
"values": [["name", "qty", "price", null],
           ["alpha", 3.0, 1.5, 4.5],
           ["beta", 7.0, 2.25, 15.75]]
```

Failures are data too, with a hint the agent can act on rather than a stack
trace to guess at:

```json
{
  "ok": false,
  "error": "Unknown handle 'xlwb:999'",
  "hint": "Handles live for one server run; re-open the file for a new one."
}
```

And a send that you have not approved returns the preview instead of sending:

```json
{
  "ok": false,
  "error": "confirm=False - nothing was sent",
  "preview": {"to": "someone@example.com", "cc": null,
              "subject": "Re: Q3 numbers",
              "body_head": "Hi — we'll have the figures by Friday..."},
  "hint": "Show this preview to the user; call again with confirm=true once they approve."
}
```

---

## Things that will confuse you if nobody says them

**Windows pop open. That is not a bug.** Excel and Word can be driven hidden,
but PowerPoint flatly refuses `Visible=False` — automating a deck always shows
one. And when the agent attaches to an app you already had open, it works in
*your* window, so you will see cells fill in and slides appear. That visibility
is the point: you can watch what it did and press Ctrl+Z yourself.

**If Office is showing a dialog, everything stops.** A "Save changes?" box, an
open cell in edit mode, a file-in-use prompt — while any of those are up, COM
calls get rejected (`RPC_E_CALL_REJECTED`) or time out. The fix is to look at
the screen and clear it. The server suppresses the alerts it can (`DisplayAlerts
= False`), but it cannot dismiss one you triggered by hand.

**The new Outlook has no COM at all.** Only classic Outlook for Windows works.
If mail tools fail with "Application is not running" while Outlook is plainly
open, you are on the new one — `m365_status` tells you which.

**OneNote needs one registry key on 64-bit Python.** Classic OneNote 2016
registers its type library only under the `win32` key, so a 64-bit caller gets
`TYPE_E_LIBNOTREGISTERED` even though OneNote is 64-bit and starts fine. Run
this once, as yourself — it writes a single key under `HKCU`, needs no admin,
and has a `-Revert` switch:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\tools\fix_onenote_typelib.ps1 -WhatIf
powershell -NoProfile -ExecutionPolicy Bypass -File .\tools\fix_onenote_typelib.ps1
```

**Numbers come back as floats.** Excel stores every number as a double, so `3`
reads back as `3.0`. That is Excel, faithfully reported, not a conversion bug.

**Handles die when the server restarts.** `xlwb:1` is valid for one run of the
server. After a restart the agent just re-opens the file — but if you reference
"that workbook from yesterday", it no longer means anything.

**A closed Excel can linger as an invisible process.** Excel started by
automation sometimes survives `Quit()` with no window. Harmless, and it holds no
documents — check `excel_list_workbooks` before killing it if you want to be
sure.

---

## Setup

Windows, desktop Office (classic Outlook, not the new one), and:

```bash
pip install -r requirements.txt
```

Register it with Claude Code:

```bash
claude mcp add m365 -e PYTHONPATH=<path-to-this-repo> -- <path-to-python.exe> -m m365_mcp.server
```

Or add the equivalent to `~/.claude.json` / `claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "m365": {
      "type": "stdio",
      "command": "C:\\path\\to\\python.exe",
      "args": ["-m", "m365_mcp.server"],
      "env": {"PYTHONPATH": "C:\\path\\to\\m365-mcp"}
    }
  }
}
```

Check it works — this drives the real applications and cleans up after itself:

```bash
python tests/smoke_test.py
```

103 checks; 101 pass on a stock Office 16 install, the two OneNote ones after
the registry fix above.

---

## Scope

It does what you ask inside the apps, and nothing on its own initiative. It
never sends mail, deletes a message, writes to a database or quits an
application without an explicit confirmation in the conversation. It has no
network path of its own: files go where you name, mail goes to whom you name.

One thing to be deliberate about: `com_eval` runs Python against the live Office
objects, which makes it exactly as powerful as the VBA editor — that is what
lets it cover the long tail. Keep this server on the single workstation it was
started for; don't expose it to anyone you wouldn't hand a macro editor.

---

## For maintainers

- [`docs/TOOLS.md`](docs/TOOLS.md) — generated index of all 132 tools
  (`python tools/dump_tool_index.py` after adding one).
- [`m365_mcp/comcore.py`](m365_mcp/comcore.py) — the invariant everything rests
  on: every COM call is marshalled onto one long-lived STA thread, because COM
  objects are thread-affine and FastMCP runs sync tools on arbitrary workers.
  Never touch a COM object from anywhere else.
- [`tests/smoke_test.py`](tests/smoke_test.py) — runs against the real apps;
  mocks would test nothing, since every bug here lives in the real object
  model's behaviour.

One trap worth repeating, because it produced a **silent** wrong answer: under
late binding, parameterised properties are evaluated *before* your parentheses.
`Range.Address(False, False)` at least raises `TypeError`. But
`ws.Range("A1").Resize(2, 3)` quietly becomes `Range("A1").Resize` — the same
range — followed by `Item(2, 3)`, i.e. the **single cell C2**. Data lands in the
wrong place, no error, and the returned `cells_written` still looks right.
`_resize()` builds blocks from their corners to avoid this; `Offset` has the
same shape. The invariant to re-verify after touching range code: write a known
grid, read it back, compare cell by cell — `smoke_test.py` does exactly that,
and it is the only reason the bug was found.
