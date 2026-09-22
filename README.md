# m365-mcp — Microsoft 365 desktop automation as MCP tools

An MCP server that drives the Office applications already installed and signed
in on this machine, through Windows COM. No passwords, no Azure app
registration, no Graph consent, no network: it talks to the same object model a
VBA macro would.

**132 tools** across Excel, Word, PowerPoint, Outlook, OneNote, Access and
Publisher, plus a generic COM layer for everything the typed tools do not cover.

## Why COM rather than Graph

| | COM (this server) | Microsoft Graph |
|---|---|---|
| Setup | none — uses the signed-in desktop apps | Azure app registration + admin consent |
| Reach | anything the app can do, including local files, macros, charts, PDF export, tracked changes | what the API exposes |
| Scope | this machine, this user, this session | tenant-wide, delegated or app-only |
| Requires | classic desktop Office running | nothing local |

COM wins for "do this to the file/mailbox on my desk". Graph wins for
unattended server-side work. This server is the first kind.

## Install

Requires Windows, desktop Office (classic Outlook, not "new Outlook"), and:

```bash
pip install mcp==1.28.1 pywin32
```

Register with Claude Code:

```bash
claude mcp add m365 -e PYTHONPATH=<path-to-this-repo> -- <path-to-python.exe> -m m365_mcp.server
```

or add to `~/.claude.json` / `claude_desktop_config.json`:

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

Verify: `python tests/smoke_test.py` — it exercises every family against the
real apps and cleans up after itself (Outlook coverage is read-only plus one
draft that gets deleted).

## Architecture

```
m365_mcp/
  comcore.py       single STA thread, handle registry, error translation
  tools_common.py  status, conversion, com_describe / com_get / com_set /
                   com_call / com_eval / com_dispatch
  tools_excel.py   30 tools
  tools_word.py    26 tools
  tools_ppt.py     25 tools
  tools_outlook.py 23 tools
  tools_onenote.py  8 tools
  tools_misc.py    Access (5) + Publisher (4)
  server.py        FastMCP registration + m365_help
```

Three design decisions worth knowing:

**One STA thread owns every COM object.** COM apartments are thread-affine and
FastMCP runs sync tools on arbitrary worker threads, so `comcore` marshals every
call onto one long-lived `CoInitializeEx(COINIT_APARTMENTTHREADED)` thread. That
thread also owns the handle registry, which is why a workbook handle stays valid
across tool calls.

**Handles, not paths, for open files.** `excel_open` returns `xlwb:1`;
`word_open` returns `wddoc:2`. Every tool also accepts a file name or full path,
and falls back to the active document when you pass nothing. Handles live for
one server run.

**A typed surface plus an escape hatch.** The Office object models are far
larger than any fixed tool list. `com_describe` introspects a live object's
methods and properties; `com_get` / `com_set` / `com_call` address one member by
path; `com_eval` evaluates a Python expression against the live objects. Between
them they reach the parts with no dedicated tool (pivot tables, mail merge,
conditional formatting, slide masters, …).

## Tool families

Call `m365_help` for the live list. Summary:

- **m365_** — `m365_status` (installed / running / what is open — start here
  when something fails), `m365_list_handles`, `m365_quit_app`
- **excel_** — workbooks, sheets, `read_range` / `write_range` /
  `append_rows`, formulas, find & replace, formatting, named ranges, charts,
  macros, PDF and PNG export, `cell_info` (value + formula + format +
  precedents + comment)
- **word_** — documents, text by paragraph range, outline, styles,
  find & replace, tables, images, breaks, table of contents, comments,
  tracked changes, macros, PDF export
- **ppt_** — presentations, slides, shapes, textboxes, tables, native charts,
  speaker notes, deck-wide text replacement, PDF and image export
- **outlook_** — accounts, folder tree, message listing with server-side
  filters, DASL search, full message read, attachment saving, drafts, replies,
  forwards, send (confirmed), move / categorise / flag, calendar, contacts,
  GAL resolution, tasks
- **onenote_** — hierarchy, full-text search, page read, page create and
  append, raw XML update, publish to PDF/DOCX
- **access_** — ADO queries against .accdb/.mdb without opening the Access UI,
  CSV export, macros
- **publisher_** — open, read text, export PDF
- **com_** — `describe`, `get`, `set`, `call`, `eval`, `dispatch`
- **office_convert** — any supported file to PDF (or another Office format)

## Safety model

Nothing transmits or destroys without an explicit flag:

| Tool | Guard |
|---|---|
| `outlook_send_draft` | `confirm=true`; without it, returns a preview of recipients + subject + body head |
| `outlook_delete_message`, `outlook_delete_appointment` | `confirm=true` |
| `outlook_create_appointment` with attendees | saves the meeting, does **not** send invitations unless `save_only=false` |
| `access_execute` | `confirm=true` |
| `m365_quit_app` | `confirm=true` |

Composition is deliberately split from sending: `outlook_create_draft`,
`outlook_reply_draft` and `outlook_forward_draft` only ever write to Drafts and
return an `entry_id`. Show the user the draft, get approval, then send.

`com_eval` runs real Python in the server process — it is as powerful as the
VBA editor. This is a single-user workstation server; do not expose it beyond
the machine it runs on.

## Known limits and gotchas

- **New Outlook has no COM.** Only classic Outlook for Windows works. Check
  with `m365_status`.
- **OneNote needs a per-user registry fix on 64-bit Python.** Classic OneNote
  2016 registers its type library only under the `win32` key, so a 64-bit
  process gets `TYPE_E_LIBNOTREGISTERED` (0x8002801D) from `GetHierarchy`. Run,
  as yourself (no admin needed):
  ```powershell
  powershell -NoProfile -ExecutionPolicy Bypass -File .\tools\fix_onenote_typelib.ps1 -WhatIf
  powershell -NoProfile -ExecutionPolicy Bypass -File .\tools\fix_onenote_typelib.ps1
  ```
  It writes one key under `HKCU\Software\Classes` and has a `-Revert` switch.
  Alternative: run the server on 32-bit Python. Every other app works
  unchanged on 64-bit.
- **Parameterised properties.** Under late binding `Range.Address(False, False)`
  raises `TypeError: 'str' object is not callable` — the property already
  returned a string. Use `.Address`. The typed Excel tools handle this; if you
  hit it in `com_eval`, the error carries the hint.
- **Modal dialogs block everything.** If Office is showing a dialog or sitting
  in cell-edit mode, calls fail with `RPC_E_CALL_REJECTED` (0x80010001) or time
  out. `DisplayAlerts` is switched off for Excel and Word to avoid causing them.
- **PowerPoint cannot be hidden.** It refuses `Visible=False`; automating it
  shows a window.
- **Access needs an ACE OLEDB provider** matching the Python bitness for
  `access_query` / `access_export`. `access_run_macro` goes through the Access
  UI instead and does not need it.
- **Handles are per server run.** After a restart, re-open the file.
- **Quit can leave an orphan.** An Excel instance that COM started sometimes
  survives `Quit()` as a window-less background process. The smoke test quits
  only the apps it started; if one lingers, check that it holds no workbooks
  (`excel_list_workbooks`) before killing it.

## Development

```bash
python tests/smoke_test.py                       # everything
python tests/smoke_test.py --skip ppt,outlook    # faster loop
```

The suite drives ~92 checks across the real applications and quits only the
apps it started itself. On a stock Office 16 install with 64-bit Python, 90
pass; the two OneNote checks pass once the type library fix above is applied.
