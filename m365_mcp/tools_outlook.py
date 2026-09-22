"""Outlook tools (COM) - classic Outlook for Windows only.

Mail, folders, search, attachments, calendar, contacts, tasks, categories.

Safety: nothing here sends anything by itself. Composition tools create DRAFTS
and return an entry_id; `outlook_send_draft` is the only tool that transmits and
it requires confirm=True, which the user has to authorise in the conversation.
"""

from __future__ import annotations

import datetime as _dt
import os
import re
from typing import Any

from .comcore import ComToolError, com_tool, get_app, jsonable

# --- Outlook enums ---
OL_FOLDERS = {
    "deleted": 3,
    "outbox": 4,
    "sent": 5,
    "inbox": 6,
    "calendar": 9,
    "contacts": 10,
    "journal": 11,
    "notes": 12,
    "tasks": 13,
    "drafts": 16,
    "junk": 23,
}
OL_MAIL = 0
OL_APPOINTMENT = 1
OL_CONTACT = 2
OL_TASK = 3
OL_FORMAT_PLAIN = 1
OL_FORMAT_HTML = 2
OL_REPLY_ALL = "ReplyAll"


def _ns() -> Any:
    return get_app("outlook").GetNamespace("MAPI")


def _fmt_dt(value: _dt.datetime) -> str:
    """Outlook Restrict wants US-style short date-time regardless of locale."""
    return value.strftime("%m/%d/%Y %I:%M %p")


def _parse_dt(value: str | None) -> _dt.datetime | None:
    if not value:
        return None
    text = str(value).strip()
    if re.fullmatch(r"-?\d+\s*[dhw]", text, re.I):  # relative: 7d, 12h, 2w
        n = int(re.findall(r"-?\d+", text)[0])
        unit = text[-1].lower()
        delta = {"d": _dt.timedelta(days=n), "h": _dt.timedelta(hours=n),
                 "w": _dt.timedelta(weeks=n)}[unit]
        return _dt.datetime.now() - abs(delta)
    for fmt in ("%Y-%m-%d %H:%M", "%Y-%m-%dT%H:%M", "%Y-%m-%d %H:%M:%S",
                "%Y-%m-%d", "%m/%d/%Y %H:%M", "%m/%d/%Y"):
        try:
            return _dt.datetime.strptime(text, fmt)
        except ValueError:
            continue
    raise ComToolError(
        "Cannot parse datetime %r" % value,
        hint="Use 'YYYY-MM-DD', 'YYYY-MM-DD HH:MM', or a relative '7d' / '12h'.",
    )


def _folder(spec: str | None = None, store: str | None = None) -> Any:
    """Resolve a folder from a name, a well-known key, a path, or an EntryID."""
    ns = _ns()
    if not spec:
        return ns.GetDefaultFolder(OL_FOLDERS["inbox"])
    text = str(spec).strip()

    if store:
        root = None
        for st in ns.Stores:
            name = str(st.DisplayName)
            smtp = ""
            try:
                smtp = str(st.GetRootFolder().Store.DisplayName)
            except Exception:  # noqa: BLE001
                pass
            if store.lower() in (name.lower(), smtp.lower()):
                root = st.GetRootFolder()
                break
        if root is None:
            names = [str(s.DisplayName) for s in ns.Stores]
            raise ComToolError(
                "No store named %r" % store, hint="Stores: " + ", ".join(names)
            )
        if text.lower() in OL_FOLDERS:
            for fld in root.Folders:
                if str(fld.Name).lower() in (text.lower(), "inbox"):
                    return fld
        return _walk_path(root, text)

    key = text.lower()
    if key in OL_FOLDERS:
        return ns.GetDefaultFolder(OL_FOLDERS[key])
    if len(text) > 40 and " " not in text:  # looks like an EntryID
        try:
            return ns.GetFolderFromID(text)
        except Exception:  # noqa: BLE001
            pass
    if "/" in text or "\\" in text:
        parts = re.split(r"[\\/]+", text.strip("\\/"))
        for st in ns.Stores:
            if str(st.DisplayName).lower() == parts[0].lower():
                return _walk_path(st.GetRootFolder(), "/".join(parts[1:]))
        return _walk_path(ns.GetDefaultFolder(OL_FOLDERS["inbox"]).Parent,
                          "/".join(parts))
    # bare name: search the default store one level deep, then everywhere
    inbox = ns.GetDefaultFolder(OL_FOLDERS["inbox"])
    for fld in inbox.Folders:
        if str(fld.Name).lower() == key:
            return fld
    for fld in inbox.Parent.Folders:
        if str(fld.Name).lower() == key:
            return fld
    raise ComToolError(
        "Folder %r not found" % spec,
        hint="Call outlook_list_folders to see the tree, or pass a path.",
    )


def _walk_path(root: Any, path: str) -> Any:
    node = root
    for part in [p for p in re.split(r"[\\/]+", path) if p]:
        found = None
        for fld in node.Folders:
            if str(fld.Name).lower() == part.lower():
                found = fld
                break
        if found is None:
            names = [str(f.Name) for f in node.Folders]
            raise ComToolError(
                "No subfolder %r under %r" % (part, node.Name),
                hint="Children: " + ", ".join(names),
            )
        node = found
    return node


def _item(entry_id: str) -> Any:
    try:
        return _ns().GetItemFromID(entry_id)
    except Exception as exc:  # noqa: BLE001
        raise ComToolError(
            "No item with EntryID %r" % entry_id,
            hint="EntryIDs come from outlook_list_messages / outlook_search.",
        ) from exc


def _mail_row(item: Any, body_chars: int = 0) -> dict[str, Any]:
    row: dict[str, Any] = {}
    for label, attr in (
        ("entry_id", "EntryID"),
        ("subject", "Subject"),
        ("sender", "SenderName"),
        ("sender_email", "SenderEmailAddress"),
        ("to", "To"),
        ("cc", "CC"),
        ("received", "ReceivedTime"),
        ("sent", "SentOn"),
        ("unread", "UnRead"),
        ("importance", "Importance"),
        ("categories", "Categories"),
        ("size", "Size"),
        ("has_attachments", "Attachments"),
        ("conversation", "ConversationTopic"),
    ):
        try:
            value = getattr(item, attr)
            if attr == "Attachments":
                value = int(value.Count) > 0
            row[label] = jsonable(value)
        except Exception:  # noqa: BLE001 - not every item class has every field
            row[label] = None
    if body_chars:
        try:
            body = str(item.Body or "")
            row["body"] = body[:body_chars]
            row["body_truncated"] = len(body) > body_chars
        except Exception:  # noqa: BLE001
            row["body"] = None
    return row


def _restrict(items: Any, clauses: list[str]) -> Any:
    if not clauses:
        return items
    return items.Restrict(" AND ".join(clauses))


# --------------------------------------------------------------------------
# accounts and folders
# --------------------------------------------------------------------------


@com_tool
def outlook_accounts() -> dict[str, Any]:
    """List the mail accounts and stores available in this Outlook profile."""
    app = get_app("outlook")
    ns = app.GetNamespace("MAPI")
    accounts = []
    for acc in ns.Accounts:
        accounts.append(
            {
                "display_name": str(acc.DisplayName),
                "smtp": jsonable(getattr(acc, "SmtpAddress", None)),
                "type": jsonable(getattr(acc, "AccountType", None)),
            }
        )
    stores = []
    for st in ns.Stores:
        try:
            stores.append(
                {
                    "name": str(st.DisplayName),
                    "path": jsonable(getattr(st, "FilePath", None)),
                    "default": bool(st.ExchangeStoreType is not None),
                }
            )
        except Exception:  # noqa: BLE001
            continue
    return {
        "outlook_version": jsonable(app.Version),
        "current_user": jsonable(ns.CurrentUser.Name),
        "accounts": accounts,
        "stores": stores,
    }


@com_tool
def outlook_list_folders(
    store: str | None = None, max_depth: int = 3, with_counts: bool = True
) -> dict[str, Any]:
    """Walk the folder tree (of one store, or of every store)."""
    ns = _ns()
    roots = []
    if store:
        roots = [_folder("", store).Parent if False else None]
        for st in ns.Stores:
            if str(st.DisplayName).lower() == store.lower():
                roots = [st.GetRootFolder()]
                break
        if roots == [None]:
            raise ComToolError("No store named %r" % store)
    else:
        roots = [st.GetRootFolder() for st in ns.Stores]

    rows: list[dict[str, Any]] = []

    def walk(folder: Any, depth: int, prefix: str) -> None:
        if depth > max_depth:
            return
        for fld in folder.Folders:
            path = prefix + "/" + str(fld.Name)
            row = {"path": path, "name": str(fld.Name), "depth": depth}
            if with_counts:
                try:
                    row["items"] = int(fld.Items.Count)
                    row["unread"] = int(fld.UnReadItemCount)
                except Exception:  # noqa: BLE001
                    pass
            rows.append(row)
            walk(fld, depth + 1, path)

    for root in roots:
        rows.append({"path": str(root.Name), "name": str(root.Name), "depth": 0,
                     "is_store_root": True})
        walk(root, 1, str(root.Name))
    return {"count": len(rows), "folders": rows}


@com_tool
def outlook_folder_stats(folder: str = "inbox", store: str | None = None) -> dict:
    """Item/unread counts and the newest item of a folder."""
    fld = _folder(folder, store)
    items = fld.Items
    newest = None
    try:
        items.Sort("[ReceivedTime]", True)
        if int(items.Count):
            newest = _mail_row(items.GetFirst())
    except Exception:  # noqa: BLE001 - non-mail folders have no ReceivedTime
        pass
    return {
        "folder": str(fld.FolderPath),
        "items": int(items.Count),
        "unread": int(fld.UnReadItemCount),
        "newest": newest,
    }


# --------------------------------------------------------------------------
# reading mail
# --------------------------------------------------------------------------


@com_tool
def outlook_list_messages(
    folder: str = "inbox",
    store: str | None = None,
    limit: int = 25,
    unread_only: bool = False,
    since: str | None = None,
    until: str | None = None,
    sender_contains: str | None = None,
    subject_contains: str | None = None,
    body_chars: int = 0,
    newest_first: bool = True,
) -> dict[str, Any]:
    """List messages in a folder with server-side filters.

    Args:
        since/until: "YYYY-MM-DD[ HH:MM]" or relative like "7d" / "12h"
        body_chars: >0 also returns that many characters of each body
    """
    fld = _folder(folder, store)
    items = fld.Items
    clauses = []
    if unread_only:
        clauses.append("[UnRead] = True")
    start, end = _parse_dt(since), _parse_dt(until)
    if start:
        clauses.append("[ReceivedTime] >= '%s'" % _fmt_dt(start))
    if end:
        clauses.append("[ReceivedTime] <= '%s'" % _fmt_dt(end))
    try:
        items.Sort("[ReceivedTime]", bool(newest_first))
    except Exception:  # noqa: BLE001
        pass
    items = _restrict(items, clauses)

    rows = []
    item = items.GetFirst()
    scanned = 0
    while item is not None and len(rows) < limit and scanned < 5000:
        scanned += 1
        try:
            if sender_contains:
                who = "%s %s" % (
                    getattr(item, "SenderName", "") or "",
                    getattr(item, "SenderEmailAddress", "") or "",
                )
                if sender_contains.lower() not in who.lower():
                    item = items.GetNext()
                    continue
            if subject_contains:
                subj = str(getattr(item, "Subject", "") or "")
                if subject_contains.lower() not in subj.lower():
                    item = items.GetNext()
                    continue
            rows.append(_mail_row(item, body_chars))
        except Exception:  # noqa: BLE001 - skip unreadable items
            pass
        item = items.GetNext()
    return {
        "folder": str(fld.FolderPath),
        "count": len(rows),
        "scanned": scanned,
        "messages": rows,
    }


@com_tool
def outlook_search(
    query: str,
    folder: str = "inbox",
    store: str | None = None,
    fields: str = "subject,body,sender",
    limit: int = 25,
    since: str | None = None,
    body_chars: int = 0,
) -> dict[str, Any]:
    """Keyword search inside a folder (DASL contains-search, subfolders excluded).

    Args:
        query: plain text; it is matched case-insensitively
        fields: comma list of subject | body | sender | to | categories
    """
    fld = _folder(folder, store)
    field_map = {
        "subject": "urn:schemas:httpmail:subject",
        "body": "urn:schemas:httpmail:textdescription",
        "sender": "urn:schemas:httpmail:fromname",
        "sender_email": "urn:schemas:httpmail:fromemail",
        "to": "urn:schemas:httpmail:displayto",
        "categories": "urn:schemas-microsoft-com:office:office#Keywords",
    }
    wanted = [f.strip().lower() for f in fields.split(",") if f.strip()]
    unknown = [f for f in wanted if f not in field_map]
    if unknown:
        raise ComToolError(
            "Unknown search field(s): " + ", ".join(unknown),
            hint="Use " + ", ".join(field_map),
        )
    escaped = query.replace("'", "''")
    parts = [
        '"%s" like \'%%%s%%\'' % (field_map[f], escaped) for f in wanted
    ]
    dasl = "@SQL=" + " OR ".join(parts)
    start = _parse_dt(since)
    if start:
        dasl = '@SQL=(%s) AND "urn:schemas:httpmail:datereceived" >= \'%s\'' % (
            " OR ".join(parts),
            start.strftime("%Y-%m-%d %H:%M"),
        )
    try:
        items = fld.Items.Restrict(dasl)
    except Exception as exc:  # noqa: BLE001
        raise ComToolError(
            "Outlook rejected the search filter",
            hint="Try a simpler query, or use outlook_list_messages with "
                 "subject_contains.",
        ) from exc
    try:
        items.Sort("[ReceivedTime]", True)
    except Exception:  # noqa: BLE001
        pass
    rows = []
    item = items.GetFirst()
    while item is not None and len(rows) < limit:
        try:
            rows.append(_mail_row(item, body_chars))
        except Exception:  # noqa: BLE001
            pass
        item = items.GetNext()
    return {
        "query": query,
        "folder": str(fld.FolderPath),
        "count": len(rows),
        "messages": rows,
    }


@com_tool
def outlook_get_message(
    entry_id: str, body_format: str = "text", max_chars: int = 20000
) -> dict[str, Any]:
    """Read one message in full. body_format: text | html."""
    item = _item(entry_id)
    row = _mail_row(item)
    body = ""
    try:
        body = str(
            item.HTMLBody if body_format.lower() == "html" else item.Body or ""
        )
    except Exception:  # noqa: BLE001
        pass
    row["body_format"] = body_format
    row["body"] = body[:max_chars]
    row["body_truncated"] = len(body) > max_chars
    try:
        row["attachments"] = [
            {"index": i + 1, "name": str(att.FileName), "size": int(att.Size)}
            for i, att in enumerate(item.Attachments)
        ]
    except Exception:  # noqa: BLE001
        row["attachments"] = []
    try:
        row["recipients"] = [
            {"name": str(r.Name), "address": jsonable(r.Address), "type": int(r.Type)}
            for r in item.Recipients
        ]
    except Exception:  # noqa: BLE001
        pass
    return row


@com_tool
def outlook_save_attachments(
    entry_id: str,
    output_dir: str,
    name_filter: str | None = None,
    index: int | None = None,
) -> dict[str, Any]:
    """Save a message's attachments to a folder."""
    item = _item(entry_id)
    out = os.path.abspath(os.path.expanduser(output_dir))
    os.makedirs(out, exist_ok=True)
    saved = []
    for i, att in enumerate(item.Attachments, start=1):
        name = str(att.FileName)
        if index and i != index:
            continue
        if name_filter and name_filter.lower() not in name.lower():
            continue
        safe = re.sub(r'[<>:"/\\|?*]', "_", name)
        path = os.path.join(out, safe)
        att.SaveAsFile(path)
        saved.append({"name": name, "path": path, "size": os.path.getsize(path)})
    return {"directory": out, "count": len(saved), "saved": saved}


# --------------------------------------------------------------------------
# composing (drafts only) and sending (explicit confirm)
# --------------------------------------------------------------------------


@com_tool
def outlook_create_draft(
    to: str,
    subject: str,
    body: str,
    cc: str | None = None,
    bcc: str | None = None,
    html: bool = False,
    attachments: list | None = None,
    account: str | None = None,
    display: bool = False,
) -> dict[str, Any]:
    """Create an unsent draft and return its entry_id. Nothing is transmitted.

    Args:
        to/cc/bcc: semicolon-separated addresses
        html: treat `body` as HTML
        attachments: list of file paths
        account: send-from account display name or SMTP address
        display: pop the draft open on screen for the user to inspect
    """
    app = get_app("outlook")
    mail = app.CreateItem(OL_MAIL)
    mail.To = to
    if cc:
        mail.CC = cc
    if bcc:
        mail.BCC = bcc
    mail.Subject = subject
    if html:
        mail.HTMLBody = body
    else:
        mail.Body = body
    for path in attachments or []:
        full = os.path.abspath(os.path.expanduser(path))
        if not os.path.exists(full):
            raise ComToolError("Attachment not found: " + full)
        mail.Attachments.Add(full)
    if account:
        ns = app.GetNamespace("MAPI")
        matched = None
        for acc in ns.Accounts:
            if account.lower() in (
                str(acc.DisplayName).lower(),
                str(getattr(acc, "SmtpAddress", "") or "").lower(),
            ):
                matched = acc
                break
        if matched is None:
            raise ComToolError("No account matching %r" % account)
        mail.SendUsingAccount = matched
    mail.Save()
    if display:
        mail.Display(False)
    return {
        "entry_id": str(mail.EntryID),
        "subject": subject,
        "to": to,
        "sent": False,
        "note": "Draft saved in Drafts. Use outlook_send_draft to transmit.",
    }


@com_tool
def outlook_reply_draft(
    entry_id: str,
    body: str,
    reply_all: bool = False,
    html: bool = False,
    prepend: bool = True,
) -> dict[str, Any]:
    """Create a reply draft to an existing message (not sent)."""
    item = _item(entry_id)
    reply = item.ReplyAll() if reply_all else item.Reply()
    if html:
        original = str(reply.HTMLBody or "")
        reply.HTMLBody = (body + original) if prepend else (original + body)
    else:
        original = str(reply.Body or "")
        reply.Body = (body + "\n\n" + original) if prepend else (original + "\n" + body)
    reply.Save()
    return {
        "entry_id": str(reply.EntryID),
        "subject": jsonable(reply.Subject),
        "reply_all": reply_all,
        "sent": False,
    }


@com_tool
def outlook_forward_draft(
    entry_id: str, to: str, comment: str = "", html: bool = False
) -> dict[str, Any]:
    """Create a forward draft of an existing message (not sent)."""
    item = _item(entry_id)
    fwd = item.Forward()
    fwd.To = to
    if comment:
        if html:
            fwd.HTMLBody = comment + str(fwd.HTMLBody or "")
        else:
            fwd.Body = comment + "\n\n" + str(fwd.Body or "")
    fwd.Save()
    return {"entry_id": str(fwd.EntryID), "to": to, "sent": False}


@com_tool
def outlook_send_draft(entry_id: str, confirm: bool = False) -> dict[str, Any]:
    """Send an existing draft. THIS TRANSMITS MAIL.

    Requires confirm=True, and the user must have approved the send in the
    conversation - never call this off your own judgement.
    """
    if not confirm:
        item = _item(entry_id)
        return {
            "ok": False,
            "error": "confirm=False - nothing was sent",
            "preview": {
                "to": jsonable(getattr(item, "To", None)),
                "cc": jsonable(getattr(item, "CC", None)),
                "subject": jsonable(getattr(item, "Subject", None)),
                "body_head": str(getattr(item, "Body", "") or "")[:500],
            },
            "hint": "Show this preview to the user; call again with confirm=true "
                    "once they approve.",
        }
    item = _item(entry_id)
    to = jsonable(getattr(item, "To", None))
    subject = jsonable(getattr(item, "Subject", None))
    item.Send()
    return {"sent": True, "to": to, "subject": subject}


@com_tool
def outlook_delete_draft(entry_id: str) -> dict[str, Any]:
    """Delete a draft you created (moves it to Deleted Items)."""
    item = _item(entry_id)
    subject = jsonable(getattr(item, "Subject", None))
    item.Delete()
    return {"deleted": True, "subject": subject}


# --------------------------------------------------------------------------
# organising mail
# --------------------------------------------------------------------------


@com_tool
def outlook_move_message(
    entry_id: str, folder: str, store: str | None = None
) -> dict[str, Any]:
    """Move a message to another folder."""
    item = _item(entry_id)
    target = _folder(folder, store)
    moved = item.Move(target)
    return {"moved_to": str(target.FolderPath), "entry_id": str(moved.EntryID)}


@com_tool
def outlook_update_message(
    entry_id: str,
    unread: bool | None = None,
    categories: str | None = None,
    add_category: str | None = None,
    remove_category: str | None = None,
    flag_status: str | None = None,
    importance: str | None = None,
) -> dict[str, Any]:
    """Mark read/unread, set categories, flag or change importance."""
    item = _item(entry_id)
    applied = []
    if unread is not None:
        item.UnRead = bool(unread)
        applied.append("unread")
    current = [c.strip() for c in str(getattr(item, "Categories", "") or "").split(",")
               if c.strip()]
    if categories is not None:
        current = [c.strip() for c in categories.split(",") if c.strip()]
        applied.append("categories")
    if add_category and add_category not in current:
        current.append(add_category)
        applied.append("add_category")
    if remove_category and remove_category in current:
        current.remove(remove_category)
        applied.append("remove_category")
    if categories is not None or add_category or remove_category:
        item.Categories = ", ".join(current)
    if flag_status is not None:
        item.FlagStatus = {"none": 0, "complete": 1, "marked": 2}.get(
            flag_status.lower(), 2
        )
        applied.append("flag_status")
    if importance is not None:
        item.Importance = {"low": 0, "normal": 1, "high": 2}.get(
            importance.lower(), 1
        )
        applied.append("importance")
    item.Save()
    return {"entry_id": entry_id, "applied": applied,
            "categories": jsonable(item.Categories)}


@com_tool
def outlook_delete_message(entry_id: str, confirm: bool = False) -> dict[str, Any]:
    """Move a message to Deleted Items. Requires confirm=True."""
    item = _item(entry_id)
    subject = jsonable(getattr(item, "Subject", None))
    if not confirm:
        return {
            "ok": False,
            "error": "confirm=False - nothing was deleted",
            "subject": subject,
            "hint": "Ask the user, then call again with confirm=true.",
        }
    item.Delete()
    return {"deleted": True, "subject": subject,
            "note": "The item is in Deleted Items, not purged."}


@com_tool
def outlook_categories(
    action: str = "list", name: str | None = None, color: int | None = None
) -> dict[str, Any]:
    """Master category list. action: list | add | remove."""
    ns = _ns()
    cats = ns.Categories
    act = action.lower()
    if act == "add":
        if not name:
            raise ComToolError("name is required to add a category")
        cats.Add(name, color if color is not None else 0)
        return {"added": name}
    if act == "remove":
        if not name:
            raise ComToolError("name is required to remove a category")
        cats.Remove(name)
        return {"removed": name}
    rows = []
    for cat in cats:
        try:
            rows.append({"name": str(cat.Name), "color": int(cat.Color)})
        except Exception:  # noqa: BLE001
            continue
    return {"count": len(rows), "categories": rows}


# --------------------------------------------------------------------------
# calendar, contacts, tasks
# --------------------------------------------------------------------------


@com_tool
def outlook_calendar_list(
    start: str | None = None,
    end: str | None = None,
    limit: int = 50,
    store: str | None = None,
    include_recurring: bool = True,
) -> dict[str, Any]:
    """List calendar items in a window (default: now to +7 days)."""
    fld = _folder("calendar", store)
    items = fld.Items
    if include_recurring:
        items.IncludeRecurrences = True
    items.Sort("[Start]")
    s = _parse_dt(start) or _dt.datetime.now().replace(hour=0, minute=0, second=0,
                                                       microsecond=0)
    e = _parse_dt(end) or (s + _dt.timedelta(days=7))
    restriction = "[Start] >= '%s' AND [Start] <= '%s'" % (_fmt_dt(s), _fmt_dt(e))
    items = items.Restrict(restriction)
    rows = []
    item = items.GetFirst()
    while item is not None and len(rows) < limit:
        try:
            rows.append(
                {
                    "entry_id": str(item.EntryID),
                    "subject": jsonable(item.Subject),
                    "start": jsonable(item.Start),
                    "end": jsonable(item.End),
                    "location": jsonable(item.Location),
                    "organizer": jsonable(getattr(item, "Organizer", None)),
                    "all_day": bool(item.AllDayEvent),
                    "recurring": bool(item.IsRecurring),
                    "required": jsonable(getattr(item, "RequiredAttendees", None)),
                    "categories": jsonable(item.Categories),
                }
            )
        except Exception:  # noqa: BLE001
            pass
        item = items.GetNext()
    return {
        "window": [s.isoformat(), e.isoformat()],
        "count": len(rows),
        "appointments": rows,
    }


@com_tool
def outlook_create_appointment(
    subject: str,
    start: str,
    end: str | None = None,
    duration_minutes: int = 60,
    location: str | None = None,
    body: str | None = None,
    attendees: str | None = None,
    reminder_minutes: int | None = 15,
    all_day: bool = False,
    busy_status: str = "busy",
    save_only: bool = True,
) -> dict[str, Any]:
    """Create a calendar appointment (or an unsent meeting request draft).

    With `attendees`, the item becomes a meeting request. It is SAVED, not sent:
    sending invitations is a user decision - open it in Outlook and send, or use
    save_only=False only after the user approves.
    """
    app = get_app("outlook")
    appt = app.CreateItem(OL_APPOINTMENT)
    appt.Subject = subject
    s = _parse_dt(start)
    appt.Start = s
    if end:
        appt.End = _parse_dt(end)
    else:
        appt.Duration = int(duration_minutes)
    if location:
        appt.Location = location
    if body:
        appt.Body = body
    if all_day:
        appt.AllDayEvent = True
    appt.BusyStatus = {"free": 0, "tentative": 1, "busy": 2, "out": 3}.get(
        busy_status.lower(), 2
    )
    if reminder_minutes is not None:
        appt.ReminderSet = True
        appt.ReminderMinutesBeforeStart = int(reminder_minutes)
    if attendees:
        appt.MeetingStatus = 1  # olMeeting
        appt.RequiredAttendees = attendees
    appt.Save()
    result = {
        "entry_id": str(appt.EntryID),
        "subject": subject,
        "start": jsonable(appt.Start),
        "end": jsonable(appt.End),
        "is_meeting": bool(attendees),
        "sent": False,
    }
    if attendees and not save_only:
        appt.Send()
        result["sent"] = True
    elif attendees:
        result["note"] = "Meeting saved but invitations NOT sent (save_only=true)."
    return result


@com_tool
def outlook_delete_appointment(entry_id: str, confirm: bool = False) -> dict:
    """Delete a calendar item. Requires confirm=True."""
    item = _item(entry_id)
    subject = jsonable(getattr(item, "Subject", None))
    if not confirm:
        return {"ok": False, "error": "confirm=False - nothing was deleted",
                "subject": subject}
    item.Delete()
    return {"deleted": True, "subject": subject}


@com_tool
def outlook_contacts_list(
    query: str | None = None, limit: int = 50, store: str | None = None
) -> dict[str, Any]:
    """List or search contacts."""
    fld = _folder("contacts", store)
    rows = []
    for item in fld.Items:
        if len(rows) >= limit:
            break
        try:
            name = str(getattr(item, "FullName", "") or "")
            email = str(getattr(item, "Email1Address", "") or "")
            company = str(getattr(item, "CompanyName", "") or "")
            blob = " ".join([name, email, company]).lower()
            if query and query.lower() not in blob:
                continue
            rows.append(
                {
                    "entry_id": str(item.EntryID),
                    "name": name,
                    "email": email,
                    "company": company,
                    "job_title": jsonable(getattr(item, "JobTitle", None)),
                    "mobile": jsonable(getattr(item, "MobileTelephoneNumber", None)),
                }
            )
        except Exception:  # noqa: BLE001
            continue
    return {"count": len(rows), "contacts": rows}


@com_tool
def outlook_resolve_recipient(name: str) -> dict[str, Any]:
    """Resolve a name against the address book (GAL) - check before sending."""
    ns = _ns()
    rec = ns.CreateRecipient(name)
    resolved = bool(rec.Resolve())
    out: dict[str, Any] = {"input": name, "resolved": resolved}
    if resolved:
        out["name"] = str(rec.Name)
        out["address"] = jsonable(rec.Address)
        try:
            entry = rec.AddressEntry.GetExchangeUser()
            if entry is not None:
                out["smtp"] = str(entry.PrimarySmtpAddress)
                out["job_title"] = jsonable(entry.JobTitle)
                out["department"] = jsonable(entry.Department)
        except Exception:  # noqa: BLE001 - non-Exchange address
            pass
    return out


@com_tool
def outlook_tasks_list(
    include_completed: bool = False, limit: int = 50, store: str | None = None
) -> dict[str, Any]:
    """List tasks."""
    fld = _folder("tasks", store)
    rows = []
    for item in fld.Items:
        if len(rows) >= limit:
            break
        try:
            complete = bool(item.Complete)
            if complete and not include_completed:
                continue
            rows.append(
                {
                    "entry_id": str(item.EntryID),
                    "subject": jsonable(item.Subject),
                    "due": jsonable(item.DueDate),
                    "status": int(item.Status),
                    "complete": complete,
                    "percent": int(item.PercentComplete),
                    "categories": jsonable(item.Categories),
                }
            )
        except Exception:  # noqa: BLE001
            continue
    return {"count": len(rows), "tasks": rows}


@com_tool
def outlook_create_task(
    subject: str,
    due: str | None = None,
    body: str | None = None,
    reminder: str | None = None,
    categories: str | None = None,
) -> dict[str, Any]:
    """Create a task in the default Tasks folder."""
    app = get_app("outlook")
    task = app.CreateItem(OL_TASK)
    task.Subject = subject
    if due:
        task.DueDate = _parse_dt(due)
    if body:
        task.Body = body
    if reminder:
        task.ReminderSet = True
        task.ReminderTime = _parse_dt(reminder)
    if categories:
        task.Categories = categories
    task.Save()
    return {"entry_id": str(task.EntryID), "subject": subject,
            "due": jsonable(task.DueDate)}


TOOLS = [
    outlook_accounts,
    outlook_list_folders,
    outlook_folder_stats,
    outlook_list_messages,
    outlook_search,
    outlook_get_message,
    outlook_save_attachments,
    outlook_create_draft,
    outlook_reply_draft,
    outlook_forward_draft,
    outlook_send_draft,
    outlook_delete_draft,
    outlook_move_message,
    outlook_update_message,
    outlook_delete_message,
    outlook_categories,
    outlook_calendar_list,
    outlook_create_appointment,
    outlook_delete_appointment,
    outlook_contacts_list,
    outlook_resolve_recipient,
    outlook_tasks_list,
    outlook_create_task,
]
