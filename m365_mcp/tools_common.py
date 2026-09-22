"""Cross-application tools: status, conversion, object-model introspection,
and the generic COM escape hatch.

The typed tools in tools_excel / tools_word / ... cover the common ground. The
Office object models are far larger than any fixed tool list, so these three
tools let an agent reach the rest:

  com_describe  - what properties/methods does this object have?
  com_get / com_call / com_set - address one member by path
  com_eval      - evaluate a Python expression against live COM objects

com_eval executes Python in this process. It is as powerful as a VBA macro, so
it is meant for a local, single-user workstation server - do not expose this
server to anyone you would not hand a VBA editor.
"""

from __future__ import annotations

import os
from typing import Any

from . import comcore
from .comcore import (
    APPS,
    ComToolError,
    com_tool,
    get_app,
    get_handle,
    jsonable,
    list_handles,
    normalize_app,
)

CONVERTERS = {
    ".xlsx": "excel", ".xlsm": "excel", ".xlsb": "excel", ".xls": "excel",
    ".csv": "excel",
    ".docx": "word", ".doc": "word", ".rtf": "word", ".txt": "word",
    ".odt": "word", ".htm": "word", ".html": "word",
    ".pptx": "powerpoint", ".ppt": "powerpoint", ".pptm": "powerpoint",
    ".pub": "publisher",
}


# --------------------------------------------------------------------------
# status
# --------------------------------------------------------------------------


@com_tool
def m365_status(probe_running: bool = True) -> dict[str, Any]:
    """Which Office apps are installed, which are running, and what is open.

    Call this first when something fails - it separates "not installed" from
    "not running" from "COM blocked".
    """
    apps = []
    for name in sorted(APPS):
        installed = comcore.app_installed(name)
        row: dict[str, Any] = {
            "app": name,
            "progid": APPS[name],
            "installed": installed,
            "running": comcore.app_running(name) if (installed and probe_running)
            else False,
        }
        if row["running"]:
            try:
                app = get_app(name)
                row["version"] = jsonable(getattr(app, "Version", None))
                if name == "excel":
                    row["open"] = [wb.Name for wb in app.Workbooks]
                elif name == "word":
                    row["open"] = [d.Name for d in app.Documents]
                elif name == "powerpoint":
                    row["open"] = [p.Name for p in app.Presentations]
                elif name == "outlook":
                    row["open"] = [jsonable(
                        app.GetNamespace("MAPI").CurrentUser.Name)]
            except Exception as exc:  # noqa: BLE001
                row["probe_error"] = str(exc)[:200]
        apps.append(row)
    return {"applications": apps, "handles": list_handles()}


@com_tool
def m365_list_handles() -> dict[str, Any]:
    """All live handles (workbooks, documents, presentations, publications)."""
    return {"handles": list_handles()}


@com_tool
def m365_quit_app(app: str, save_changes: bool = False, confirm: bool = False) -> dict:
    """Quit an Office application. Requires confirm=True - it can lose work."""
    key = normalize_app(app)
    if not confirm:
        return {
            "ok": False,
            "error": "confirm=False - nothing was closed",
            "hint": "Quitting closes the user's open files; ask first.",
        }
    application = get_app(key)
    if key == "excel" and not save_changes:
        for wb in application.Workbooks:
            try:
                wb.Saved = True
            except Exception:  # noqa: BLE001
                pass
    application.Quit()
    comcore.forget_app(key)
    return {"quit": key, "saved": bool(save_changes)}


@com_tool
def office_convert(
    input_path: str, output_path: str, keep_open: bool = False
) -> dict[str, Any]:
    """Convert an Office file to PDF (or another format) by round-tripping COM.

    Routes .xlsx/.docx/.pptx/.pub to the right application automatically.
    """
    src = os.path.abspath(os.path.expanduser(input_path))
    dst = os.path.abspath(os.path.expanduser(output_path))
    if not os.path.exists(src):
        raise ComToolError("File not found: " + src)
    os.makedirs(os.path.dirname(dst) or ".", exist_ok=True)
    ext = os.path.splitext(src)[1].lower()
    app_key = CONVERTERS.get(ext)
    if app_key is None:
        raise ComToolError(
            "Do not know which app opens %r" % ext,
            hint="Supported: " + ", ".join(sorted(CONVERTERS)),
        )
    out_ext = os.path.splitext(dst)[1].lower()
    app = get_app(app_key)

    def _already_open(collection: Any) -> Any:
        """Never close a file the user (or another tool) already had open."""
        for item in collection:
            try:
                if os.path.abspath(str(item.FullName)).lower() == src.lower():
                    return item
            except Exception:  # noqa: BLE001
                continue
        return None

    if app_key == "excel":
        existing = _already_open(app.Workbooks)
        keep_open = keep_open or existing is not None
        doc = existing or app.Workbooks.Open(src, UpdateLinks=0, ReadOnly=True)
        try:
            if out_ext == ".pdf":
                doc.ExportAsFixedFormat(Type=0, Filename=dst)
            else:
                from .tools_excel import FILE_FORMATS

                fmt = FILE_FORMATS.get(out_ext.lstrip("."))
                if fmt is None:
                    raise ComToolError("Unsupported Excel target " + out_ext)
                doc.SaveAs(Filename=dst, FileFormat=fmt)
        finally:
            if not keep_open:
                doc.Close(SaveChanges=False)
    elif app_key == "word":
        existing = _already_open(app.Documents)
        keep_open = keep_open or existing is not None
        doc = existing or app.Documents.Open(src, ReadOnly=True)
        try:
            if out_ext == ".pdf":
                doc.ExportAsFixedFormat(OutputFileName=dst, ExportFormat=17)
            else:
                from .tools_word import WD_FORMAT

                fmt = WD_FORMAT.get(out_ext.lstrip("."))
                if fmt is None:
                    raise ComToolError("Unsupported Word target " + out_ext)
                doc.SaveAs2(FileName=dst, FileFormat=fmt)
        finally:
            if not keep_open:
                doc.Close(SaveChanges=0)
    elif app_key == "powerpoint":
        existing = _already_open(app.Presentations)
        keep_open = keep_open or existing is not None
        doc = existing or app.Presentations.Open(src, ReadOnly=-1, WithWindow=-1)
        try:
            from .tools_ppt import PP_SAVE

            fmt = PP_SAVE.get(out_ext.lstrip("."))
            if fmt is None:
                raise ComToolError("Unsupported PowerPoint target " + out_ext)
            doc.SaveAs(FileName=dst, FileFormat=fmt)
        finally:
            if not keep_open:
                doc.Saved = -1
                doc.Close()
    else:  # publisher
        doc = app.Open(src)
        try:
            doc.ExportAsFixedFormat(1, dst)
        finally:
            if not keep_open:
                doc.Close()
    return {"input": src, "output": dst, "exists": os.path.exists(dst),
            "via": app_key}


# --------------------------------------------------------------------------
# generic COM access
# --------------------------------------------------------------------------


def _root(app: str | None, handle: str | None) -> tuple[Any, dict[str, Any]]:
    """Return (root object, namespace) for a generic COM expression."""
    aliases = {"xlwb": ("wb", "excel"), "wddoc": ("doc", "word"),
               "pppres": ("pres", "powerpoint"), "pbdoc": ("pub", "publisher")}
    namespace: dict[str, Any] = {}
    obj: Any = None
    if handle:
        obj = get_handle(handle)
        alias, owner = aliases.get(handle.split(":", 1)[0], (None, None))
        if alias:
            namespace[alias] = obj
        if app is None:
            app = owner
    if app:
        application = get_app(app)
        namespace["app"] = application
        if obj is None:
            obj = application
    elif obj is None:
        raise ComToolError("Pass `app` (excel/word/...) or `handle`")
    namespace["obj"] = obj
    return obj, namespace


def _eval(expression: str, namespace: dict[str, Any]) -> Any:
    env = dict(namespace)
    env.setdefault("os", os)
    import win32com.client as _w32

    env.setdefault("win32com", _w32)
    try:
        return eval(expression, {"__builtins__": __builtins__}, env)  # noqa: S307
    except ComToolError:
        raise
    except SyntaxError as exc:
        raise ComToolError("Cannot parse expression: %s" % exc) from exc
    except TypeError as exc:
        if "not callable" in str(exc):
            raise ComToolError(
                "%s" % exc,
                hint="Parameterised properties (Range.Address, Range.Item, "
                     "Comment.Text) come back as plain values under late binding - "
                     "drop the parentheses, e.g. `.Address` instead of "
                     "`.Address(False, False)`.",
            ) from exc
        raise


@com_tool
def com_describe(
    app: str | None = None,
    handle: str | None = None,
    path: str | None = None,
    max_members: int = 200,
) -> dict[str, Any]:
    """Introspect a live COM object: its methods and properties.

    Args:
        app: excel | word | powerpoint | outlook | onenote | access | publisher
        handle: a handle from excel_open / word_open / ... (overrides app)
        path: member path from the root, e.g. "ActiveSheet.Range('A1')"

    Use this to discover object-model members that have no dedicated tool, then
    call them with com_get / com_call / com_eval.
    """
    obj, namespace = _root(app, handle)
    if path:
        obj = _eval("obj." + path if not path.startswith("(") else path, namespace)
    info: dict[str, Any] = {
        "target": path or (handle or app),
        "repr": comcore.describe_dispatch(obj) if hasattr(obj, "_oleobj_") else repr(obj),
        "type": type(obj).__name__,
    }
    methods: list[dict[str, Any]] = []
    properties: list[dict[str, Any]] = []
    try:
        ole = obj._oleobj_
        if ole.GetTypeInfoCount():
            ti = ole.GetTypeInfo()
            attr = ti.GetTypeAttr()
            for i in range(attr.cFuncs):
                fd = ti.GetFuncDesc(i)
                names = ti.GetNames(fd.memid)
                if not names:
                    continue
                entry = {
                    "name": names[0],
                    "args": list(names[1:]) or [],
                    "arg_count": int(fd.cParams),
                }
                if fd.invkind == 1:
                    methods.append(entry)
                else:
                    properties.append(
                        {"name": names[0],
                         "access": {2: "get", 4: "put", 8: "putref"}.get(
                             fd.invkind, str(fd.invkind))}
                    )
            for i in range(attr.cVars):
                vd = ti.GetVarDesc(i)
                names = ti.GetNames(vd.memid)
                if names:
                    properties.append({"name": names[0], "access": "get"})
            info["type_name"] = ti.GetDocumentation(-1)[0]
    except Exception as exc:  # noqa: BLE001 - late-bound objects may refuse
        info["introspection_note"] = "TypeInfo unavailable: %s" % str(exc)[:160]

    if not methods and not properties:
        # Fallback: pywin32's cached dispatch maps.
        for attr_name in ("_prop_map_get_", "_prop_map_put_"):
            mapping = getattr(obj, attr_name, None) or {}
            for key in mapping:
                properties.append({"name": key, "access": attr_name})
        for key in dir(obj):
            if not key.startswith("_"):
                methods.append({"name": key})

    seen = set()
    props_unique = []
    for p in properties:
        if p["name"] not in seen:
            seen.add(p["name"])
            props_unique.append(p)
    info["methods"] = methods[:max_members]
    info["properties"] = props_unique[:max_members]
    info["counts"] = {"methods": len(methods), "properties": len(props_unique)}
    return info


@com_tool
def com_get(
    path: str, app: str | None = None, handle: str | None = None,
    max_items: int = 200
) -> dict[str, Any]:
    """Read any property by path, e.g. path="ActiveSheet.UsedRange.Address()".

    Collections are expanded to a list of their items' names where possible.
    """
    _obj, namespace = _root(app, handle)
    value = _eval("obj." + path, namespace)
    if hasattr(value, "Count") and not isinstance(value, (str, bytes)):
        try:
            items = []
            for i, element in enumerate(value):
                if i >= max_items:
                    break
                items.append(jsonable(element))
            return {"path": path, "count": int(value.Count), "items": items}
        except Exception:  # noqa: BLE001 - not iterable after all
            pass
    return {"path": path, "value": jsonable(value)}


@com_tool
def com_set(
    path: str, value: Any, app: str | None = None, handle: str | None = None
) -> dict[str, Any]:
    """Set any property by path, e.g. path="ActiveSheet.Range('A1').Value"."""
    _obj, namespace = _root(app, handle)
    if "." not in path:
        raise ComToolError('path must address a member, e.g. "Range(\'A1\').Value"')
    parent_path, member = path.rsplit(".", 1)
    parent = _eval("obj." + parent_path, namespace)
    setattr(parent, member, value)
    return {"path": path, "set_to": jsonable(value)}


@com_tool
def com_call(
    path: str,
    args: list | None = None,
    app: str | None = None,
    handle: str | None = None,
    kwargs: dict | None = None,
) -> dict[str, Any]:
    """Call any method by path with positional/keyword arguments.

    Example: path="Worksheets(1).Range('A1:C9').Sort", kwargs={"Key1": ...}
    """
    _obj, namespace = _root(app, handle)
    method = _eval("obj." + path, namespace)
    if not callable(method):
        raise ComToolError("%r is not callable - use com_get instead" % path)
    result = method(*(args or []), **(kwargs or {}))
    return {"path": path, "result": jsonable(result)}


@com_tool
def com_eval(
    expression: str,
    app: str | None = None,
    handle: str | None = None,
    assign: bool = False,
) -> dict[str, Any]:
    """Escape hatch: evaluate a Python expression against live COM objects.

    Bound names: `app` (the Application), `obj` (handle target or Application),
    plus `wb` / `doc` / `pres` / `pub` when a handle of that kind is passed, and
    `win32com`, `os`.

    Examples:
        expression="app.ActiveSheet.UsedRange.Address()"
        expression="[ws.Name for ws in wb.Worksheets]"
        expression="app.ActiveDocument.Paragraphs.Count"

    With assign=True the text is exec'd instead of eval'd, so you can run
    several statements; the value of a variable named `result` is returned.
    This runs real Python in the server process - treat it like a VBA macro.
    """
    _obj, namespace = _root(app, handle)
    if assign:
        env = dict(namespace)
        env["os"] = os
        import win32com.client as _w32

        env["win32com"] = _w32
        exec(expression, {"__builtins__": __builtins__}, env)  # noqa: S102
        return {"expression": expression[:200],
                "result": jsonable(env.get("result"))}
    return {"expression": expression[:200],
            "result": jsonable(_eval(expression, namespace))}


@com_tool
def com_dispatch(
    progid: str, expression: str | None = None, new_instance: bool = False
) -> dict[str, Any]:
    """Talk to ANY COM server on this machine, not just Office.

    Example: progid="Scripting.FileSystemObject", expression="obj.Drives.Count",
    or progid="Shell.Application", or a Visio/Project install.
    """
    import win32com.client as _w32

    if new_instance:
        obj = _w32.DispatchEx(progid)
    else:
        try:
            obj = _w32.GetActiveObject(progid)
        except Exception:  # noqa: BLE001 - nothing running, create one
            obj = _w32.Dispatch(progid)
    if not expression:
        return {"progid": progid, "repr": comcore.describe_dispatch(obj)}
    return {
        "progid": progid,
        "result": jsonable(_eval(expression, {"obj": obj})),
    }


TOOLS = [
    m365_status,
    m365_list_handles,
    m365_quit_app,
    office_convert,
    com_describe,
    com_get,
    com_set,
    com_call,
    com_eval,
    com_dispatch,
]
