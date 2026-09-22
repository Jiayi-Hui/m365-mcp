"""COM core for the M365 MCP server.

Everything that touches COM goes through ONE dedicated STA thread:

  * COM apartments are thread-affine. pywin32 objects created on thread A
    cannot be safely used from thread B, and FastMCP runs sync tools on
    arbitrary anyio worker threads. So we own a single long-lived thread,
    CoInitialize it once, and marshal every call onto it.
  * That thread also owns the handle registry (open workbooks, documents,
    presentations, mail items), so handles stay valid across tool calls.

Public surface used by the tools_* modules:

  run(fn, timeout)      -> execute fn() on the COM thread, return its value
  com_tool(fn)          -> decorator: run on COM thread + uniform error dict
  get_app(name, ...)    -> attach to (or launch) an Office application
  put_handle / get_handle -> handle registry
  jsonable(value)       -> COM value -> JSON-safe Python
"""

from __future__ import annotations

import datetime as _dt
import functools
import itertools
import os
import queue
import threading
import traceback
from concurrent.futures import Future
from typing import Any, Callable

import pythoncom
import win32com.client

# --------------------------------------------------------------------------
# Single STA worker thread
# --------------------------------------------------------------------------

DEFAULT_TIMEOUT = 180.0


class _ComThread:
    def __init__(self) -> None:
        self._q: queue.Queue = queue.Queue()
        self._ident: int | None = None
        self._thread = threading.Thread(
            target=self._loop, name="m365-com-sta", daemon=True
        )
        self._thread.start()

    def _loop(self) -> None:
        # STA: Office automation requires it; MTA causes marshalling errors.
        self._ident = threading.get_ident()
        pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
        try:
            while True:
                item = self._q.get()
                if item is None:
                    return
                fn, fut = item
                if not fut.set_running_or_notify_cancel():
                    continue
                try:
                    fut.set_result(fn())
                except BaseException as exc:  # noqa: BLE001 - surfaced to caller
                    fut.set_exception(exc)
        finally:
            pythoncom.CoUninitialize()

    def call(self, fn: Callable[[], Any], timeout: float = DEFAULT_TIMEOUT) -> Any:
        # Re-entrant: a tool that calls another tool is already ON this thread.
        # Queueing from here would wait for a worker that is busy waiting for us.
        if threading.get_ident() == self._ident:
            return fn()
        fut: Future = Future()
        self._q.put((fn, fut))
        return fut.result(timeout=timeout)


_worker = _ComThread()


def run(fn: Callable[[], Any], timeout: float = DEFAULT_TIMEOUT) -> Any:
    """Run fn() on the COM thread and wait for its result."""
    return _worker.call(fn, timeout=timeout)


# --------------------------------------------------------------------------
# Errors
# --------------------------------------------------------------------------


class ComToolError(Exception):
    """Expected, user-facing failure: reported without a traceback."""

    def __init__(self, message: str, hint: str | None = None) -> None:
        super().__init__(message)
        self.hint = hint


def describe_com_error(exc: BaseException) -> tuple[str, str | None]:
    """Turn a pywin32 com_error into (message, hint)."""
    msg = str(exc)
    hint = None
    args = getattr(exc, "args", ())
    if args and isinstance(args[0], int):
        hresult = args[0] & 0xFFFFFFFF
        detail = args[2] if len(args) > 2 and args[2] else None
        scode_text = None
        if isinstance(detail, tuple) and len(detail) > 2 and detail[2]:
            scode_text = str(detail[2]).strip()
        msg = "COM error 0x%08X" % hresult
        if scode_text:
            msg = msg + ": " + scode_text
        if hresult == 0x800401E3:  # MK_E_UNAVAILABLE
            hint = "Application is not running; pass new_instance=true to launch it."
        elif hresult == 0x80010001:  # RPC_E_CALL_REJECTED
            hint = (
                "Office rejected the call - it is usually showing a modal dialog "
                "or is in cell-edit mode. Clear the dialog and retry."
            )
        elif hresult == 0x800A03EC:
            hint = "Office rejected the argument (bad range/name/format for this call)."
        elif hresult in (0x80080005, 0x800706BA):
            hint = "The Office process died or is not reachable; reopen the app."
        elif hresult == 0x8002801D:  # TYPE_E_LIBNOTREGISTERED
            hint = (
                "The app's type library is not registered for this process "
                "bitness. Classic OneNote registers only the win32 key, so a "
                "64-bit Python cannot drive it. Run "
                "tools/fix_onenote_typelib.ps1 (writes HKCU, no admin needed) "
                "and restart the server, or run this server on 32-bit Python."
            )
    return msg, hint


def com_tool(
    fn: Callable[..., Any] | None = None, *, timeout: float = DEFAULT_TIMEOUT
) -> Callable[..., Any]:
    """Decorator for MCP tool functions.

    Runs the body on the COM thread and normalises the result to
    {"ok": True, ...} / {"ok": False, "error": ..., "hint": ...}.

    Use `@com_tool(timeout=900)` for surveys that legitimately take minutes;
    the default guards against Office sitting behind a modal dialog forever.
    """
    if fn is None:
        return lambda f: com_tool(f, timeout=timeout)
    default_timeout = timeout

    @functools.wraps(fn)
    def wrapper(*args: Any, **kwargs: Any) -> dict[str, Any]:
        timeout = float(kwargs.pop("_timeout", None) or default_timeout)
        try:
            result = run(lambda: fn(*args, **kwargs), timeout=timeout)
        except ComToolError as exc:
            return {"ok": False, "error": str(exc), "hint": exc.hint}
        except TimeoutError:
            return {
                "ok": False,
                "error": "COM call timed out after %.0fs" % timeout,
                "hint": "Office is likely blocked on a modal dialog on the desktop.",
            }
        except Exception as exc:  # noqa: BLE001
            message, hint = describe_com_error(exc)
            return {
                "ok": False,
                "error": type(exc).__name__ + ": " + message,
                "hint": hint,
                "traceback": traceback.format_exc(limit=4),
            }
        if isinstance(result, dict):
            result.setdefault("ok", True)
            return result
        return {"ok": True, "result": jsonable(result)}

    return wrapper


# --------------------------------------------------------------------------
# JSON encoding of COM values
# --------------------------------------------------------------------------


def jsonable(value: Any, _depth: int = 0) -> Any:
    """Convert a COM/pywin32 value into something json.dumps can handle."""
    if value is None or isinstance(value, (bool, int, float, str)):
        return value
    if isinstance(value, (_dt.datetime, _dt.date, _dt.time)):
        try:
            return value.isoformat()
        except Exception:  # noqa: BLE001 - pywintypes edge cases
            return str(value)
    if isinstance(value, (bytes, bytearray)):
        return "<%d bytes>" % len(value)
    if isinstance(value, dict):
        return {str(k): jsonable(v, _depth + 1) for k, v in value.items()}
    if isinstance(value, (list, tuple, set)):
        if _depth > 6:
            return "<nested sequence len=%d>" % len(value)
        return [jsonable(v, _depth + 1) for v in value]
    if isinstance(value, win32com.client.CDispatch):
        return describe_dispatch(value)
    # pywintypes.datetime and friends
    for attr in ("isoformat", "Format"):
        method = getattr(value, attr, None)
        if callable(method):
            try:
                return method()
            except Exception:  # noqa: BLE001
                pass
    return str(value)


def describe_dispatch(obj: Any) -> str:
    for attr in ("Name", "Caption"):
        try:
            name = getattr(obj, attr)
            if name:
                return "<COM object " + str(name) + ">"
        except Exception:  # noqa: BLE001
            continue
    return "<COM object>"


# --------------------------------------------------------------------------
# Applications
# --------------------------------------------------------------------------

APPS: dict[str, str] = {
    "excel": "Excel.Application",
    "word": "Word.Application",
    "powerpoint": "PowerPoint.Application",
    "outlook": "Outlook.Application",
    "onenote": "OneNote.Application",
    "access": "Access.Application",
    "publisher": "Publisher.Application",
    "visio": "Visio.Application",
    "project": "MSProject.Application",
}

# Apps that refuse to be hidden / have no meaningful Visible property.
_NO_VISIBLE = {"outlook", "onenote"}

# OneNote's IDispatch does not resolve its own method names, so late binding
# raises AttributeError on GetHierarchy/FindPages/... It needs a makepy-generated
# early-bound wrapper.
_EARLY_BINDING = {"onenote"}

_app_cache: dict[str, Any] = {}


def normalize_app(name: str) -> str:
    key = (name or "").strip().lower()
    aliases = {
        "xl": "excel",
        "ppt": "powerpoint",
        "pp": "powerpoint",
        "doc": "word",
        "msword": "word",
        "ol": "outlook",
        "on": "onenote",
        "pub": "publisher",
    }
    key = aliases.get(key, key)
    if key not in APPS:
        raise ComToolError(
            "Unknown application %r. Known: %s" % (name, ", ".join(sorted(APPS)))
        )
    return key


def _local_server_path(progid: str) -> str | None:
    """Look up the EXE that implements a ProgID (used to load its typelib)."""
    import winreg

    try:
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, progid + r"\CLSID") as key:
            clsid = winreg.QueryValueEx(key, "")[0]
        with winreg.OpenKey(
            winreg.HKEY_CLASSES_ROOT, r"CLSID\%s\LocalServer32" % clsid
        ) as key:
            return str(winreg.QueryValueEx(key, "")[0]).strip('"')
    except OSError:
        return None


def _early_bound(progid: str) -> Any:
    """Build an early-bound wrapper for apps whose IDispatch cannot resolve names.

    OneNote is the case that forces this: GetIDsOfNames fails for its own
    methods, so late binding raises AttributeError. EnsureDispatch normally
    fixes that, but OneNote also refuses to hand out its typelib through COM, so
    we fall back to reading the typelib resource out of ONENOTE.EXE.
    """
    from win32com.client import gencache, makepy

    gencache.is_readonly = False
    try:
        return win32com.client.gencache.EnsureDispatch(progid)
    except Exception:  # noqa: BLE001 - "can not automate the makepy process"
        pass

    exe = _local_server_path(progid)
    if not exe or not os.path.exists(exe):
        raise ComToolError(
            "Cannot locate the executable behind %s" % progid,
            hint="Is the desktop app installed? The store/UWP build has no COM.",
        )
    last_error: Exception | None = None
    for resource in ("\\3", "\\1", ""):
        try:
            tlb = pythoncom.LoadTypeLib(exe + resource)
        except Exception as exc:  # noqa: BLE001 - try the next resource index
            last_error = exc
            continue
        makepy.GenerateFromTypeLibSpec(tlb)
        attr = tlb.GetLibAttr()
        module = gencache.GetModuleForTypelib(attr[0], attr[1], attr[3], attr[4])
        raw = win32com.client.Dispatch(progid)
        for cls_name in ("IApplication", "Application", "Application2"):
            cls = getattr(module, cls_name, None)
            if cls is not None:
                return cls(raw._oleobj_)
        return raw
    raise ComToolError(
        "Cannot load the type library of %s: %s" % (progid, last_error),
        hint="Run `python -m win32com.client.makepy` against the app once.",
    )


def get_app(name: str, new_instance: bool = False, visible: bool | None = None) -> Any:
    """Attach to a running Office app, or launch one.

    Attaching is the default: the user usually has the document open already and
    a second hidden instance would not see it.
    """
    key = normalize_app(name)
    progid = APPS[key]

    app = _app_cache.get(key)
    if app is not None and not new_instance:
        try:
            _ = app.Name  # liveness probe: raises if the process is gone
            return app
        except Exception:  # noqa: BLE001
            _app_cache.pop(key, None)

    app = None
    if key in _EARLY_BINDING:
        app = _early_bound(progid)
    elif not new_instance:
        try:
            app = win32com.client.GetActiveObject(progid)
        except Exception:  # noqa: BLE001 - nothing running, fall through
            app = None

    if app is None:
        try:
            app = (
                win32com.client.DispatchEx(progid)
                if new_instance
                else win32com.client.Dispatch(progid)
            )
        except Exception as exc:  # noqa: BLE001
            message, _ = describe_com_error(exc)
            raise ComToolError(
                "Cannot start %s: %s" % (progid, message),
                hint="Is %s installed on this machine? Check with m365_status." % key,
            ) from exc

    if visible is not None and key not in _NO_VISIBLE:
        try:
            app.Visible = bool(visible)
        except Exception:  # noqa: BLE001 - some apps refuse while loading
            pass
    if key == "excel":
        # Otherwise a single bad filename hangs the server behind a modal dialog.
        try:
            app.DisplayAlerts = False
            app.AskToUpdateLinks = False
        except Exception:  # noqa: BLE001
            pass
    if key == "word":
        try:
            app.DisplayAlerts = 0  # wdAlertsNone
        except Exception:  # noqa: BLE001
            pass

    _app_cache[key] = app
    return app


def forget_app(name: str) -> None:
    _app_cache.pop(normalize_app(name), None)


def app_installed(name: str) -> bool:
    """Registry check - cheaper and safer than launching the app."""
    import winreg

    progid = APPS[normalize_app(name)]
    try:
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, progid):
            return True
    except OSError:
        return False


def app_running(name: str) -> bool:
    key = normalize_app(name)
    cached = _app_cache.get(key)
    if cached is not None:
        try:
            _ = cached.Name
            return True
        except Exception:  # noqa: BLE001
            _app_cache.pop(key, None)
    try:
        win32com.client.GetActiveObject(APPS[key])
        return True
    except Exception:  # noqa: BLE001
        return False


# --------------------------------------------------------------------------
# Handle registry
# --------------------------------------------------------------------------

_handles: dict[str, Any] = {}
_handle_meta: dict[str, dict[str, Any]] = {}
_counter = itertools.count(1)


def put_handle(prefix: str, obj: Any, **meta: Any) -> str:
    hid = "%s:%d" % (prefix, next(_counter))
    _handles[hid] = obj
    _handle_meta[hid] = meta
    return hid


def get_handle(hid: str) -> Any:
    obj = _handles.get(hid)
    if obj is None:
        raise ComToolError(
            "Unknown handle %r" % hid,
            hint="Handles live for one server run; re-open the file for a new one.",
        )
    return obj


def find_handle_by(prefix: str, key: str, value: Any) -> str | None:
    for hid, meta in _handle_meta.items():
        if hid.startswith(prefix + ":") and meta.get(key) == value:
            return hid
    return None


def drop_handle(hid: str) -> None:
    _handles.pop(hid, None)
    _handle_meta.pop(hid, None)


def handle_meta(hid: str) -> dict[str, Any]:
    return _handle_meta.get(hid, {})


def list_handles(prefix: str | None = None) -> list[dict[str, Any]]:
    out = []
    for hid, obj in list(_handles.items()):
        if prefix and not hid.startswith(prefix + ":"):
            continue
        alive = True
        try:
            _ = obj.Name
        except Exception:  # noqa: BLE001
            alive = False
        row = {"handle": hid, "alive": alive}
        row.update(_handle_meta.get(hid, {}))
        out.append(row)
    return out
