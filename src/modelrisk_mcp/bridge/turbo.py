"""ModelRisk's Turbo engine: whether it can run a workbook, and a watch that
answers the message boxes its command shows.

Turbo is ModelRisk's fast simulation engine, MREngineRun.exe. Its command,
`VoseStartFastSimulation`, simulates the active workbook
(ModelRiskCloude/MREngine.cpp:1024): the add-in saves a copy,
`vsmre_<book>.xlsx`, runs the engine on it, moves the engine's run file into
its simulation folder as `f<hWndExcel>vsmre_<stem>.dmr`
(SimulationObj.cpp:14690) and shows that run in the Results Viewer. Two
things stand in the way of running it unattended:

- Turbo cannot evaluate every function; the time-series family is the
  largest gap. An output that depends on such a function comes back NaN in
  every iteration, with no message. The engine's own check,
  `mr_model_unsupported_functions` in mrengine_native.dll, names those
  functions and needs no activation. It runs in a child process
  (`modelrisk-mcp turbo-check`), so a crash in the native engine cannot take
  the server down.
- The command talks through system-modal message boxes, which keep
  Application.Run waiting for a click. Before a workbook's first verified
  Turbo run it asks whether to check Turbo against the classic engine, and a
  failed run ends with an OK box. The question comes every time here,
  because VoseSetSimulOptions12 resets the flag that silences it.
  `DialogWatch` answers them (No to a question, OK to a message) and keeps
  their text.
"""

from __future__ import annotations

import json
import os
import subprocess
import sys
import threading
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Protocol

DLL_NAME = "mrengine_native.dll"

# The engine reads OOXML packages only (MREngine.cpp:210).
TURBO_FORMATS = (".xlsx", ".xlsm")

_DEFAULT_DLL_CANDIDATES: tuple[str, ...] = (
    rf"C:\Program Files\Vose Software\ModelRisk\{DLL_NAME}",
    rf"C:\Program Files (x86)\Vose Software\ModelRisk\{DLL_NAME}",
)


# ---------------------------------------------------------------------------
# Can Turbo run this workbook?
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class TurboCheck:
    """The engine's verdict on a workbook. `problem` is set when the check
    itself could not run; `unsupported` is then empty and means nothing."""

    unsupported: tuple[str, ...] = ()
    problem: str | None = None


def find_engine_dll() -> Path | None:
    """The installed ModelRisk's mrengine_native.dll. `MRENGINE_NATIVE_DLL`
    overrides; a path that is set but missing is skipped."""
    override = os.environ.get("MRENGINE_NATIVE_DLL")
    for candidate in ((override,) if override else ()) + _DEFAULT_DLL_CANDIDATES:
        if Path(candidate).is_file():
            return Path(candidate)
    return None


def unsupported_functions(workbook_path: str, dll_path: Path) -> list[str]:
    """The distinct functions Turbo cannot evaluate among the cells the
    workbook's outputs are recomputed from, as the engine names them
    (DT/mrengine engine_capi.cpp, #unsupported-scan). Loads the native engine
    into THIS process: call it from the `turbo-check` child, not the server.
    The workbook must be a saved copy; the engine cannot open a file Excel
    has open."""
    import ctypes

    os.add_dll_directory(str(dll_path.parent))
    dll = ctypes.CDLL(str(dll_path))
    dll.mr_last_error.restype = ctypes.c_char_p
    dll.mr_model_load_file.argtypes = [
        ctypes.c_char_p, ctypes.POINTER(ctypes.c_void_p),
    ]
    dll.mr_model_load_file.restype = ctypes.c_int
    dll.mr_model_unsupported_functions.argtypes = [
        ctypes.c_void_p, ctypes.c_char_p, ctypes.c_int32,
    ]
    dll.mr_model_unsupported_functions.restype = ctypes.c_int32
    dll.mr_model_free.argtypes = [ctypes.c_void_p]
    dll.mr_model_free.restype = None

    model = ctypes.c_void_p()
    status = dll.mr_model_load_file(
        workbook_path.encode("utf-8"), ctypes.byref(model)
    )
    if status != 0:
        detail = (dll.mr_last_error() or b"").decode(errors="replace")
        raise RuntimeError(
            f"the engine could not read the workbook (status {status}: {detail})"
        )
    try:
        buf = ctypes.create_string_buffer(1 << 16)
        dll.mr_model_unsupported_functions(model, buf, len(buf))
        names = buf.value.decode(errors="replace")
    finally:
        dll.mr_model_free(model)
    return [name for name in (n.strip() for n in names.split(",")) if name]


def check_workbook(workbook_path: str, *, timeout: float = 120.0) -> TurboCheck:
    """Ask the installed engine, in a child process, which functions in the
    saved workbook at `workbook_path` it cannot evaluate."""
    if getattr(sys, "frozen", False):  # the standalone exe is the CLI itself
        command = [sys.executable, "turbo-check", workbook_path]
    else:
        command = [sys.executable, "-m", "modelrisk_mcp", "turbo-check", workbook_path]
    try:
        done = subprocess.run(
            command,
            stdin=subprocess.DEVNULL,  # never the server's own MCP stream
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=timeout,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
    except subprocess.TimeoutExpired:
        return TurboCheck(problem=f"it did not finish within {timeout:.0f} s")
    except OSError as exc:
        return TurboCheck(problem=f"it could not be started ({exc})")
    report = _last_json_line(done.stdout)
    if isinstance(report.get("unsupported"), list):
        return TurboCheck(unsupported=tuple(str(n) for n in report["unsupported"]))
    if report.get("error"):
        return TurboCheck(problem=str(report["error"]))
    code = done.returncode
    shown = str(code) if 0 <= code <= 255 else f"0x{code & 0xFFFFFFFF:08X}"
    return TurboCheck(problem=f"it ended with exit code {shown}")


def _last_json_line(stdout: str) -> dict[str, Any]:
    for line in reversed(stdout.strip().splitlines()):
        try:
            report = json.loads(line)
        except json.JSONDecodeError:
            continue
        if isinstance(report, dict):
            return report
    return {}


def run_check_command(argv: list[str]) -> int:
    """`modelrisk-mcp turbo-check <workbook>`: one JSON line on stdout,
    {"unsupported": [...]} or {"error": "..."}. Hidden; run_simulation's
    engine="turbo" runs it on a saved copy of the workbook."""
    if len(argv) != 1:
        print(json.dumps({"error": "usage: modelrisk-mcp turbo-check <workbook>"}))
        return 2
    dll = find_engine_dll()
    if dll is None:
        print(json.dumps(
            {"error": f"ModelRisk's Turbo engine ({DLL_NAME}) is not installed"}
        ))
        return 1
    try:
        names = unsupported_functions(argv[0], dll)
    except Exception as exc:
        print(json.dumps({"error": str(exc)}))
        return 1
    print(json.dumps({"unsupported": names}))
    return 0


# ---------------------------------------------------------------------------
# Answering the Turbo command's message boxes
# ---------------------------------------------------------------------------

IDOK, IDCANCEL, IDABORT, IDRETRY, IDIGNORE, IDYES, IDNO = 1, 2, 3, 4, 5, 6, 7
_BUTTONS = (IDOK, IDCANCEL, IDABORT, IDRETRY, IDIGNORE, IDYES, IDNO)
# The button that declines, best first. A message's only button, OK, has
# the id IDCANCEL in a Windows MB_OK box (seen on Windows 11).
_DECLINE = (IDNO, IDCANCEL, IDABORT, IDOK)
# The control holding the text of a Windows message box.
_MESSAGE_TEXT_ID = 0xFFFF
# A box still up this long after its button was pressed is pressed again.
_PRESS_AGAIN_AFTER = 2.0


@dataclass(frozen=True)
class MessageBox:
    hwnd: int
    title: str
    text: str
    buttons: frozenset[int]

    @property
    def is_question(self) -> bool:
        return len(self.buttons) > 1

    def decline(self) -> int | None:
        return next((b for b in _DECLINE if b in self.buttons), None)


class Windows(Protocol):
    def message_boxes(self, pid: int) -> list[MessageBox]: ...

    def press(self, hwnd: int, button: int) -> None: ...


class _Win32Windows:
    """The visible message boxes of a process, through pywin32. A message
    box is a dialog (#32770) with the message-text control."""

    def message_boxes(self, pid: int) -> list[MessageBox]:
        import win32gui
        import win32process

        hwnds: list[int] = []

        def collect(hwnd: int, _: Any) -> bool:
            hwnds.append(hwnd)
            return True

        win32gui.EnumWindows(collect, None)
        boxes: list[MessageBox] = []
        for hwnd in hwnds:
            try:
                if win32process.GetWindowThreadProcessId(hwnd)[1] != pid:
                    continue
                if win32gui.GetClassName(hwnd) != "#32770":
                    continue
                if not win32gui.IsWindowVisible(hwnd):
                    continue
                text_control = _dialog_item(hwnd, _MESSAGE_TEXT_ID)
                if not text_control:
                    continue
                boxes.append(MessageBox(
                    hwnd=hwnd,
                    title=win32gui.GetWindowText(hwnd),
                    text=win32gui.GetWindowText(text_control),
                    buttons=frozenset(b for b in _BUTTONS if _dialog_item(hwnd, b)),
                ))
            except Exception:
                continue  # the window closed while it was being read
        return boxes

    def press(self, hwnd: int, button: int) -> None:
        import win32api
        import win32con

        win32api.PostMessage(hwnd, win32con.WM_COMMAND, button, 0)


def _dialog_item(hwnd: int, item_id: int) -> int:
    import win32gui

    try:
        return int(win32gui.GetDlgItem(hwnd, item_id) or 0)
    except Exception:  # pywin32 raises when there is no such control
        return 0


class DialogWatch:
    """While a ModelRisk command runs in Excel process `pid`, answer every
    message box that process shows with the button that declines: No, else
    Cancel, else a message's only button. `answered` lists the boxes in the
    order they came."""

    def __init__(
        self, pid: int, *, windows: Windows | None = None, interval: float = 0.2
    ) -> None:
        self.pid = pid
        self.answered: list[MessageBox] = []
        self._windows: Windows = windows or _Win32Windows()
        self._interval = interval
        self._stop = threading.Event()
        self._thread = threading.Thread(
            target=self._watch, name="modelrisk-dialog-watch", daemon=True
        )

    def __enter__(self) -> DialogWatch:
        self._thread.start()
        return self

    def __exit__(self, *exc: object) -> None:
        self._stop.set()
        self._thread.join(timeout=5)

    def _watch(self) -> None:
        pressed: dict[int, float] = {}
        while not self._stop.wait(self._interval):
            try:
                boxes = self._windows.message_boxes(self.pid)
            except Exception:
                continue
            now = time.monotonic()
            present = {box.hwnd for box in boxes}
            for hwnd in [h for h in pressed if h not in present]:
                del pressed[hwnd]  # closed; a new box may get its handle
            for box in boxes:
                button = box.decline()
                last = pressed.get(box.hwnd)
                if button is None or (last is not None and now - last < _PRESS_AGAIN_AFTER):
                    continue
                if last is None:
                    self.answered.append(box)
                pressed[box.hwnd] = now
                try:
                    self._windows.press(box.hwnd, button)
                except Exception:
                    pass


__all__ = [
    "DLL_NAME",
    "TURBO_FORMATS",
    "DialogWatch",
    "MessageBox",
    "TurboCheck",
    "check_workbook",
    "find_engine_dll",
    "run_check_command",
    "unsupported_functions",
]
