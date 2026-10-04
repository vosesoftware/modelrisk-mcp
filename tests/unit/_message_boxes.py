"""A fake of the Windows message boxes the Turbo dialog watch answers.

`show` raises a box and blocks, as MessageBox does, until the watch presses
one of its buttons, then returns that button. `ignore_presses` makes the
first presses do nothing, as when a click is lost.
"""

from __future__ import annotations

import threading
from collections.abc import Iterable

from modelrisk_mcp.bridge.turbo import MessageBox


class FakeWindows:
    def __init__(self, *, ignore_presses: int = 0) -> None:
        self._lock = threading.Lock()
        self._open: dict[int, tuple[MessageBox, threading.Event]] = {}
        self._answers: dict[int, int] = {}
        self._next_hwnd = 0x1000
        self._ignore = ignore_presses
        self.pids: set[int] = set()
        self.presses: list[tuple[int, int]] = []

    def show(
        self, text: str, buttons: Iterable[int], title: str = "ModelRisk"
    ) -> int:
        done = threading.Event()
        with self._lock:
            hwnd = self._next_hwnd
            self._next_hwnd += 1
            box = MessageBox(hwnd, title, text, frozenset(buttons))
            self._open[hwnd] = (box, done)
        if not done.wait(timeout=5):
            with self._lock:
                self._open.pop(hwnd, None)
            raise AssertionError(f"nobody answered the message box {text!r}")
        return self._answers[hwnd]

    def message_boxes(self, pid: int) -> list[MessageBox]:
        with self._lock:
            self.pids.add(pid)
            return [box for box, _ in self._open.values()]

    def press(self, hwnd: int, button: int) -> None:
        with self._lock:
            self.presses.append((hwnd, button))
            if self._ignore > 0:
                self._ignore -= 1
                return
            _, done = self._open.pop(hwnd)
            self._answers[hwnd] = button
        done.set()
