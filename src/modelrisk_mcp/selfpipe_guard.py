"""Keep asyncio's Windows event loop from spinning after the PC sleeps.

On Windows, asyncio's event loop wakes itself through a pair of connected
loopback TCP sockets, its "self-pipe". A thread that hands the loop work (the
stdin reader does, for every message) writes a byte to one end, and the loop
always has a read pending on the other. On some PCs Windows tears idle
loopback connections down across a power or session change, such as sleep and
wake. The pending read then ends in one of two ways, and CPython handles
neither (https://github.com/python/cpython/issues/156333, open):

- End of file. asyncio arms the read again on the dead socket, the read ends at
  once, and so on forever: the idle server uses a full CPU core until it is
  killed. Reported on 13 October 2026 for every modelrisk-mcp and
  modelchoice-mcp process on one PC, after every sleep.
- An error (a reset or aborted read). asyncio reports it and stops reading, so
  the loop no longer wakes for work handed over from threads: the server stops
  answering, and does not exit when the client closes its stdin.

`install()` makes the loop replace the socket pair in both cases and read from
the new one. The server calls it once, before its event loop starts.
"""

from __future__ import annotations

import functools
import logging
import signal
import socket
import sys
import threading
import time
from typing import Any

logger = logging.getLogger(__name__)

# A pair that dies within this many seconds of the last replacement is
# replaced only after this long, so pairs that die as fast as they are made
# cannot spin the loop either.
MIN_REBUILD_INTERVAL = 1.0

_installed = False


def install() -> bool:
    """Patch asyncio's Proactor event loop to replace a dead self-pipe.

    Windows only; calling it again does nothing. Returns whether the patch is
    in place.
    """
    global _installed
    if _installed:
        return True
    if sys.platform != "win32":
        return False

    from asyncio import proactor_events

    loop_cls: Any = proactor_events.BaseProactorEventLoop
    original = getattr(loop_cls, "_loop_self_reading", None)
    if original is None or not hasattr(loop_cls, "_make_self_pipe"):
        logger.warning(
            "asyncio's event loop has changed; the self-pipe guard is not installed"
        )
        return False

    @functools.wraps(original)
    def _loop_self_reading(self: Any, f: Any = None) -> None:
        # `f` is the read that just ended. A read that brought bytes is the
        # normal wake-up; one that ended empty or in an error means the pair is
        # dead, and arming the next read on it would spin or deafen the loop.
        if (
            f is not None
            and f is self._self_reading_future
            and f.done()
            and not f.cancelled()
            and (f.exception() is not None or not f.result())
        ):
            exc = f.exception()
            self._self_reading_future = None
            if not _rebuild(self, "end of file" if exc is None else repr(exc)):
                return
            f = None  # read from the new pair
        original(self, f)

    loop_cls._loop_self_reading = _loop_self_reading
    _installed = True
    return True


def _rebuild(loop: Any, reason: str, *, throttle: bool = True) -> bool:
    """Give `loop` a new self-pipe.

    Returns False when that has to wait. A retry is then scheduled, and the
    loop wakes no later than the retry.
    """
    now = time.monotonic()
    last = getattr(loop, "_selfpipe_rebuilt_at", None)
    if throttle and last is not None and now - last < MIN_REBUILD_INTERVAL:
        loop.call_later(MIN_REBUILD_INTERVAL, _retry, loop, reason)
        return False
    try:
        ssock, csock = socket.socketpair()
    except OSError as exc:
        logger.warning(
            "The event loop's self-pipe was lost (%s) and could not be "
            "replaced (%r); trying again in %.0f s.",
            reason, exc, MIN_REBUILD_INTERVAL,
        )
        loop.call_later(MIN_REBUILD_INTERVAL, _retry, loop, reason)
        return False
    ssock.setblocking(False)
    csock.setblocking(False)

    old_ssock, old_csock = loop._ssock, loop._csock
    if threading.current_thread() is threading.main_thread():
        # Signals (Ctrl+C) wake the loop by writing to the old write end.
        # Point them at the new one, unless something else has taken them.
        old_fd = old_csock.fileno()
        try:
            previous = signal.set_wakeup_fd(csock.fileno())
            if previous != old_fd:
                signal.set_wakeup_fd(previous)
        except (OSError, ValueError):
            pass
    loop._ssock, loop._csock = ssock, csock
    old_ssock.close()
    old_csock.close()

    loop._selfpipe_rebuilt_at = now
    loop._selfpipe_rebuilds = getattr(loop, "_selfpipe_rebuilds", 0) + 1
    logger.warning(
        "The event loop's self-pipe was lost (%s), as Windows can do on sleep "
        "and wake; replaced it (rebuild %d, CPython gh-156333).",
        reason, loop._selfpipe_rebuilds,
    )
    return True


def _retry(loop: Any, reason: str) -> None:
    if loop.is_closed() or loop._self_reading_future is not None:
        return  # closed, or a read is armed again already
    if _rebuild(loop, reason, throttle=False):
        loop._loop_self_reading()
