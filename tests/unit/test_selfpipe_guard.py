"""The self-pipe guard keeps an idle event loop idle and awake when Windows
tears its loopback socket pair down (CPython gh-156333).

Each test runs a fresh Proactor loop, breaks the pair the way Windows does (a
clean end of file, or a read that ends in an error), and then checks the two
things that go wrong without the guard: the loop must not use the CPU while
idle, and a callback handed over from another thread must still wake it.
"""

from __future__ import annotations

import asyncio
import contextlib
import itertools
import signal
import socket
import sys
import threading
import time
from collections.abc import Callable
from typing import Any

import pytest

from modelrisk_mcp import selfpipe_guard

windows_only = pytest.mark.skipif(
    sys.platform != "win32", reason="the Proactor event loop is Windows-only"
)


def _end_of_file(loop: Any) -> None:
    # What a torn-down pair looks like from the reading end.
    loop._csock.shutdown(socket.SHUT_WR)


def _aborted_read(loop: Any) -> None:
    # The pending read ends in an error (ERROR_OPERATION_ABORTED).
    loop._self_reading_future._ov.cancel()


async def _idle_cpu(seconds: float = 1.0) -> float:
    """CPU seconds this process uses while the loop has nothing to do."""
    start = time.process_time()
    await asyncio.sleep(seconds)
    return time.process_time() - start


async def _wakeup_delay(loop: Any) -> float:
    """Seconds between another thread handing the loop a callback and the
    callback running. Nothing else is due for 2 s, so only the self-pipe can
    wake the loop sooner."""
    ran: asyncio.Future[float] = loop.create_future()
    sent: list[float] = []

    def hand_over() -> None:
        sent.append(time.monotonic())
        loop.call_soon_threadsafe(lambda: ran.set_result(time.monotonic()))

    threading.Timer(0.05, hand_over).start()
    await asyncio.wait({ran}, timeout=2.0)
    return ran.result() - sent[0]


def _run(scenario: Callable[[Any], Any]) -> Any:
    assert selfpipe_guard.install()

    async def main() -> Any:
        loop = asyncio.get_running_loop()
        await asyncio.sleep(0.05)  # the first read is armed
        return await scenario(loop)

    return asyncio.run(main())


@windows_only
@pytest.mark.parametrize("break_pipe", [_end_of_file, _aborted_read])
def test_a_dead_self_pipe_is_replaced(break_pipe: Callable[[Any], None]) -> None:
    async def scenario(loop: Any) -> tuple[bool, float, float]:
        old = loop._ssock
        break_pipe(loop)
        await asyncio.sleep(0.1)
        return loop._ssock is not old, await _idle_cpu(), await _wakeup_delay(loop)

    replaced, cpu, delay = _run(scenario)
    assert replaced
    assert cpu < 0.3, f"the idle loop used {cpu:.2f} s of CPU in 1 s"
    assert delay < 0.5, f"a thread's callback waited {delay:.2f} s"


@windows_only
def test_signals_follow_the_new_pipe() -> None:
    async def scenario(loop: Any) -> tuple[int, int, int]:
        old_fd = loop._csock.fileno()
        _end_of_file(loop)
        await asyncio.sleep(0.1)
        fd = signal.set_wakeup_fd(-1)
        with contextlib.suppress(OSError, ValueError):
            signal.set_wakeup_fd(fd)  # as it was, if that socket is still open
        return old_fd, fd, loop._csock.fileno()

    old_fd, wakeup_fd, csock_fd = _run(scenario)
    assert csock_fd != old_fd
    assert wakeup_fd == csock_fd


@windows_only
def test_pairs_that_keep_dying_are_replaced_once_a_second() -> None:
    async def scenario(loop: Any) -> tuple[list[float], float]:
        start = time.monotonic()
        cpu = time.process_time()
        rebuilt_at: list[float] = []
        while time.monotonic() - start < 2.5:
            old = loop._ssock
            if loop._self_reading_future is not None:
                _end_of_file(loop)
            while loop._ssock is old and time.monotonic() - start < 2.5:
                await asyncio.sleep(0.01)
            if loop._ssock is not old:
                rebuilt_at.append(time.monotonic() - start)
        return rebuilt_at, time.process_time() - cpu

    rebuilt_at, cpu = _run(scenario)
    assert 2 <= len(rebuilt_at) <= 3, rebuilt_at
    gaps = [b - a for a, b in itertools.pairwise(rebuilt_at)]
    assert all(g >= 0.9 for g in gaps), rebuilt_at
    assert cpu < 0.5, f"the loop used {cpu:.2f} s of CPU in 2.5 s"


def test_the_server_installs_the_guard_before_it_serves(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from modelrisk_mcp import __main__ as cli

    calls: list[str] = []
    monkeypatch.setattr(cli.selfpipe_guard, "install", lambda: calls.append("guard"))
    monkeypatch.setattr(cli, "_run_stdio", lambda: calls.append("serve"))
    cli.main([])
    assert calls == ["guard", "serve"]
