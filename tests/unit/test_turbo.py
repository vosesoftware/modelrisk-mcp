"""Unit tests for bridge/turbo.py: the Turbo engine's own compatibility
check, run in a child process, and the watch that answers the message boxes
ModelRisk's Turbo command shows while Application.Run waits on it."""

from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path
from typing import Any

import pytest

from modelrisk_mcp.bridge import turbo
from modelrisk_mcp.bridge.turbo import (
    IDABORT,
    IDCANCEL,
    IDIGNORE,
    IDNO,
    IDOK,
    IDRETRY,
    IDYES,
    DialogWatch,
    TurboCheck,
)
from tests.unit._message_boxes import FakeWindows

_COPY = r"C:\Temp\modelrisk-mcp-turbo-1\turbo-check.xlsx"


def _child(
    monkeypatch: pytest.MonkeyPatch,
    *,
    stdout: str = "",
    returncode: int = 0,
    raises: Exception | None = None,
) -> dict[str, Any]:
    """Stand in for the `turbo-check` child process; returns what it was
    started with."""
    seen: dict[str, Any] = {}

    def run(command: list[str], **kwargs: Any) -> subprocess.CompletedProcess[str]:
        seen["command"] = command
        seen["kwargs"] = kwargs
        if raises is not None:
            raise raises
        return subprocess.CompletedProcess(command, returncode, stdout, "")

    monkeypatch.setattr(turbo.subprocess, "run", run)
    return seen


class TestCheckWorkbook:
    def test_the_childs_report_is_read(self, monkeypatch: pytest.MonkeyPatch) -> None:
        seen = _child(
            monkeypatch, stdout='loading\n{"unsupported": ["VOSETIMEGBMVR"]}\n'
        )

        check = turbo.check_workbook(_COPY)

        assert check == TurboCheck(unsupported=("VOSETIMEGBMVR",))
        assert seen["command"] == [
            sys.executable, "-m", "modelrisk_mcp", "turbo-check", _COPY,
        ]
        # Never the server's own stdin: that is the MCP stream.
        assert seen["kwargs"]["stdin"] is subprocess.DEVNULL

    def test_the_standalone_exe_runs_itself(
        self, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        monkeypatch.setattr(sys, "frozen", True, raising=False)
        seen = _child(monkeypatch, stdout='{"unsupported": []}')

        assert turbo.check_workbook(_COPY) == TurboCheck()
        assert seen["command"] == [sys.executable, "turbo-check", _COPY]

    def test_the_engines_error_is_the_problem(
        self, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        error = "the engine could not read the workbook (status 1: not a zip)"
        _child(monkeypatch, stdout=json.dumps({"error": error}), returncode=1)

        assert turbo.check_workbook(_COPY) == TurboCheck(problem=error)

    def test_a_crash_is_reported_with_its_exit_code(
        self, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        _child(monkeypatch, returncode=0xC0000005)

        check = turbo.check_workbook(_COPY)

        assert check.problem == "it ended with exit code 0xC0000005"

    def test_a_check_that_hangs_is_given_up(
        self, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        _child(monkeypatch, raises=subprocess.TimeoutExpired("turbo-check", 120))

        check = turbo.check_workbook(_COPY, timeout=120)

        assert check.problem == "it did not finish within 120 s"


class TestCheckCommand:
    def test_the_names_are_printed(
        self, monkeypatch: pytest.MonkeyPatch, capsys: pytest.CaptureFixture[str]
    ) -> None:
        monkeypatch.setattr(turbo, "find_engine_dll", lambda: Path("engine.dll"))
        monkeypatch.setattr(
            turbo, "unsupported_functions", lambda path, dll: ["VOSETIMEGBMVR"]
        )

        assert turbo.run_check_command([_COPY]) == 0
        assert json.loads(capsys.readouterr().out) == {"unsupported": ["VOSETIMEGBMVR"]}

    def test_no_engine_installed(
        self, monkeypatch: pytest.MonkeyPatch, capsys: pytest.CaptureFixture[str]
    ) -> None:
        monkeypatch.setattr(turbo, "find_engine_dll", lambda: None)

        assert turbo.run_check_command([_COPY]) == 1
        assert "is not installed" in json.loads(capsys.readouterr().out)["error"]

    def test_an_engine_failure_is_printed(
        self, monkeypatch: pytest.MonkeyPatch, capsys: pytest.CaptureFixture[str]
    ) -> None:
        def fail(path: str, dll: Path) -> list[str]:
            raise RuntimeError("the engine could not read the workbook")

        monkeypatch.setattr(turbo, "find_engine_dll", lambda: Path("engine.dll"))
        monkeypatch.setattr(turbo, "unsupported_functions", fail)

        assert turbo.run_check_command([_COPY]) == 1
        assert json.loads(capsys.readouterr().out) == {
            "error": "the engine could not read the workbook"
        }

    def test_the_cli_runs_it(self, capsys: pytest.CaptureFixture[str]) -> None:
        from modelrisk_mcp.__main__ import main

        with pytest.raises(SystemExit) as exc:
            main(["turbo-check"])

        assert exc.value.code == 2
        assert "usage" in json.loads(capsys.readouterr().out)["error"]


class TestFindEngineDll:
    def test_the_override_wins(
        self, monkeypatch: pytest.MonkeyPatch, tmp_path: Path
    ) -> None:
        dll = tmp_path / "mrengine_native.dll"
        dll.write_bytes(b"MZ")
        monkeypatch.setenv("MRENGINE_NATIVE_DLL", str(dll))

        assert turbo.find_engine_dll() == dll

    def test_a_missing_override_is_skipped(
        self, monkeypatch: pytest.MonkeyPatch, tmp_path: Path
    ) -> None:
        monkeypatch.setenv("MRENGINE_NATIVE_DLL", str(tmp_path / "gone.dll"))
        monkeypatch.setattr(turbo, "_DEFAULT_DLL_CANDIDATES", ())

        assert turbo.find_engine_dll() is None


class TestDialogWatch:
    @pytest.mark.parametrize(
        ("buttons", "answer"),
        [
            ({IDYES, IDNO}, IDNO),
            ({IDYES, IDNO, IDCANCEL}, IDNO),
            ({IDOK, IDCANCEL}, IDCANCEL),
            ({IDABORT, IDRETRY, IDIGNORE}, IDABORT),
            # A Windows MB_OK box: its OK button has the id IDCANCEL.
            ({IDCANCEL}, IDCANCEL),
            ({IDOK}, IDOK),
        ],
    )
    def test_each_box_is_declined(self, buttons: set[int], answer: int) -> None:
        windows = FakeWindows()

        with DialogWatch(4242, windows=windows, interval=0.01) as watch:
            pressed = windows.show("Go on?", buttons)

        assert pressed == answer
        assert [box.text for box in watch.answered] == ["Go on?"]
        assert windows.pids == {4242}

    def test_boxes_are_kept_in_order(self) -> None:
        windows = FakeWindows()
        offer = (
            "Check that the Turbo engine matches the classic ModelRisk engine "
            "for this workbook before running it?"
        )
        failure = "The simulation could not be completed:\nFAIL|no result"

        with DialogWatch(4242, windows=windows, interval=0.01) as watch:
            windows.show(offer, {IDYES, IDNO})
            windows.show(failure, {IDCANCEL})

        assert [(box.is_question, box.text) for box in watch.answered] == [
            (True, offer),
            (False, failure),
        ]

    def test_a_box_still_up_is_pressed_again(
        self, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        monkeypatch.setattr(turbo, "_PRESS_AGAIN_AFTER", 0.05)
        windows = FakeWindows(ignore_presses=1)

        with DialogWatch(4242, windows=windows, interval=0.01) as watch:
            pressed = windows.show("Stuck?", {IDOK})

        assert pressed == IDOK
        assert len(windows.presses) == 2
        assert len(watch.answered) == 1


__all__ = [
    "TestCheckCommand",
    "TestCheckWorkbook",
    "TestDialogWatch",
    "TestFindEngineDll",
]
