"""Tests for the new v0.3 MRService.dll bridge.

Without the actual DLL these tests stay focused on:
- Path discovery (env var + standard install paths)
- The activation flow's environment-variable contract
- Error message clarity when activation fails / DLL is missing

A separate integration test (gated on MRService.dll being installed)
exercises the full open-vmrs / read-samples flow.
"""

from __future__ import annotations

from pathlib import Path

import pytest

from modelrisk_mcp.bridge._keymat import decode_bundled_key
from modelrisk_mcp.bridge.mrservice import (
    MrServiceBridge,
    VmrsHandle,
    _call_with_timeout,
    find_latest_vmrs,
    find_mrservice_dll,
)
from modelrisk_mcp.errors import ModelRiskNotLoadedError, SimulationFailedError


class TestFindDll:
    def test_returns_none_when_no_dll_found(
        self, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        monkeypatch.delenv("MRSERVICE_DLL", raising=False)
        monkeypatch.setattr(
            "modelrisk_mcp.bridge.mrservice._DEFAULT_DLL_CANDIDATES",
            (r"C:\nonexistent\MRService.dll",),
        )
        assert find_mrservice_dll() is None

    def test_env_override_wins(
        self, monkeypatch: pytest.MonkeyPatch, tmp_path: Path
    ) -> None:
        fake = tmp_path / "MRService.dll"
        fake.write_bytes(b"")
        monkeypatch.setenv("MRSERVICE_DLL", str(fake))
        assert find_mrservice_dll() == str(fake)


class TestBridgeLifecycle:
    def test_missing_dll_raises_clear_error(
        self, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        monkeypatch.delenv("MRSERVICE_DLL", raising=False)
        monkeypatch.setattr(
            "modelrisk_mcp.bridge.mrservice._DEFAULT_DLL_CANDIDATES",
            (r"C:\nonexistent\MRService.dll",),
        )
        bridge = MrServiceBridge()
        with pytest.raises(ModelRiskNotLoadedError) as exc:
            bridge.ensure_ready()
        msg = str(exc.value)
        assert "MRService.dll not found" in msg
        assert "MRSERVICE_DLL" in msg

    def test_bundled_key_used_when_no_env(
        self, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        """No env override → the bundled activation key is tried via
        `MRLIB_SetOfflineActivationKey`. Tests fake the lib so we
        record the call without needing the real DLL."""
        calls: list[int] = []

        class _FakeLib:
            def MRLIB_SetOfflineActivationKey(self, key: object) -> bool:  # noqa: N802
                calls.append(int(getattr(key, "value", key)))  # type: ignore[arg-type]
                return True

        bridge = MrServiceBridge()
        bridge._lib = _FakeLib()  # type: ignore[assignment]
        monkeypatch.delenv("MRSERVICE_ACTIVATION_KEY", raising=False)
        monkeypatch.delenv("MRSERVICE_ACTIVATION_KEY1", raising=False)
        monkeypatch.delenv("MRSERVICE_ACTIVATION_KEY2", raising=False)
        monkeypatch.delenv("MRSERVICE_DISABLE_BUNDLED_KEY", raising=False)
        bridge._activate()
        assert bridge._activated
        assert len(calls) == 1
        assert calls[0] > 0  # the bundled key, whatever its current value

    def test_env_override_beats_bundled(
        self, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        """An explicit MRSERVICE_ACTIVATION_KEY env var is preferred
        over the bundled fallback."""
        calls: list[int] = []

        class _FakeLib:
            def MRLIB_SetOfflineActivationKey(self, key: object) -> bool:  # noqa: N802
                calls.append(int(getattr(key, "value", key)))  # type: ignore[arg-type]
                return True

        bridge = MrServiceBridge()
        bridge._lib = _FakeLib()  # type: ignore[assignment]
        monkeypatch.setenv("MRSERVICE_ACTIVATION_KEY", "9999")
        bridge._activate()
        assert calls == [9999]

    def test_bundled_key_can_be_disabled(
        self, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        """MRSERVICE_DISABLE_BUNDLED_KEY=1 suppresses the fallback,
        restoring the original 'must configure env var' behaviour."""
        bridge = MrServiceBridge()
        bridge._lib = object()  # type: ignore[assignment]
        monkeypatch.delenv("MRSERVICE_ACTIVATION_KEY", raising=False)
        monkeypatch.delenv("MRSERVICE_ACTIVATION_KEY1", raising=False)
        monkeypatch.delenv("MRSERVICE_ACTIVATION_KEY2", raising=False)
        monkeypatch.setenv("MRSERVICE_DISABLE_BUNDLED_KEY", "1")
        with pytest.raises(SimulationFailedError) as exc:
            bridge._activate()
        msg = str(exc.value)
        assert "MRSERVICE_ACTIVATION_KEY" in msg
        assert "bundled" in msg.lower()

    def test_non_integer_key_rejected(
        self, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        bridge = MrServiceBridge()
        bridge._lib = object()  # type: ignore[assignment]
        monkeypatch.setenv("MRSERVICE_ACTIVATION_KEY", "not-a-number")
        with pytest.raises(SimulationFailedError, match="not an integer"):
            bridge._activate()


class TestKeyObfuscation:
    """The bundled activation key must never appear as a plain integer
    or as its decimal string in any shipped source file. These tests
    guard against regressions — if someone accidentally inlines the
    literal again, CI will fail."""

    def test_decode_returns_positive_int64(self) -> None:
        key = decode_bundled_key()
        assert isinstance(key, int)
        assert 0 < key < (1 << 63)

    def test_decode_is_deterministic(self) -> None:
        assert decode_bundled_key() == decode_bundled_key()

    def test_no_literal_in_package_sources(self) -> None:
        """Recursively scan every .py file in the package for the
        decoded key's decimal representation. Skip the test module
        itself (this assertion would otherwise be self-defeating)."""
        from pathlib import Path

        import modelrisk_mcp

        key_str = str(decode_bundled_key())
        pkg_root = Path(modelrisk_mcp.__file__).parent
        offenders: list[str] = []
        for py in pkg_root.rglob("*.py"):
            text = py.read_text(encoding="utf-8", errors="ignore")
            if key_str in text:
                offenders.append(str(py.relative_to(pkg_root)))
        assert not offenders, (
            f"Plain activation key found in shipped sources: {offenders}. "
            "Run scripts/encode_activation_key.py and replace the literal."
        )


class TestCallWithTimeout:
    """The timeout wrapper guards individual MRService.dll calls from
    pathological-input hangs (most notably MRLIB_GetModelVarID stuck
    on names containing `?`, `(`, or `)`)."""

    def test_returns_result_when_call_completes(self) -> None:
        assert _call_with_timeout(lambda: 42, timeout=1.0, label="fast") == 42

    def test_propagates_exception_from_call(self) -> None:
        def boom() -> int:
            raise RuntimeError("inner failure")

        with pytest.raises(RuntimeError, match="inner failure"):
            _call_with_timeout(boom, timeout=1.0, label="boom")

    def test_raises_simulation_failed_on_timeout(self) -> None:
        import time

        def hang() -> int:
            time.sleep(10)
            return 0  # pragma: no cover - unreached

        with pytest.raises(SimulationFailedError) as exc:
            _call_with_timeout(hang, timeout=0.1, label="hang")
        msg = str(exc.value)
        # The message must be actionable — name the likely cause and
        # the workaround (rename the variable).
        assert "did not return" in msg
        assert "rename" in msg.lower()
        assert "?" in msg or "(" in msg


class TestVmrsHandleLookupVarId:
    """`VmrsHandle.lookup_var_id` is the timeout-protected wrapper around
    MRLIB_GetModelVarID. The tests use a fake lib so we don't need the
    real DLL — they verify the success path, the not-found path, and
    the timeout path (the bug #16 fix)."""

    def _make_handle(self, fake_lib: object) -> VmrsHandle:
        return VmrsHandle(
            fake_lib,  # type: ignore[arg-type]
            model_ptr=12345,
            path="fake.vmrs",
        )

    def test_returns_var_id_on_success(self) -> None:
        class _FakeLib:
            def MRLIB_GetModelVarID(  # noqa: N802
                self, _model_ptr: object, name: str, var_id_ptr: object
            ) -> bool:
                # ctypes byref(c_int) — assign .value via _obj_
                var_id_ptr._obj.value = 7  # type: ignore[attr-defined]
                return True

        lib = _FakeLib()
        lib._modelrisk_mcp_var_id_configured = True  # type: ignore[attr-defined]
        handle = self._make_handle(lib)
        assert handle.lookup_var_id("WidgetCost", timeout=1.0) == 7

    def test_returns_none_when_lib_reports_failure(self) -> None:
        class _FakeLib:
            def MRLIB_GetModelVarID(  # noqa: N802
                self, _model_ptr: object, _name: str, _var_id_ptr: object
            ) -> bool:
                return False

        lib = _FakeLib()
        lib._modelrisk_mcp_var_id_configured = True  # type: ignore[attr-defined]
        handle = self._make_handle(lib)
        assert handle.lookup_var_id("Missing", timeout=1.0) is None

    def test_returns_none_when_var_id_negative(self) -> None:
        class _FakeLib:
            def MRLIB_GetModelVarID(  # noqa: N802
                self, _model_ptr: object, _name: str, var_id_ptr: object
            ) -> bool:
                # SDK occasionally returns True but leaves var_id at -1.
                # Treat that as not-found, same as ok=False.
                var_id_ptr._obj.value = -1  # type: ignore[attr-defined]
                return True

        lib = _FakeLib()
        lib._modelrisk_mcp_var_id_configured = True  # type: ignore[attr-defined]
        handle = self._make_handle(lib)
        assert handle.lookup_var_id("EdgeCase", timeout=1.0) is None

    def test_raises_on_pathological_name_timeout(self) -> None:
        """Regression for bug #16: names with `?`, `(`, or `)` caused
        MRService.dll to hang for the full 4-minute Claude Desktop
        timeout. With the wrapper, lookup_var_id raises a clear
        SimulationFailedError within the configured budget."""
        import time

        class _HangingLib:
            def MRLIB_GetModelVarID(  # noqa: N802
                self, _model_ptr: object, _name: str, _var_id_ptr: object
            ) -> bool:
                time.sleep(10)
                return False  # pragma: no cover

        lib = _HangingLib()
        lib._modelrisk_mcp_var_id_configured = True  # type: ignore[attr-defined]
        handle = self._make_handle(lib)
        with pytest.raises(SimulationFailedError) as exc:
            handle.lookup_var_id("Conservatives get in? (1=yes)", timeout=0.1)
        msg = str(exc.value)
        assert "GetModelVarID" in msg
        assert "Conservatives get in? (1=yes)" in msg
        assert "rename" in msg.lower()

    def test_timeout_env_override(
        self, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        """MRSERVICE_VARID_TIMEOUT_S overrides the default budget for
        environments where the SDK is unusually slow."""
        import time

        class _SlowLib:
            def __init__(self) -> None:
                self.delay = 0.3

            def MRLIB_GetModelVarID(  # noqa: N802
                self, _model_ptr: object, _name: str, var_id_ptr: object
            ) -> bool:
                time.sleep(self.delay)
                var_id_ptr._obj.value = 3  # type: ignore[attr-defined]
                return True

        lib = _SlowLib()
        lib._modelrisk_mcp_var_id_configured = True  # type: ignore[attr-defined]
        handle = self._make_handle(lib)

        # Default budget (8 s) would let this through, but the env
        # override clamps the budget below the delay → expect timeout.
        monkeypatch.setenv("MRSERVICE_VARID_TIMEOUT_S", "0.05")
        with pytest.raises(SimulationFailedError, match="did not return"):
            handle.lookup_var_id("Slow")


class TestVmrsDiscovery:
    def test_returns_sibling_vmrs(self, tmp_path: Path) -> None:
        workbook = tmp_path / "model.xlsx"
        workbook.write_text("")
        vmrs = tmp_path / "model_1.vmrs"
        vmrs.write_bytes(b"")
        assert find_latest_vmrs(workbook) == str(vmrs)

    def test_returns_most_recent_when_multiple(self, tmp_path: Path) -> None:
        import time

        workbook = tmp_path / "model.xlsx"
        workbook.write_text("")
        old = tmp_path / "model_1.vmrs"
        old.write_bytes(b"")
        time.sleep(0.05)
        new = tmp_path / "model_2.vmrs"
        new.write_bytes(b"")
        assert find_latest_vmrs(workbook) == str(new)

    def test_returns_none_when_no_vmrs(self, tmp_path: Path) -> None:
        workbook = tmp_path / "model.xlsx"
        workbook.write_text("")
        assert find_latest_vmrs(workbook) is None

    def test_returns_none_when_workbook_missing(self, tmp_path: Path) -> None:
        assert find_latest_vmrs(tmp_path / "nope.xlsx") is None
