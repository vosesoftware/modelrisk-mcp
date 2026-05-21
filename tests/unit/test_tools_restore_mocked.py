"""MCP-wrapper tests for `tools/restore.py`.

The bridge-level restore logic (audit-log replay, writer-mutex acquire)
is covered by the in-repo audit-log tests in test_safety.py and the
restore round-trip in test_tools_building_mocked.py. What this file
guards is the MCP tool's input handling: ISO-8601 parsing, CellRef
construction, and pass-through to bridge.restore_cell.
"""

from __future__ import annotations

from collections.abc import Iterator
from datetime import UTC, datetime
from unittest.mock import MagicMock

import pytest

from modelrisk_mcp.schemas.distributions import InsertResult
from modelrisk_mcp.schemas.workbook import CellRef
from modelrisk_mcp.tools import reading, restore


@pytest.fixture
def mock_bridge() -> Iterator[MagicMock]:
    bridge = MagicMock()
    bridge.restore_cell.return_value = InsertResult(
        cell=CellRef(workbook="b.xlsx", sheet="S1", cell="A1"),
        formula="=42",
        written=True,
        previous_formula="=99",
    )
    reading.set_bridge_for_testing(bridge)  # type: ignore[arg-type]
    yield bridge
    reading.set_bridge_for_testing(None)


class TestRestoreCellArguments:
    def test_constructs_cell_ref(self, mock_bridge: MagicMock) -> None:
        restore.restore_cell("b.xlsx", "Sheet1", "B2")
        called_ref = mock_bridge.restore_cell.call_args.args[0]
        assert isinstance(called_ref, CellRef)
        assert called_ref.workbook == "b.xlsx"
        assert called_ref.sheet == "Sheet1"
        assert called_ref.cell == "B2"

    def test_since_omitted_passes_none(self, mock_bridge: MagicMock) -> None:
        restore.restore_cell("b.xlsx", "S1", "A1")
        kwargs = mock_bridge.restore_cell.call_args.kwargs
        assert kwargs == {"since": None}

    def test_since_parsed_to_datetime(self, mock_bridge: MagicMock) -> None:
        restore.restore_cell(
            "b.xlsx", "S1", "A1", since="2026-05-20T10:30:00"
        )
        kwargs = mock_bridge.restore_cell.call_args.kwargs
        assert isinstance(kwargs["since"], datetime)
        assert kwargs["since"].year == 2026
        assert kwargs["since"].hour == 10

    def test_since_with_timezone_parsed(self, mock_bridge: MagicMock) -> None:
        restore.restore_cell(
            "b.xlsx", "S1", "A1", since="2026-05-20T10:30:00+00:00"
        )
        since = mock_bridge.restore_cell.call_args.kwargs["since"]
        assert since.tzinfo is not None
        assert since.utcoffset() == UTC.utcoffset(since)

    def test_invalid_since_raises_value_error(
        self, mock_bridge: MagicMock
    ) -> None:
        with pytest.raises(ValueError, match="ISO-8601"):
            restore.restore_cell("b.xlsx", "S1", "A1", since="not-a-date")
        # Bridge must NOT be called when input is invalid — fail fast.
        mock_bridge.restore_cell.assert_not_called()

    def test_returns_bridge_result_unchanged(
        self, mock_bridge: MagicMock
    ) -> None:
        result = restore.restore_cell("b.xlsx", "S1", "A1")
        assert isinstance(result, InsertResult)
        assert result.formula == "=42"
        assert result.previous_formula == "=99"
