"""Tests for the CLI argument parser."""

from __future__ import annotations

import pytest

from modelrisk_mcp.__main__ import _build_parser


class TestCliParser:
    def test_defaults_to_stdio(self) -> None:
        args = _build_parser().parse_args([])
        assert args.transport == "stdio"
        assert args.host == "127.0.0.1"
        assert args.port == 8000
        assert args.token is None

    def test_streamable_http_flag(self) -> None:
        args = _build_parser().parse_args(["--transport=streamable-http"])
        assert args.transport == "streamable-http"

    def test_sse_flag(self) -> None:
        args = _build_parser().parse_args(["--transport=sse"])
        assert args.transport == "sse"

    def test_invalid_transport_rejected(self) -> None:
        with pytest.raises(SystemExit):
            _build_parser().parse_args(["--transport=carrier-pigeon"])

    def test_token_flag(self) -> None:
        args = _build_parser().parse_args(["--token=secret-123"])
        assert args.token == "secret-123"

    def test_host_and_port(self) -> None:
        args = _build_parser().parse_args(
            ["--transport=streamable-http", "--host=0.0.0.0", "--port=9000"]
        )
        assert args.host == "0.0.0.0"
        assert args.port == 9000

    def test_mount_path(self) -> None:
        args = _build_parser().parse_args(["--mount-path=/custom"])
        assert args.mount_path == "/custom"
