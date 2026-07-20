import os
from dataclasses import dataclass, field
from pathlib import Path


def _default_log_dir() -> Path:
    base = os.environ.get("LOCALAPPDATA") or str(Path.home() / "AppData" / "Local")
    return Path(base) / "VoseSoftware" / "modelrisk-mcp"


_TRUTHY = {"1", "true", "yes", "on"}


def read_only_active(settings: "Settings | None" = None) -> bool:
    """True when the server should refuse writes/simulations/saves.

    Checked AT CALL TIME so the `--read-only` CLI flag (which sets
    MODELRISK_MCP_READ_ONLY before the server starts) and programmatic
    `Settings(read_only=True)` both work regardless of import order.
    Until 0.3.11 `Settings.read_only` existed but nothing set or
    enforced it — documented-but-not-implemented (field bug report,
    2026-07-20)."""
    if settings is not None and settings.read_only:
        return True
    return os.environ.get("MODELRISK_MCP_READ_ONLY", "").strip().lower() in _TRUTHY


@dataclass(frozen=True)
class Settings:
    read_only: bool = False
    log_dir: Path = field(default_factory=_default_log_dir)
    writes_log_name: str = "writes.log"

    @property
    def writes_log_path(self) -> Path:
        return self.log_dir / self.writes_log_name
