import os
from dataclasses import dataclass, field
from pathlib import Path


def _default_log_dir() -> Path:
    base = os.environ.get("LOCALAPPDATA") or str(Path.home() / "AppData" / "Local")
    return Path(base) / "VoseSoftware" / "modelrisk-mcp"


@dataclass(frozen=True)
class Settings:
    read_only: bool = False
    log_dir: Path = field(default_factory=_default_log_dir)
    writes_log_name: str = "writes.log"

    @property
    def writes_log_path(self) -> Path:
        return self.log_dir / self.writes_log_name
