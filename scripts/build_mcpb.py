"""Build the Claude Desktop Extension (.mcpb) from the built exe.

A .mcpb is a ZIP with manifest.json at the root and the standalone server
binary under server/. This stages that layout (injecting the release version
into the manifest) and packs it with the official `@anthropic-ai/mcpb` CLI via
`npx` — which validates the manifest during packing. If node/npx isn't
available (e.g. a bare local checkout), it falls back to a plain-zip pack that
produces a byte-structurally identical bundle.

Usage:
    python scripts/build_mcpb.py <exe_path> <version> <out.mcpb>
"""

from __future__ import annotations

import json
import shutil
import subprocess
import sys
import zipfile
from pathlib import Path

_MCPB_PKG = "@anthropic-ai/mcpb@latest"


def _stage(exe_path: Path, version: str, stage: Path) -> str:
    """Lay out manifest.json + server/<exe> under `stage`; return the exe name."""
    root = Path(__file__).resolve().parent.parent
    manifest = json.loads((root / "packaging" / "manifest.json").read_text(encoding="utf-8"))
    manifest["version"] = version

    exe_name = exe_path.name
    entry = manifest["server"]["entry_point"]
    if Path(entry).name != exe_name:
        raise SystemExit(f"manifest entry_point {entry!r} does not match exe {exe_name!r}")

    if stage.exists():
        shutil.rmtree(stage)
    (stage / "server").mkdir(parents=True)
    shutil.copy2(exe_path, stage / "server" / exe_name)
    (stage / "manifest.json").write_text(json.dumps(manifest, indent=2) + "\n", encoding="utf-8")
    return exe_name


def _pack_with_cli(npx: str, stage: Path, out: Path) -> bool:
    """Validate + pack with the official mcpb CLI. Returns True on success."""
    try:
        subprocess.run([npx, "-y", _MCPB_PKG, "validate", str(stage / "manifest.json")], check=True)
        subprocess.run([npx, "-y", _MCPB_PKG, "pack", str(stage), str(out)], check=True)
        return True
    except (subprocess.CalledProcessError, OSError) as exc:
        print(f"mcpb CLI pack failed ({exc}); falling back to plain-zip.")
        return False


def _pack_plain_zip(stage: Path, out: Path) -> None:
    """Fallback: zip the staged dir (manifest.json at root)."""
    if out.exists():
        out.unlink()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as z:
        for f in sorted(stage.rglob("*")):
            if f.is_file():
                z.write(f, f.relative_to(stage).as_posix())


def main() -> int:
    if len(sys.argv) != 4:
        print(__doc__)
        return 2
    exe_path, version, out = Path(sys.argv[1]), sys.argv[2], Path(sys.argv[3])

    stage = (out.parent if out.parent.name else Path(".")) / "_mcpb_stage"
    _stage(exe_path, version, stage)

    out.parent.mkdir(parents=True, exist_ok=True)
    npx = shutil.which("npx")
    if not (npx and _pack_with_cli(npx, stage, out)):
        _pack_plain_zip(stage, out)

    size = out.stat().st_size
    print(f"Built {out} ({size:,} bytes) — manifest version {version}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
