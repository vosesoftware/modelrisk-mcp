"""Assemble the Claude Desktop Extension (.mcpb) bundle from the built exe.

A .mcpb is a ZIP archive with manifest.json at its root and the standalone
server binary under server/. Claude Desktop installs it in one click — no
Python, no config editing. This is CLI-free (plain zipfile) so it doesn't
depend on the external `mcpb` packer being installed in CI.

Usage:
    python scripts/build_mcpb.py <exe_path> <version> <out.mcpb>

The version (e.g. derived from the release tag) is injected into the bundled
manifest's `version` field so it always matches the release.
"""

from __future__ import annotations

import json
import os
import shutil
import sys
import zipfile
from pathlib import Path


def main() -> int:
    if len(sys.argv) != 4:
        print(__doc__)
        return 2
    exe_path, version, out = sys.argv[1], sys.argv[2], sys.argv[3]

    root = Path(__file__).resolve().parent.parent
    manifest_src = root / "packaging" / "manifest.json"
    manifest = json.loads(manifest_src.read_text(encoding="utf-8"))
    manifest["version"] = version

    # Sanity: the manifest must reference the exe we're bundling by basename.
    exe_name = os.path.basename(exe_path)
    entry = manifest["server"]["entry_point"]
    if os.path.basename(entry) != exe_name:
        raise SystemExit(
            f"manifest entry_point {entry!r} does not match exe {exe_name!r}"
        )

    build = root / "build" / "mcpb"
    if build.exists():
        shutil.rmtree(build)
    (build / "server").mkdir(parents=True)
    shutil.copy2(exe_path, build / "server" / exe_name)
    (build / "manifest.json").write_text(
        json.dumps(manifest, indent=2) + "\n", encoding="utf-8"
    )

    out_path = Path(out)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    if out_path.exists():
        out_path.unlink()
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as z:
        for f in sorted(build.rglob("*")):
            if f.is_file():
                z.write(f, f.relative_to(build).as_posix())

    size = out_path.stat().st_size
    print(f"Built {out_path} ({size:,} bytes) — manifest version {version}")
    # Echo the manifest for the CI log.
    print(json.dumps(manifest, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
