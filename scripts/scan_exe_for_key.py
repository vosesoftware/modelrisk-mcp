"""Scan the built modelrisk-mcp.exe for any encoding of the plain
activation key. Use after every PyInstaller build to verify the
obfuscation is intact.

Usage:
    python scripts/scan_exe_for_key.py [path_to_exe]

Exits 0 if clean, 1 if any plain-key occurrence is detected.

This is a paranoid scan — it looks for the key as:
- ASCII decimal string
- UTF-16 LE wide string (Windows convention)
- Raw int64 little-endian bytes (the form a ctypes c_int64 takes in
  memory; if the value ever got stored as a Python int constant in a
  .pyc, it could appear here)
- Raw int64 big-endian bytes (defence in depth)
- 8-byte big-endian compact form (matches how _keymat encodes)
- Hex string form
- Composite: any ASCII run containing both the key's first 4 chars
  AND last 4 chars (defeats simple obfuscation that breaks the digits
  across runs)

The bundled key value is decoded via `_keymat` so the literal is never
written to the source of this script either — same threat model.
"""

from __future__ import annotations

import re
import struct
import sys
from pathlib import Path


def _decode_key() -> int:
    """Import the decoder and decode the bundled key. We do this rather
    than hardcoding the value to keep the literal out of this script
    too."""
    sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))
    from modelrisk_mcp.bridge._keymat import decode_bundled_key

    return decode_bundled_key()


def scan(exe_path: Path) -> int:
    if not exe_path.is_file():
        print(f"error: {exe_path} not found", file=sys.stderr)
        return 2

    key = _decode_key()
    key_str = str(key)

    print(f"target: {exe_path}")
    print(f"size:   {exe_path.stat().st_size:,} bytes")
    print()

    blob = exe_path.read_bytes()
    hits: dict[str, int] = {}

    # 1. ASCII decimal
    hits["ASCII decimal"] = blob.count(key_str.encode())

    # 2. UTF-16 LE
    hits["UTF-16 LE wide"] = blob.count(key_str.encode("utf-16-le"))

    # 3-5. Raw int64 packings
    hits["int64 LE bytes"] = blob.count(struct.pack("<q", key))
    hits["int64 BE bytes"] = blob.count(struct.pack(">q", key))
    hits["int64 BE compact"] = blob.count(key.to_bytes(8, "big"))

    # 6. Hex string (both cases)
    hex_str = format(key, "X").encode()
    hits["hex string (upper)"] = blob.count(hex_str)
    hits["hex string (lower)"] = blob.count(hex_str.lower())

    # 7. Composite: any printable ASCII run containing first+last 4 chars
    strings_pattern = re.compile(rb"[\x20-\x7e]{4,}")
    runs = strings_pattern.findall(blob)
    head = key_str[:4].encode()
    tail = key_str[-4:].encode()
    composite = sum(1 for r in runs if head in r and tail in r)
    hits["composite (first4+last4)"] = composite

    print("per-encoding scan:")
    for label, count in hits.items():
        flag = "OK" if count == 0 else f"WARNING ({count} hits)"
        print(f"  {label:<28}  {flag}")

    total = sum(hits.values())
    print()
    if total == 0:
        print("RESULT: clean — no plain-key encoding found.")
        return 0
    print(f"RESULT: FAILED — {total} plain-key occurrences across encodings.")
    return 1


def main(argv: list[str]) -> int:
    if len(argv) > 1 and argv[1] in {"-h", "--help"}:
        print(__doc__.strip())
        return 0
    exe = Path(argv[1]) if len(argv) > 1 else Path("dist/modelrisk-mcp.exe")
    return scan(exe)


if __name__ == "__main__":
    sys.exit(main(sys.argv))
