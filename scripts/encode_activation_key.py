"""Encode an MRService.dll activation key for bundling.

Usage:
    python scripts/encode_activation_key.py <int64_key>
    python scripts/encode_activation_key.py <int64_key> <16-byte-salt-hex>

The output is two base85 strings: a salt and a blob. Paste them into
`src/modelrisk_mcp/bridge/_keymat.py` to rotate the bundled key. The
algorithm is intentionally public — security depends on the key value
being unknown to attackers, not on hiding how it's encoded
(Kerckhoffs's principle). What this script buys us is that the literal
integer never appears in shipped source, .pyc caches, or `strings`
output of a PyInstaller bundle.

Layered obfuscation:
1. Pack the int64 as 8 big-endian bytes.
2. XOR with a per-byte rotating pad (16-byte salt repeated as needed).
3. Base85-encode both salt and ciphertext.

A `strings` scan of the result yields high-entropy base85 chunks with
no recognisable substring of the original key.
"""

from __future__ import annotations

import base64
import secrets
import sys


def _encode(key: int, salt: bytes) -> bytes:
    if not (0 <= key < (1 << 63)):
        raise ValueError(
            f"Key {key!r} must fit in a positive int64 (0..2^63-1)."
        )
    if len(salt) != 16:
        raise ValueError(f"Salt must be 16 bytes, got {len(salt)}.")
    plain = key.to_bytes(8, byteorder="big", signed=False)
    pad = (salt * ((len(plain) // len(salt)) + 1))[: len(plain)]
    return bytes(b ^ p for b, p in zip(plain, pad, strict=True))


def main(argv: list[str]) -> int:
    if len(argv) < 2 or argv[1] in {"-h", "--help"}:
        print(__doc__.strip())
        return 0
    try:
        key = int(argv[1])
    except ValueError:
        print(f"error: {argv[1]!r} is not an integer", file=sys.stderr)
        return 1
    if len(argv) >= 3:
        try:
            salt = bytes.fromhex(argv[2])
        except ValueError:
            print(f"error: {argv[2]!r} is not valid hex", file=sys.stderr)
            return 1
    else:
        salt = secrets.token_bytes(16)
    blob = _encode(key, salt)

    salt_b85 = base64.b85encode(salt).decode("ascii")
    blob_b85 = base64.b85encode(blob).decode("ascii")
    print("Paste these into src/modelrisk_mcp/bridge/_keymat.py:")
    print()
    print(f'_SALT_B85 = b"{salt_b85}"')
    print(f'_BLOB_B85 = b"{blob_b85}"')
    print()

    # Round-trip sanity check.
    decoded_plain = bytes(b ^ p for b, p in zip(
        blob, (salt * ((len(blob) // len(salt)) + 1))[: len(blob)], strict=True,
    ))
    decoded = int.from_bytes(decoded_plain, byteorder="big", signed=False)
    if decoded != key:
        print(f"error: round-trip mismatch: {decoded} != {key}", file=sys.stderr)
        return 1
    print(f"# round-trip OK ({decoded})", file=sys.stderr)
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
