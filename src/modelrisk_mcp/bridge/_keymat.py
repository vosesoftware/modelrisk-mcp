"""Obfuscated bundled activation key material for MRService.dll.

This module exists so the literal int64 key never appears as a constant
in shipped source, .pyc caches, or `strings` output of a PyInstaller
bundle. The key is XOR-encoded against a 16-byte salt and base85-stored
in two short blobs. The plain integer only exists transiently — as the
return value of `decode_bundled_key()` — and is consumed immediately by
`MRLIB_SetOfflineActivationKey` in `bridge/mrservice.py`.

This is light obfuscation, not real cryptography. A determined attacker
with a debugger can extract the value by breakpointing on the DLL entry
point. The bar this raises:

- `strings` on the wheel / PyInstaller exe finds no integer or substring
  resembling the original key.
- Grep through the repo finds neither the integer nor a decimal-shaped
  string that decodes to it.
- Casual readers of the source see only opaque base85 blobs and need to
  either run the code or trace the algorithm by hand to recover the key.

The encoding algorithm is intentionally documented (in
`scripts/encode_activation_key.py`): security depends on the key value
being unknown to the attacker, not on the algorithm being secret
(Kerckhoffs's principle).

Rotation:
    python scripts/encode_activation_key.py <new_int64>
and paste the resulting `_SALT_B85` / `_BLOB_B85` over the values below.
"""

from __future__ import annotations

import base64

# Updated by `scripts/encode_activation_key.py`. Both blobs are short
# base85 strings — they look like random tokens to `strings` and to a
# human skimming the file.
_SALT_B85 = b"gQALc9uunRIzibSr<+%a"
_BLOB_B85 = b"gQALk6Le1w"


def decode_bundled_key() -> int:
    """Reconstruct the bundled int64 activation key.

    Returns the plain integer; callers should pass it straight to the
    DLL entry point and not store it. Catching this value in a debugger
    or memory dump is trivially possible — that's an acceptable cost.
    """
    salt = base64.b85decode(_SALT_B85)
    blob = base64.b85decode(_BLOB_B85)
    pad = (salt * ((len(blob) // len(salt)) + 1))[: len(blob)]
    plain = bytes(b ^ p for b, p in zip(blob, pad, strict=True))
    return int.from_bytes(plain, byteorder="big", signed=False)


__all__ = ["decode_bundled_key"]
