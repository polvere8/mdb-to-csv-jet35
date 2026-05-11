"""
diagnose.py
-----------
Inspect the raw header of a Jet 3.5 / Access 97 .mdb file to detect:
  - Jet version (3.5 vs 4.0)
  - Encryption flag
  - Database password (tries all known XOR keys)
  - Whether the file is a normal unprotected "Standard Jet DB"

Run this on one sample file before attempting batch export.
It costs nothing and tells you exactly what you are dealing with.

Usage:
    python diagnose.py path/to/file.mdb

Note: use forward slashes or quotes around paths with spaces.
"""

import sys
import os

# ── Known Jet 3.5 XOR keys for password decoding ─────────────────────────────
# Jet 3.5 stores the database password at offset 0x18, XOR-encoded.
# The standard key is 0x86, but some OEM builds used different keys.
JET3_SINGLE_KEYS = [
    0x86, 0x00, 0xFF, 0xA1, 0x6B, 0xC4,
    0x39, 0x19, 0x2F, 0x55, 0xAA, 0x35, 0xD3,
]

# Some Jet 3.5 SP3+ builds used a rolling (per-byte) key
ROLLING_KEY = [
    0x86, 0xFB, 0xEC, 0x37, 0x5D, 0x44,
    0x9C, 0x35, 0xC5, 0xFB, 0x4B, 0x15, 0xA5, 0x22,
]


def is_printable(b: bytes, threshold: float = 0.7) -> bool:
    """Return True if most bytes are printable ASCII."""
    ratio = sum(32 <= x < 127 for x in b) / max(len(b), 1)
    return ratio >= threshold


def decode_password(raw: bytes, key) -> str:
    if isinstance(key, list):
        decoded = bytes([b ^ key[i % len(key)] for i, b in enumerate(raw)])
    else:
        decoded = bytes([b ^ key for b in raw])
    return decoded.split(b"\x00")[0].decode("latin-1", errors="replace")


def hex_dump(data: bytes, width: int = 16) -> str:
    lines = []
    for i in range(0, len(data), width):
        chunk = data[i : i + width]
        hex_part = " ".join(f"{b:02x}" for b in chunk)
        asc_part = "".join(chr(b) if 32 <= b < 127 else "." for b in chunk)
        lines.append(f"  {i:04x}: {hex_part:<{width*3}}  {asc_part}")
    return "\n".join(lines)


def diagnose(filepath: str) -> None:
    if not os.path.isfile(filepath):
        print(f"ERROR: File not found: {filepath}")
        sys.exit(1)

    size = os.path.getsize(filepath)
    print(f"\n{'='*60}")
    print(f"File  : {os.path.basename(filepath)}")
    print(f"Size  : {size:,} bytes")

    with open(filepath, "rb") as f:
        header = f.read(256)

    if len(header) < 32:
        print("ERROR: File too small to be a valid MDB.")
        return

    # ── Magic bytes ───────────────────────────────────────────────────────────
    magic = header[:4].hex()
    valid = magic == "00010000"
    print(f"\n[1] Magic bytes : {magic}  {'OK: valid Jet file' if valid else 'ERROR: unexpected -- may be corrupted or non-Jet'}")

    # ── Jet version ───────────────────────────────────────────────────────────
    ver_byte = header[0x14]
    ver_str = {0: "Jet 3.5 (Access 97)", 1: "Jet 4.0 (Access 2000/XP/2003)"}.get(ver_byte, f"Unknown ({ver_byte})")
    print(f"[2] Jet version : byte 0x14 = {ver_byte}  ->  {ver_str}")

    # ── Encryption flag ───────────────────────────────────────────────────────
    enc_byte = header[0x15]
    print(f"[3] Encryption  : byte 0x15 = {enc_byte}  ->  {'WARNING: encrypted' if enc_byte != 0 else 'not encrypted'}")

    # ── Database name tag ─────────────────────────────────────────────────────
    tag_bytes = header[4:20]
    tag_str = tag_bytes.decode("latin-1", errors="replace").rstrip("\x00")
    print(f"[4] DB tag      : \"{tag_str}\"")
    if "Standard Jet DB" in tag_str:
        print("    -> Normal unprotected Standard Jet DB file.")

    # ── Password region ───────────────────────────────────────────────────────
    raw_pw = header[0x18:0x2C]
    print(f"\n[5] Password region (offset 0x18, 20 bytes):")
    print(f"    Raw hex : {raw_pw.hex()}")

    found_password = None
    print("    Trying XOR keys:")
    for key in JET3_SINGLE_KEYS:
        candidate = decode_password(raw_pw, key)
        marker = ""
        if is_printable(candidate.encode("latin-1", errors="replace")) and candidate.strip():
            marker = "  <- READABLE - likely password!"
            found_password = candidate
        print(f"      key 0x{key:02X} : {repr(candidate)}{marker}")

    candidate_rolling = decode_password(raw_pw, ROLLING_KEY)
    marker = "  <- READABLE!" if is_printable(candidate_rolling.encode("latin-1", errors="replace")) else ""
    print(f"      rolling  : {repr(candidate_rolling)}{marker}")

    # ── Summary ───────────────────────────────────────────────────────────────
    print(f"\n[6] Summary:")
    if not valid:
        print("    ERROR: Not a valid Jet file. Cannot process.")
    elif enc_byte != 0:
        print("    WARNING: File is encrypted (workgroup security).")
        print("    -> You need the .mdw workgroup file alongside the .mdb.")
        print("    -> Try: mdb-tables --sysdb=System.mdw file.mdb")
    elif found_password:
        print(f"    WARNING: Possible password detected: \"{found_password}\"")
        print(f"    -> Note: this may be a false positive.")
        print(f"    -> First try mdb-tables WITHOUT -p flag.")
        print(f"    -> Only use -p if that returns blank: mdb-tables -p \"{found_password}\" file.mdb")
    else:
        print("    OK: File appears to be a normal, unprotected Jet 3.5 database.")
        print("    -> Use lsgunth/mdbtools-win to export.")
        print("    -> See batch_export.py")

    # ── Full hex dump ─────────────────────────────────────────────────────────
    print(f"\n[7] Full header hex dump (first 128 bytes):")
    print(hex_dump(header[:128]))
    print()


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python diagnose.py path/to/file.mdb")
        sys.exit(1)
    diagnose(sys.argv[1])
