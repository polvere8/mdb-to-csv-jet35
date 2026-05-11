"""
batch_export.py
---------------
Batch export all tables from all .mdb files (Jet 3.5 / Access 97)
to CSV using lsgunth/mdbtools-win.

Prerequisites:
  1. Download mdbtools-win from:
     https://github.com/lsgunth/mdbtools-win/releases
  2. Extract the zip somewhere on your machine
  3. Set MDBTOOLS below to that folder path

No Python packages required -- uses only standard library.

Usage:
    python batch_export.py

Output structure:
    OUTPUT_ROOT/
    └── filename_without_extension/
        ├── Table1.csv
        ├── Table2.csv
        └── ...
"""

import subprocess
import os
import time
from pathlib import Path

# -- Configuration: change these three paths --

MDBTOOLS    = r"C:\mdbtools"            # folder containing mdb-export.exe
MDB_ROOT    = r"C:\path\to\mdb\files"  # folder containing your .mdb files (searched recursively)
OUTPUT_ROOT = r"C:\path\to\output"     # where CSVs will be written

# If your files have a password, set it here. Leave as None if no password.
PASSWORD = None   # e.g. PASSWORD = "mypassword"

# -- Paths to executables --

MDB_EXPORT = str(Path(MDBTOOLS) / "mdb-export.exe")
MDB_TABLES = str(Path(MDBTOOLS) / "mdb-tables.exe")
MDB_VER    = str(Path(MDBTOOLS) / "mdb-ver.exe")


def check_mdbtools() -> bool:
    """Verify mdbtools executables exist and work."""
    for exe in [MDB_EXPORT, MDB_TABLES, MDB_VER]:
        if not os.path.isfile(exe):
            print(f"ERROR: Not found: {exe}")
            print(f"       Download from: https://github.com/lsgunth/mdbtools-win/releases")
            return False
    return True


def get_tables(mdb_path: Path) -> list:
    """Return list of user table names in the given .mdb file."""
    cmd = [MDB_TABLES, "-1"]
    if PASSWORD:
        cmd += ["-p", PASSWORD]
    cmd.append(str(mdb_path))

    result = subprocess.run(cmd, capture_output=True, text=True)
    return [t.strip() for t in result.stdout.splitlines() if t.strip()]


def export_table(mdb_path: Path, table: str, out_csv: Path) -> bool:
    """Export one table to CSV. Returns True on success."""
    cmd = [MDB_EXPORT]
    if PASSWORD:
        cmd += ["-p", PASSWORD]
    cmd += [str(mdb_path), table]

    try:
        with open(out_csv, "w", newline="", encoding="utf-8") as f:
            result = subprocess.run(cmd, stdout=f, stderr=subprocess.PIPE)
        return result.returncode == 0
    except Exception as e:
        print(f"      ERROR exporting: {e}")
        return False


def export_mdb(mdb_path: Path, output_root: Path) -> tuple:
    """
    Export all tables from one .mdb file.
    Returns (tables_exported, status_string).
    """
    rel     = mdb_path.relative_to(Path(MDB_ROOT))
    out_dir = output_root / rel.parent / rel.stem
    out_dir.mkdir(parents=True, exist_ok=True)

    tables = get_tables(mdb_path)

    if not tables:
        return 0, "no_tables"

    exported = 0
    for table in tables:
        out_csv = out_dir / f"{table}.csv"
        ok = export_table(mdb_path, table, out_csv)
        if ok:
            exported += 1

    return exported, "ok"


def main():
    print("MDB -> CSV Batch Exporter")
    print("=" * 50)

    # Sanity checks
    if not check_mdbtools():
        return

    if not os.path.isdir(MDB_ROOT):
        print(f"ERROR: MDB_ROOT does not exist: {MDB_ROOT}")
        return

    # Find all .mdb files
    files = list(Path(MDB_ROOT).rglob("*.mdb"))
    if not files:
        print(f"No .mdb files found in: {MDB_ROOT}")
        return

    print(f"Found   : {len(files)} .mdb files")
    print(f"Output  : {OUTPUT_ROOT}")
    if PASSWORD:
        print(f"Password: {'*' * len(PASSWORD)}")
    print()

    # Verify first file reads correctly
    print("Checking first file with mdb-ver...")
    ver_result = subprocess.run([MDB_VER, str(files[0])], capture_output=True, text=True)
    print(f"  {files[0].name} -> {ver_result.stdout.strip()}")
    if "JET3" not in ver_result.stdout and "JET4" not in ver_result.stdout:
        print("  WARNING: mdb-ver did not return a known Jet version.")
        print("  Run diagnose.py on this file first to understand the format.")
    print()

    # Batch export
    start_time = time.time()
    ok_count, no_tables, errors = 0, [], []

    for i, mdb_file in enumerate(files):
        count, status = export_mdb(mdb_file, Path(OUTPUT_ROOT))

        if status == "ok":
            ok_count += 1
            print(f"[{i+1:>4}/{len(files)}] OK  {mdb_file.name}  ({count} tables)")
        elif status == "no_tables":
            no_tables.append(str(mdb_file))
            print(f"[{i+1:>4}/{len(files)}] --  {mdb_file.name}  (no tables found)")
        else:
            errors.append((str(mdb_file), status))
            print(f"[{i+1:>4}/{len(files)}] ERR {mdb_file.name}  ({status})")

    elapsed = time.time() - start_time

    # Final summary
    print()
    print("=" * 50)
    print(f"Done in {elapsed:.1f}s")
    print(f"Exported  : {ok_count}")
    print(f"No tables : {len(no_tables)}")
    print(f"Errors    : {len(errors)}")

    if no_tables:
        print("\nFiles with no tables (may be password protected):")
        for f in no_tables:
            print(f"  {f}")
        print("  -> Run diagnose.py on these files to investigate.")

    if errors:
        print("\nFailed files:")
        for path, err in errors:
            print(f"  {path}: {err}")


if __name__ == "__main__":
    main()
