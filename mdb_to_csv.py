"""
mdb_to_csv.py
-------------
Batch export all tables from all .mdb / .accdb files (Jet 4.0 / Access 2000+)
to CSV using pyodbc and the Microsoft Access Database Engine.

Use this script for modern Access files (Jet 4.0).
For legacy Jet 3.5 / Access 97 files, use batch_export.py instead.

Not sure which version your files are? Run diagnose.py first:
    python diagnose.py path/to/file.mdb

Prerequisites:
  1. Install Python libraries:
     pip install pyodbc pandas

  2. Install the Microsoft Access Database Engine:
     https://www.microsoft.com/en-us/download/details.aspx?id=54920

     To find which version (32 or 64 bit) to download, run:
     python -c "import struct; print(struct.calcsize('P')*8)"

  3. Edit the two paths below (INPUT_FOLDER and OUTPUT_FOLDER)

Usage:
    python mdb_to_csv.py
"""

import os
import sys
from pathlib import Path
import pandas as pd
import pyodbc

# -- Configuration: change these two paths --

INPUT_FOLDER  = r"C:\path\to\mdb\files"   # folder containing your .mdb files
OUTPUT_FOLDER = r"C:\path\to\output"      # where CSVs will be written

# 


def get_driver():
    drivers = [d for d in pyodbc.drivers() if "Access" in d]
    if not drivers:
        print("ERROR: Microsoft Access driver not found.")
        print("Download: https://www.microsoft.com/en-us/download/details.aspx?id=54920")
        sys.exit(1)
    return drivers[0]


def find_mdb_files(folder):
    folder = Path(folder)
    files = []
    for ext in ["*.mdb", "*.accdb", "*.MDB", "*.ACCDB"]:
        files.extend(folder.rglob(ext))
    seen, unique = set(), []
    for f in files:
        k = str(f).lower()
        if k not in seen:
            seen.add(k)
            unique.append(f)
    return sorted(unique)


def export_file(mdb_path, output_dir, driver):
    os.makedirs(output_dir, exist_ok=True)

    conn_str = (
        f"Driver={{{driver}}};"
        f"DBQ={os.path.abspath(mdb_path)};"
        "ExtendedAnsiSQL=1;"
    )

    try:
        conn = pyodbc.connect(conn_str)
    except Exception as e:
        print(f"  Cannot open file: {e}")
        print("  If this is a Jet 3.5 / Access 97 file, use batch_export.py instead.")
        return 0, 1

    cursor = conn.cursor()
    tables = [
        row.table_name
        for row in cursor.tables(tableType="TABLE")
        if not row.table_name.startswith("MSys")
    ]

    if not tables:
        print("  No tables found -- skipping")
        conn.close()
        return 0, 0

    success, failed = 0, 0

    for table in tables:
        print(f"  {table} ...", end=" ", flush=True)
        try:
            df = pd.read_sql(f"SELECT * FROM [{table}]", conn)

            if df.empty:
                print("empty -- skipped")
                continue

            safe = "".join(
                c if c.isalnum() or c in ("_", "-") else "_"
                for c in table
            )

            out_path = os.path.join(output_dir, f"{safe}.csv")
            df.to_csv(out_path, index=False, encoding="utf-8-sig")
            print(f"OK  ({len(df):,} rows, {len(df.columns)} cols)")
            success += 1

        except Exception as e:
            print(f"FAILED -- {e}")
            failed += 1

    conn.close()
    return success, failed


def main():
    print("=" * 60)
    print("  MDB to CSV -- Jet 4.0 / Access 2000+ files")
    print("=" * 60)
    print(f"  Input  : {INPUT_FOLDER}")
    print(f"  Output : {OUTPUT_FOLDER}")
    print("=" * 60)

    if not os.path.isdir(INPUT_FOLDER):
        print(f"\nERROR: Input folder not found:\n  {INPUT_FOLDER}")
        sys.exit(1)

    driver = get_driver()
    print(f"\nDriver found: {driver}\n")

    files = find_mdb_files(INPUT_FOLDER)

    if not files:
        print(f"No .mdb or .accdb files found in: {INPUT_FOLDER}")
        sys.exit(1)

    print(f"Found {len(files)} file(s)\n")

    total_success = 0
    total_failed  = 0
    results       = []

    for i, mdb_path in enumerate(files, 1):
        relative        = mdb_path.relative_to(Path(INPUT_FOLDER))
        file_output_dir = Path(OUTPUT_FOLDER) / relative.parent / mdb_path.stem

        print(f"[{i}/{len(files)}] {relative}")

        s, f = export_file(
            mdb_path   = str(mdb_path),
            output_dir = str(file_output_dir),
            driver     = driver
        )

        total_success += s
        total_failed  += f
        results.append((str(relative), s, f))
        print()

    print("=" * 60)
    print("  SUMMARY")
    print("=" * 60)
    print(f"  {'File':<40} {'OK':>4} {'Fail':>4}")
    print("-" * 60)
    for name, s, f in results:
        flag = "!" if f > 0 else " "
        print(f" {flag} {name:<40} {s:>4} {f:>4}")
    print("-" * 60)
    print(f"  {'TOTAL':<40} {total_success:>4} {total_failed:>4}")
    print("=" * 60)
    print(f"\nCSV files saved to:\n  {os.path.abspath(OUTPUT_FOLDER)}")

    if total_failed > 0:
        print(f"\nWARNING: {total_failed} table(s) failed -- check output above.")
        print("If you are getting 'Unrecognized database format', your files")
        print("may be Jet 3.5 / Access 97. Run diagnose.py to check, then")
        print("use batch_export.py instead.")


if __name__ == "__main__":
    main()
