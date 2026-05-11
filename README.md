# MDB to CSV — Jet 3.5 / Access 97 on Modern Windows

> **The complete guide to extracting data from legacy `.mdb` files on Windows 10/11 with Python — including every approach that fails and exactly why.**

⭐ **If this guide saved you hours of pain, star the repo** — it helps others find it before they waste time on the approaches that don't work.

---

## Step 0 — Find out what kind of files you have

Before doing anything else, run `diagnose.py` on one sample file:

```powershell
python diagnose.py "C:/path/to/your/file.mdb"
```

Look at line `[2] Jet version`:
- `Jet 4.0 (Access 2000+)` → your files are modern, go to the **Jet 4.0 section** below
- `Jet 3.5 (Access 97)` → your files are legacy, go to the **Jet 3.5 section** below

---

## Jet 4.0 files (Access 2000+) — pyodbc method

Install the required driver and library:

```
pip install pyodbc pandas
```

Download the Microsoft Access Database Engine from:
https://www.microsoft.com/en-us/download/details.aspx?id=54920

To know which version (32 or 64 bit) to download, run:
```
python -c "import struct; print(struct.calcsize('P')*8)"
```

Then use `mdb_to_csv.py` from this repo. Edit the two paths at the top and run:
```
python mdb_to_csv.py
```

---

## Jet 3.5 files (Access 97) — mdbtools method

pyodbc **will not work** on these files. Do not try. See below for exactly why.

### The Problem

You have `.mdb` files created by old industrial machines (CNC controllers, winding machines, PLCs, lab equipment) running **Access 97 / Jet 3.5** from the late 1990s or early 2000s.

You open Python, try the obvious solution, and hit a wall:

```
pyodbc.Error: ('HY000', '[HY000] Unrecognized database format')
```

You search Stack Overflow. Every answer says the same thing: install the ACE driver, use pyodbc. **It does not work.** Here is why, and here is what actually works.

---

### Why the Obvious Solutions Fail

#### ❌ pyodbc + Microsoft Access Driver
```python
conn = pyodbc.connect(r"Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=file.mdb")
# → HY000: Unrecognized database format
```
The `Microsoft Access Driver (*.mdb, *.accdb)` that ships with Microsoft 365 is the **ACE 14.0 engine (Jet 4.0)**. It explicitly refuses to open Jet 3.5 files. Microsoft dropped backwards compatibility.

#### ❌ DAO.DBEngine.36
```python
engine = win32com.client.Dispatch("DAO.DBEngine.36")
# → (-2147221164, 'Class not registered', None, None)
```
`DBEngine.36` is the Jet 3.5 COM object. Microsoft 365 **does not install it**. It was removed. The only DAO engine registered is `DBEngine.120` (ACE), which has the same version mismatch problem.

#### ❌ Access.Application via win32com
```python
access = win32com.client.Dispatch("Access.Application")
access.OpenCurrentDatabase(path)
db = access.CurrentDb()  # → returns None
```
Access.Application silently fails to open Jet 3.5 files when the installed Access version is 365/2019/2021. `CurrentDb()` returns `None` without throwing any error.

#### ❌ Windows port of MDBTools (most GitHub links)
The commonly linked Windows builds of MDBTools (e.g. `cyrez/mdbtools`) return empty output:
```
mdb-tables.exe file.mdb
→ (blank)
```
These builds have incomplete Jet 3.5 support and fail silently on many files.

---

### The Solution That Works

**Use the `lsgunth/mdbtools-win` build** — a newer, better-compiled Windows port of MDBTools with full Jet 3.5 support.

#### Step 1 — Download the right MDBTools build

1. Go to: [github.com/lsgunth/mdbtools-win/releases](https://github.com/lsgunth/mdbtools-win/releases)
2. Download the latest `.zip`
3. Extract it anywhere (e.g. `C:\mdbtools\`)

#### Step 2 — Verify the installation is complete

⚠️ **This step is mandatory.** A common mistake is downloading only the `.exe` files without the full package. An incomplete installation returns blank output with no error.

After extracting, check that your folder contains **at least these files**:

```
mdb-export.exe
mdb-tables.exe
mdb-ver.exe
mdb-count.exe
mdb-schema.exe
mdb-sql.exe
mdb-json.exe
mdb-queries.exe
prcat.exe
prdump.exe
```

If you only have `mdb-export.exe` and `mdb-tables.exe` — **your installation is incomplete**. Re-download the full `.zip` from the releases page.

#### Step 3 — Verify it works on your file

Open PowerShell and run:

```powershell
& "C:\mdbtools\mdb-ver.exe" "C:\path\to\your\file.mdb"
# Must output: JET3
```

If this returns blank or an error, your mdbtools installation is broken — re-download.

```powershell
& "C:\mdbtools\mdb-tables.exe" "C:\path\to\your\file.mdb"
# Must output your table names (e.g. TRAB)
```

If `mdb-ver` shows `JET3` and `mdb-tables` lists tables — you are ready.

#### Step 4 — Batch export all files to CSV

Download [`batch_export.py`](batch_export.py) from this repo. Open it and edit the three paths at the top:

```python
MDBTOOLS    = r"C:\mdbtools"            # folder containing mdb-export.exe
MDB_ROOT    = r"C:\path\to\mdb\files"  # folder containing your .mdb files
OUTPUT_ROOT = r"C:\path\to\output"     # where CSVs will be written
```

Then run:

```
python batch_export.py
```

The script will verify the mdbtools executables exist, check the first file with `mdb-ver.exe`, then export every table from every file — printing progress as it goes and a summary at the end. No Python packages required, uses only the standard library.

---

## Diagnosing Your Files (diagnose.py)

Run the diagnostic script on one sample file before attempting batch export. It reads the raw file header and tells you exactly what you are dealing with.

```powershell
python diagnose.py "C:/path/to/your/file.mdb"
```

**What the header tells you:**

| Header value | Meaning |
|---|---|
| Magic bytes `00 01 00 00` | Valid Jet file |
| Version byte `0x14 = 0` | Jet 3.5 (Access 97) |
| Version byte `0x14 = 1` | Jet 4.0 (Access 2000+) |
| Encryption flag `0x15 = 0` | No encryption |
| Text in password region | Database has a password |
| `"Standard Jet DB"` at offset 4 | Normal unprotected file |

**Note on false positives:** The password detection in `diagnose.py` can produce false positives — it may report a password when there is none. Always try `mdb-tables` without the `-p` flag first. If that returns your table names, no password is needed.

---

## Password-Protected Files

If your files genuinely have a password, pass it to MDBTools:

```powershell
& "C:\mdbtools\mdb-tables.exe" -p "yourpassword" file.mdb
& "C:\mdbtools\mdb-export.exe" -p "yourpassword" file.mdb TableName > output.csv
```

In `batch_export.py`, set the `PASSWORD` variable at the top of the script.

Jet 3.5 stores passwords XOR-encoded at offset `0x18`. The diagnostic script tries all known XOR keys automatically.

---

## Troubleshooting

| Symptom | Cause | Fix |
|---|---|---|
| `mdb-tables` returns blank | Incomplete mdbtools install | Re-download full `.zip` from lsgunth/mdbtools-win |
| `mdb-ver.exe` not found | Same — incomplete install | Same |
| `mdb-tables` returns blank even with full install | File is empty or corrupted | Run `diagnose.py` to inspect the header |
| `HY000: Unrecognized database format` | Used pyodbc on a Jet 3.5 file | Use mdbtools instead |
| `SyntaxError: (unicode error) 'unicodeescape'` | Windows path with backslashes in a string | Use forward slashes or raw strings: `r"C:\path"` |
| `diagnose.py` reports password but mdb-tables works fine | False positive from XOR key matching | Ignore it — no password needed |
| File returns blank even with correct install | Workgroup security (`.mdw` file) | See section below |

---

## What If MDBTools Still Returns Empty?

A small number of Access 97 files use workgroup security (`.mdw` files). Check:
- Is there a `System.mdw` or `*.mdw` file alongside your `.mdb` files?
- If yes: `mdb-tables --sysdb=System.mdw file.mdb`

If none of this works, the nuclear option is a **Windows XP VM** with Access 97 Runtime installed — the original Jet 3.5 engine that created the files will always be able to read them.

---

## Real-World Example

This solution was developed while extracting **1,476 `.mdb` files** (1.35 million rows) from **Marsilli coil winding machines** running Access 97 data loggers from the early 2000s. The files had been locked inside proprietary industrial software for 20+ years.

After all the approaches above failed, `lsgunth/mdbtools-win` exported all 1,476 files in under 5 minutes with zero errors.

---

## Requirements

- Python 3.6+
- No Python packages required for Jet 3.5 — uses only `subprocess` and `pathlib` (standard library)
- For Jet 4.0: `pip install pyodbc pandas`
- [lsgunth/mdbtools-win](https://github.com/lsgunth/mdbtools-win/releases) binaries (Jet 3.5 only)

---

## Files in This Repo

| File | What it does |
|---|---|
| `diagnose.py` | Inspects file headers to detect format, passwords, encryption — run this first |
| `batch_export.py` | Batch exports all tables from all `.mdb` files to CSV — **Jet 3.5 / Access 97** |
| `mdb_to_csv.py` | Batch exports all tables from all `.mdb` files to CSV — **Jet 4.0 / Access 2000+** |
| `README.md` | This file |

---

## Contributing

If you found a case this doesn't handle (different encryption, workgroup security, corrupted files), open an issue or PR. Industrial legacy data is a mess and the more edge cases we document, the more useful this becomes.

---

## License

MIT — do whatever you want with it.
