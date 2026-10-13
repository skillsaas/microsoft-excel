#!/usr/bin/env python3
"""Convert a CSV to .xlsx without Excel's silent corruption, and without formula injection.

Usage:
  python csv_to_xlsx.py IN.csv OUT.xlsx [--text-cols "ID,Zip"] [--number-cols "Amount"]
                        [--date-cols "Date"] [--date-format %Y-%m-%d] [--delimiter ,]
                        [--decimal .] [--encoding auto] [--allow-formulas]

- Every cell is written as text unless its column is listed in --number-cols or
  --date-cols, so leading zeros, long IDs and gene names survive.
- Encoding "auto" tries utf-8-sig, then utf-8, then cp1252. The delimiter is
  sniffed if not given.
- Values starting with = + - @ tab or CR are stored as literal text (cell type
  "s"), so they never run as formulas. --allow-formulas disables this, for
  trusted input only.
- Refuses to overwrite an existing OUT file.
Requires: pip install openpyxl
"""
import argparse
import csv
import datetime as dt
import io
import sys
from pathlib import Path

from openpyxl import Workbook

RISKY = ("=", "+", "-", "@", "\t", "\r")


def read_text(path, enc):
    data = Path(path).read_bytes()
    for e in ([enc] if enc != "auto" else ["utf-8-sig", "utf-8", "cp1252"]):
        try:
            return data.decode(e), e
        except UnicodeDecodeError:
            continue
    sys.exit("error: could not decode; pass --encoding")


def cols(arg):
    return {c.strip() for c in arg.split(",") if c.strip()} if arg else set()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("src")
    ap.add_argument("dst")
    ap.add_argument("--text-cols", default="")
    ap.add_argument("--number-cols", default="")
    ap.add_argument("--date-cols", default="")
    ap.add_argument("--date-format", default="%Y-%m-%d")
    ap.add_argument("--delimiter")
    ap.add_argument("--decimal", default=".", choices=[".", ","], help="decimal separator in number columns")
    ap.add_argument("--encoding", default="auto")
    ap.add_argument("--allow-formulas", action="store_true")
    a = ap.parse_args()
    if Path(a.dst).exists():
        sys.exit(f"error: {a.dst} exists; choose a new name")

    text, enc = read_text(a.src, a.encoding)
    delim = a.delimiter
    if not delim:
        try:
            delim = csv.Sniffer().sniff(text[:65536], delimiters=",;\t|").delimiter
        except csv.Error:
            delim = ","
    rows = list(csv.reader(io.StringIO(text, newline=""), delimiter=delim))  # keeps newlines inside quoted fields
    if not rows:
        sys.exit("error: empty CSV")
    header = rows[0]
    num, dates = cols(a.number_cols), cols(a.date_cols)
    unknown = (num | dates | cols(a.text_cols)) - set(header)
    if unknown:
        sys.exit(f"error: columns not in header: {sorted(unknown)}")

    wb = Workbook()
    ws = wb.active
    ws.title = "Data"
    escaped = bad = 0
    for r, row in enumerate(rows, 1):
        for c, raw in enumerate(row, 1):
            name = header[c - 1] if c <= len(header) else ""
            cell = ws.cell(r, c)
            val = raw
            if r > 1 and name in num and raw.strip():
                try:
                    thousands = "," if a.decimal == "." else "."
                    val = float(raw.strip().replace(thousands, "").replace(a.decimal, "."))
                except ValueError:
                    bad += 1
            elif r > 1 and name in dates and raw.strip():
                try:
                    val = dt.datetime.strptime(raw.strip(), a.date_format).date()
                    cell.number_format = "yyyy-mm-dd"
                except ValueError:
                    bad += 1
            cell.value = val
            if isinstance(val, str):
                if val.startswith(RISKY) and not a.allow_formulas:
                    escaped += 1
                if not a.allow_formulas:
                    cell.data_type = "s"  # literal text even if it starts with "="
                cell.number_format = "@"
    ws.freeze_panes = "A2"
    wb.save(a.dst)
    print(f"wrote {a.dst}: rows={len(rows)} encoding={enc} delimiter={delim!r} "
          f"formula_like_values_kept_as_text={escaped} unparsed_number_or_date_cells={bad}")


if __name__ == "__main__":
    main()
