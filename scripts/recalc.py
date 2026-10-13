#!/usr/bin/env python3
"""Force a full recalculation of a workbook with headless LibreOffice, then verify.

Usage:
  python recalc.py INPUT.xlsx OUTDIR [--timeout 180] [--soffice PATH] [--force] [--sandboxed]

Writes OUTDIR/<input name> with fresh cached values. Never touches INPUT, and
refuses to replace an existing OUTDIR/<input name> unless --force is given.
A plain LibreOffice load may keep the cached values stored in the file; this
script uses a throwaway profile that sets "recalculate on load: always",
"update links: never" and disables macros, then reports formulas still missing
a cached value and error cells.

Network: a forced recalc evaluates WEBSERVICE and similar functions, so a
hostile workbook could send cell data to the internet. On Linux the script runs
LibreOffice inside a new network namespace with no network (`unshare -rn`).
Where that is unavailable (macOS, Windows, locked-down containers) it refuses to
run unless you pass --sandboxed, which asserts that you already run it in a
sandbox without network access (for example `docker run --network none`).

Limits: LibreOffice is a different engine from Excel. GROUPBY, PIVOTBY, REGEX*
and some LAMBDA helpers return #NAME?. Say which engine produced the numbers.
A cached value being present does not prove the result matches Excel.
Exit codes: 0 ok, 1 recalculated but errors or missing values, 2 failure.
Requires: LibreOffice (soffice), pip install openpyxl defusedxml
"""
import argparse
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

XCU = """<?xml version="1.0" encoding="UTF-8"?>
<oor:items xmlns:oor="http://openoffice.org/2001/registry" xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<item oor:path="/org.openoffice.Office.Calc/Formula/Load"><prop oor:name="OOXMLRecalcMode" oor:op="fuse"><value>0</value></prop></item>
<item oor:path="/org.openoffice.Office.Calc/Formula/Load"><prop oor:name="ODFRecalcMode" oor:op="fuse"><value>0</value></prop></item>
<item oor:path="/org.openoffice.Office.Calc/Content/Update"><prop oor:name="Link" oor:op="fuse"><value>1</value></prop></item>
<item oor:path="/org.openoffice.Office.Common/Security/Scripting"><prop oor:name="MacroSecurityLevel" oor:op="fuse"><value>3</value></prop></item>
<item oor:path="/org.openoffice.Office.Common/Security/Scripting"><prop oor:name="DisableMacrosExecution" oor:op="fuse"><value>true</value></prop></item>
<item oor:path="/org.openoffice.Office.Common/Security/Scripting"><prop oor:name="BlockUntrustedRefererLinks" oor:op="fuse"><value>true</value></prop></item>
</oor:items>
"""
# Calc Content/Update/Link: 0 always, 1 never, 2 on request (LibreOffice default 2).
ERRORS = {"#N/A", "#VALUE!", "#REF!", "#NAME?", "#DIV/0!", "#NUM!", "#NULL!", "#SPILL!", "#CALC!", "Err:502", "Err:504"}
FILTERS = {".xlsx": "xlsx:Calc MS Excel 2007 XML", ".xlsm": "xlsm:Calc MS Excel 2007 VBA XML"}


def find_soffice(explicit):
    for c in [explicit, shutil.which("soffice"), shutil.which("libreoffice"),
              "/Applications/LibreOffice.app/Contents/MacOS/soffice",
              r"C:\Program Files\LibreOffice\program\soffice.exe"]:
        if c and Path(c).exists():
            return c
    return None


def no_network_prefix():
    """Return a command prefix that runs a child with no network, or None."""
    unshare = shutil.which("unshare")
    if not unshare or not sys.platform.startswith("linux"):
        return None
    for prefix in ([unshare, "-n"], [unshare, "-rn"]):
        try:
            if subprocess.run(prefix + ["true"], capture_output=True, timeout=10).returncode == 0:
                return prefix
        except (OSError, subprocess.TimeoutExpired):
            pass
    return None


def verify(path):
    """Count formulas, formulas without a cached value, and error cells.
    Cells are matched by coordinate, so differing row lengths cannot hide cells."""
    import defusedxml  # noqa: F401  safe XML parsing in openpyxl
    from openpyxl import load_workbook
    wf = load_workbook(path, read_only=True)
    wv = load_workbook(path, read_only=True, data_only=True)
    missing = errors = formulas = 0
    for ws in wf.worksheets:
        values = {}
        for row in wv[ws.title].iter_rows():
            for c in row:
                if getattr(c, "row", None) is not None and c.value is not None:
                    values[(c.row, c.column)] = c.value
        for row in ws.iter_rows():
            for c in row:
                if getattr(c, "row", None) is None:
                    continue
                v = values.get((c.row, c.column))
                if c.data_type == "f":
                    formulas += 1
                    if v is None:
                        missing += 1
                if isinstance(v, str) and v in ERRORS:
                    errors += 1
    wf.close()
    wv.close()
    return formulas, missing, errors


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("input")
    ap.add_argument("outdir")
    ap.add_argument("--timeout", type=int, default=180)
    ap.add_argument("--soffice")
    ap.add_argument("--force", action="store_true", help="replace an existing OUTDIR/<input name>")
    ap.add_argument("--sandboxed", action="store_true",
                    help="run without unshare; you assert the process already has no network access")
    a = ap.parse_args()

    src = Path(a.input).resolve()
    out = Path(a.outdir).resolve()
    if not src.is_file():
        sys.exit(f"error: {src} not found")
    if src.suffix.lower() not in FILTERS:
        sys.exit("error: only .xlsx and .xlsm are supported")
    if out == src.parent:
        sys.exit("error: OUTDIR must differ from the input folder so the original is never overwritten")
    dest = out / src.name
    if dest.exists() and not a.force:
        sys.exit(f"error: {dest} already exists; choose another OUTDIR or pass --force")
    soffice = find_soffice(a.soffice)
    if not soffice:
        sys.exit("error: LibreOffice not found (install it or pass --soffice)")
    prefix = [] if a.sandboxed else no_network_prefix()
    if prefix is None:
        sys.exit("error: cannot disable network for LibreOffice here (unshare unavailable). Run inside a "
                 "sandbox with no network, for example docker --network none, and pass --sandboxed")

    with tempfile.TemporaryDirectory(prefix="xlrecalc-") as prof:
        (Path(prof) / "user").mkdir()
        (Path(prof) / "user" / "registrymodifications.xcu").write_text(XCU, encoding="utf-8")
        tmp_out = Path(prof) / "out"
        tmp_out.mkdir()
        cmd = prefix + [soffice, f"-env:UserInstallation={Path(prof).as_uri()}", "--headless", "--norestore",
                        "--nolockcheck", "--convert-to", FILTERS[src.suffix.lower()],
                        "--outdir", str(tmp_out), str(src)]
        env = {"PATH": os.environ.get("PATH", ""), "HOME": prof, "TMPDIR": prof,
               "LANG": os.environ.get("LANG", "C.UTF-8"), "SAL_USE_VCLPLUGIN": "svp"}
        if os.name == "nt":
            env.update({k: os.environ[k] for k in ("SYSTEMROOT", "TEMP", "TMP") if k in os.environ})
        try:
            p = subprocess.run(cmd, capture_output=True, text=True, timeout=a.timeout, env=env)
        except subprocess.TimeoutExpired:
            sys.stderr.write(f"error: LibreOffice timed out after {a.timeout}s\n")
            return 2
        except OSError as e:
            sys.stderr.write(f"error: could not start LibreOffice: {e}\n")
            return 2
        produced = tmp_out / src.name
        if p.returncode != 0 or not produced.exists():
            sys.stderr.write(p.stdout + p.stderr)
            return 2
        out.mkdir(parents=True, exist_ok=True)
        shutil.move(str(produced), str(dest))  # only now is an existing file replaced (with --force)

    formulas, missing, errors = verify(dest)
    print(f"recalculated with LibreOffice (network {'sandbox asserted' if a.sandboxed else 'disabled'}) -> {dest}")
    print(f"formulas={formulas} missing_cached_value={missing} error_cells={errors}")
    return 1 if (missing or errors) else 0


if __name__ == "__main__":
    sys.exit(main())
