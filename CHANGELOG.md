# Changelog

## 0.2.2 (2026-10-13)
Fixes from the third review round (12 reviewers).

### Security
- `recalc.py --sandboxed` is now documented as operator-only in SKILL.md, security.md, AGENTS.md, README and the script help: an agent must never pass it on its own judgment or to get past a refusal. Under `--sandboxed` the script first tries 1-second TCP connects to public addresses and refuses (exit 2) if any succeeds.
- `recalc.py` runs LibreOffice on a private copy of the input, with HOME, TMPDIR and the working directory in a temp folder, so relative link paths cannot reach the user's folders. security.md documents that a network namespace does not hide local files or filesystem sockets, and recommends a container or bwrap that mounts only the input. Tested: `WEBSERVICE` of `file://` and relative paths, `FILTERXML`, and `INDIRECT` of a `file://` target bring no file content into the output.
- On timeout `recalc.py` kills the whole LibreOffice process group, not just the first process.
- Refusals and setup errors now exit 2 (they exited 1 before). The no-isolation error mentions AppArmor on Ubuntu 24.04+ and tells agents not to add `--sandboxed`.

### Added
- `.xlsm` recalc warning (SKILL.md section 5, security.md, script help): LibreOffice may not keep the VBA project, form controls or ActiveX byte for byte. `recalc.py` prints `vba_project=identical|changed|missing` and exits 1 on changed or missing.
- `scripts/requirements.txt` (openpyxl, defusedxml).
- security.md: static inspection, library edits and recalculation as three separate risk levels.
- Tests (31 run, 1 skipped where `unshare` works): `--sandboxed` refusal when a connection works, `--sandboxed` inside a no-network namespace, local-file canary, `.xlsm` VBA status, `vba_status`, help text, all CSV formula triggers stay literal, literal `=text` survives reload, package name and version consistency, package files. The test runner prints SKIP lines and a summary.

### Fixed
- recipes.md contradicted itself on charts (table: kept but rebuilt; chart section: dropped). Both now say kept but rebuilt, as tested, with a caveat that Excel-authored charts may lose more. The matrix states it was tested on an xlsxwriter-authored file, not an Excel-authored one.
- Pivot wording softened: openpyxl has limited read and preservation support, but no reliable editing API.
- recipes.md pivot section: GROUPBY and PIVOTBY become `#NAME?` under `recalc.py`.
- ooxml.md: built-in number formats 18 to 21 are times; remove `calcChain.xml` only when it is corrupt.

### Changed
- npm package: `@skillsaas/microsoft-excel` with `publishConfig.access` public. `files` now lists README, AGENTS.md, LICENSE and `scripts/requirements.txt`. `engines.node` removed (nothing runs on Node).
- SKILL.md: shorter description with a "not needed" line; ExcelJS dropped from description, tags, table and README (no recipe covered it); `metadata.last-verified`; LibreOffice "7.6+ (tested 25.2)"; inspector exit codes and JSON contents without `--show-links`; a failed baseline inspection makes preservation inconclusive; inventory comparison is triage, not proof; `.xlsm` checklist item; `read_only`/`write_only` limits; compare with real Excel for important results; the "any format" row qualified.
- formulas.md PIVOTBY wording follows Microsoft's documented versions with a verify-first note. CSV apostrophe prefix called a display-level mitigation.
- README: `npm install` line, catalog marked planned, `--sandboxed` note.

## 0.2.1 (2026-10-13)
Fixes from the second review round (12 reviewers).

### Security
- `recalc.py` no longer reaches the network. The throwaway profile sets Calc link updates to never and blocks untrusted referer links. On Linux, LibreOffice runs under `unshare -n` or `unshare -rn` (new network namespace, no network); where that is unavailable the script refuses unless `--sandboxed` is passed. Tested: a `WEBSERVICE` call to a local server makes no request, and a socket connect inside the namespace fails.
- `recalc.py` passes LibreOffice a minimal environment, so cloud tokens in the caller's environment are not inherited.
- `inspect_xlsx.py` `clean()` also strips C1 controls (including `\x9b` CSI), bidi embeddings, overrides and isolates, zero-width characters, word joiner and BOM.
- `inspect_xlsx.py --show-links` parse errors give a clean exit 2 instead of a traceback.

### Fixed
- `recalc.py` no longer deletes an existing `OUTDIR/<name>`. It refuses unless `--force` is given, converts into a temp folder, and moves the result only after LibreOffice succeeds.
- `recalc.py` verification matches cells by coordinate (no `zip()` truncation) and counts formulas by `data_type == "f"`, so literal `=text` is not a formula.
- `recalc.py` reports a clean exit 2 when LibreOffice cannot start, instead of a traceback.
- `csv_to_xlsx.py` keeps newlines inside quoted fields (it used `splitlines()`, which joined `"a\r\nb"` into `ab`).
- recipes.md: the `sheet_name=None` pandas snippet returned a dict and then indexed it like a DataFrame. Now it picks a sheet (or concatenates).

### Changed
- SKILL.md format table: the openpyxl preservation claim is now conservative, with a tested round-trip table in recipes.md (openpyxl 3.1.5, xlsxwriter 3.2.9, Pillow; Oct 13, 2026). Cached values are dropped on save; charts are rebuilt; pivots and slicers are untested.
- `--show-links` described the same way everywhere: the inspector always counts external links; targets need `--show-links`. JSON also lists the link part names.
- Script paths in SKILL.md, AGENTS.md and security.md are `<skill-dir>/scripts/...`, so they resolve from the skill folder instead of the user's working directory.
- Workflow: inspect to `before.json`, inspect the output to `after.json`, and explain every unintended difference.
- LibreOffice stale-cache sentence turned into a verification rule; a cached value is not proof the result matches Excel.
- GROUPBY and PIVOTBY wording in SKILL.md aligned with formulas.md. Magic-byte check says `PK` is only a zip container. `.xlsm` row says `keep_vba=True` is needed at load time. Literal-text example is complete. CSV `'` prefix marked as CSV-only.
- README: the unpublished `@skillsaas/cli` line is marked as planned; install note on folder name and script paths. Version headers in recipes.md say pandas 2.2+ or 3.x.
- Tests: 11 new checks in `tests/smoke_test.py` (22 in total; the "refuses without network isolation" check runs only where `unshare` is unavailable, the two network checks only where it works).

## 0.2.0 (2026-10-13)
Reworked from reviews by 10 reviewers (8 AI models plus a platform and a security reviewer).

### Added
- Decision guide at the top of SKILL.md (task, format sniffing, Excel availability, reader version), plus a format capability table.
- Standard workflow: inspect, copy, minimal change, recalculate, verify, report.
- Verification checklist, and a table of common symptoms and fixes.
- Untrusted-file rules in SKILL.md, and `references/security.md`: formula and CSV injection, macros, external data, sandboxed recalculation, Graph scopes.
- `references/recipes.md`: pivot summaries (and where real pivots need Excel), charts, data validation, conditional formatting, named ranges, merged cells, protection, encrypted files (msoffcrypto-tool), pandas without data loss, xlsxwriter, calamine.
- `references/automation.md`: Office.js, Office Scripts (with guards), Graph (sessions, scopes, recalculation, throttling), COM and xlwings with macros disabled, VBA, Power Query.
- `scripts/recalc.py`: forced LibreOffice recalculation in a throwaway profile, macros disabled, timeout, separate output folder, built-in verification. The plain `soffice --convert-to` approach was shown to keep stale values.
- `scripts/csv_to_xlsx.py`: text-safe conversion with encoding fallback, delimiter sniffing, a decimal-comma option, injection protection, and no overwrite.
- `tests/smoke_test.py` (repo only).
- `AGENTS.md`, plus discovery metadata in README and package.json keywords.

### Changed
- Frontmatter follows the Agent Skills spec: `name`, `description`, `license` and `compatibility`, with project fields (version, updated, risk, formats, software versions, tags) under `metadata`.
- `risk` raised from `safe` to `medium`, with notes.
- Function availability qualified by version (GROUPBY, PIVOTBY and REGEX* depend on the update channel).
- Date guidance: use the workbook epoch and library conversion, and take care with dates before 1900-03-01 (the fake 1900-02-29).
- Dynamic arrays: one recommended path (write through Excel); the `ArrayFormula` workaround removed.
- openpyxl example cleaned up: no unused imports, the table is named `tblSales` and created before the formula that refers to it.
- SKILL.md is now a router of about 140 lines; VBA, Graph, Office.js and Power Query moved to references.
- `inspect_xlsx.py` hardened:
  - requires defusedxml;
  - checks the zip size before loading;
  - does one full load plus one streaming pass for cached values;
  - sanitizes and labels file content;
  - hides external link targets unless `--show-links` is given;
  - adds `--json` output and clean errors for OLE2 (encrypted or .xls), .xlsb, corrupt zips and missing parts;
  - reports hidden sheets, rows and columns, merged ranges, validation, conditional formats, protection, charts, pivots and connections.

### Fixed
- Literal text starting with `=` is no longer counted as a formula by the inspector.

## 0.1.0 (2026-10-13)
- Initial hand-written skill.
