# Changelog

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
