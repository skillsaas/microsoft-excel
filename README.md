# Microsoft Excel agent skill

**TL;DR:** `microsoft-excel` is a free, MIT-licensed agent skill that teaches AI agents (Claude Code, Cursor, Codex, Gemini CLI and other agents that read SKILL.md skills) to read, create, edit, recalculate and verify Excel workbooks without breaking them, with openpyxl, xlsxwriter, pandas, Office.js, Office Scripts, VBA and Microsoft Graph. Install: `npx skills add skillsaas/microsoft-excel` or `npm install microsoft-excel`.

`microsoft-excel` is an [Agent Skills](https://agentskills.io/specification) skill that teaches AI coding agents (Claude Code, Cursor, Codex, Gemini CLI and others) to read, create, edit, recalculate and verify Excel workbooks without breaking them.

Unofficial. Microsoft and Excel are trademarks of Microsoft Corporation. This project is not affiliated with or endorsed by Microsoft.

## Use it when
- Writing or debugging Excel formulas (XLOOKUP, dynamic arrays, LAMBDA, SUMIFS).
- Creating or editing .xlsx, .xlsm or .csv files with openpyxl, xlsxwriter or pandas when Excel is not installed.
- Getting calculated values without Excel (forced LibreOffice recalculation).
- Automating Excel with Office.js, Office Scripts, VBA, COM/xlwings or Microsoft Graph.
- Fixing wrong totals, lost leading zeros, shifted dates, CSV import damage, `#NAME?` or `#SPILL!`.
- Handling workbooks from untrusted sources safely.

## Install
```bash
npx skills add skillsaas/microsoft-excel            # vercel-labs skills CLI, reads SKILL.md at the repo root
npm install microsoft-excel              # files land in node_modules/microsoft-excel/
# Planned, not yet published: npx @skillsaas/cli add microsoft-excel --claude
```
Manual: copy `SKILL.md`, `references/` and `scripts/` into your agent's skills folder, for example `~/.claude/skills/microsoft-excel/`. Keep the folder name `microsoft-excel`. Typical locations: Claude Code `~/.claude/skills/`, other agents in their own skills folder (check your agent's docs). SKILL.md calls scripts as `python <skill-dir>/scripts/NAME.py`, so they work whatever the agent's working directory is.

## What it does
- `inspect_xlsx.py`: read-only workbook inventory with `--json` (defusedxml, size cap, sanitized output).
- `recalc.py`: forced LibreOffice recalculation in a throwaway profile (macros off, link updates never, no network), then a check of the results.
- `csv_to_xlsx.py`: CSV to `.xlsx` with text-safe columns (leading zeros kept) and formula-injection protection.
- References for formulas by Excel version, error codes, pivots, charts, validation, Office.js, Office Scripts, VBA, Power Query and OOXML repair.
- Python 3.9+ with openpyxl and defusedxml; Excel itself is optional.

## FAQ
**What is the microsoft-excel skill?**
An Agent Skills package (a `SKILL.md` plus references and Python scripts) that tells an AI agent how to work with Excel files and automation correctly. It is unofficial and not affiliated with Microsoft.

**Why does openpyxl return None for formula cells?**
openpyxl reads the value cached by the last app that calculated the file; a file written by openpyxl has none. `recalc.py` recalculates it with LibreOffice so the values can be read.

**How do I stop Excel losing leading zeros in a CSV?**
Convert it with `csv_to_xlsx.py`, which writes every cell as text unless its column is listed as a number or date column, instead of opening the CSV directly.

**Do I need Excel installed?**
No. The scripts use Python; `recalc.py` needs LibreOffice. Excel is only needed for features LibreOffice does not calculate the same way.

**How do I install it in Claude Code?**
`npx skills add skillsaas/microsoft-excel`, or copy `SKILL.md`, `references/` and `scripts/` to `~/.claude/skills/microsoft-excel/`.

**Is it safe with workbooks from strangers?**
The inspector only reads. `references/security.md` covers macros, external data links and formula injection, and `recalc.py` runs LibreOffice with macros and link updates off and no network.

## Contents
| Path | Purpose |
| :--- | :--- |
| `SKILL.md` | Decision guide, workflow, security rules, hard limits, verification checklist, symptom table |
| `references/formulas.md` | Function choice by Excel version, error codes with fixes, input and expected output |
| `references/recipes.md` | Pivots, charts, data validation, conditional formatting, names, merged cells, protection, encryption, pandas, xlsxwriter |
| `references/automation.md` | Office.js, Office Scripts, Graph, COM/xlwings, VBA, Power Query |
| `references/security.md` | Untrusted files, macros, external data, formula injection, parsing safety |
| `references/ooxml.md` | .xlsx internals and repair |
| `scripts/inspect_xlsx.py` | Read-only inventory, `--json`, hardened (defusedxml, size cap, sanitized output) |
| `scripts/recalc.py` | Forced LibreOffice recalculation in a throwaway profile (macros off, link updates never, no network), then verification |
| `scripts/csv_to_xlsx.py` | CSV to .xlsx with text-safe columns and formula-injection protection |

Script dependencies: Python 3.9+, `pip install -r scripts/requirements.txt` (openpyxl, defusedxml), plus LibreOffice for `recalc.py`. The tests in `tests/smoke_test.py` are in the repo only, not in the npm package; run `python tests/smoke_test.py` (it prints PASS, FAIL and SKIP lines and a summary).

## Discovery metadata
- Skill name: `microsoft-excel` (same as the folder, the GitHub repo `skillsaas/microsoft-excel` and the npm package `microsoft-excel`).
- Formats: xlsx, xlsm, xlsb, xls, csv. Tested with openpyxl 3.1.5, xlsxwriter 3.2.9, LibreOffice 25.2, Python 3.13; pandas 3.0.
- Catalog of all skillsaas skills (planned, not yet published): https://github.com/skillsaas/index
- recalc.py `--sandboxed` is for operators only: agents must never pass it themselves (`references/security.md`).
- GitHub topics: `agent-skill`, `agent-skills`, `skill-md`, `claude-code-skill`, `excel`, `xlsx`, `openpyxl` and more.

## License
MIT
