# AGENTS.md

This repository is an Agent Skill for Microsoft Excel.

- Entry point: `SKILL.md`. Load it when a task involves Excel workbooks, formulas, .xlsx, .xlsm, .xlsb, .xls or .csv files.
- Load files in `references/` only when SKILL.md points to them.
- Scripts in `scripts/` (call them by absolute path, `python <skill-dir>/scripts/NAME.py`) are read-only or write only to new paths. Run `python tests/smoke_test.py` after changing them.
- Contributing: keep SKILL.md under 500 lines, test every code block, cite official docs, and do not use em dashes.
