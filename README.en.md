# wind-edb

[中文](README.md) | **English**

An AI agent skill for working with **Wind Financial Terminal's Excel EDB (economic database) connections**: write, refresh, query by exact statistical date, and verify what was saved.

On the tested setup, Wind's EDB formula only works as a no-argument call, `='WindFunc.xla'!edb()`. The indicator, the date window and the export layout all live in the **legacy cell comment** attached to that cell: Base64 → a ZIP local header with no central directory → `EDB\0` + a version word + UTF-8 JSON. This skill documents that undocumented format and ships offline inspection, encode/decode, metric-pool, native-refresh and PowerShell write tools.

## What it does

- **Offline inventory** of every EDB / WSET anchor in a workbook — decodes comment configs without launching Excel and without modifying the file
- **Cross-workbook metric pool**: batch-decode many workbooks into a searchable index of code / name / unit / freq
- **Dated queries**: a single observation, a date range, or the last N observations
- **Native write and refresh**: write the comment, write the formula, calculate, read back the real used range, then re-verify the connection survived the save
- **Failure triage**: tell apart "the add-in isn't loaded", "the comment is gone" and "the code is wrong" — three completely different failures that look alike

## Install

Drop the `wind-edb` folder into your agent's skill directory:

- Codex: `~/.codex/skills/wind-edb`
- Other agents that support `SKILL.md`: follow that agent's install convention

You can also just point the agent at `wind-edb/SKILL.md`.

## Requirements

- Windows
- Excel
- Wind Financial Terminal, logged in with data permission (tested on Wind 26.6.1)
- Python 3 (the offline tools use only the standard library)
- `pywin32` (only for `probe_excel.py`; use the bundled `edb_write.ps1` if it is missing)

## Quick start

```powershell
# 1. see what a workbook already has
python scripts/edb_tools.py inspect workbook.xlsx --out inspect.json

# 2. build a metric pool across several workbooks and find a code
python scripts/pool_metrics.py --roots "<metrics-dir>" --terms consumption retail --out pool.json

# 3. prepare a dated query from a verified anchor
python scripts/workbook_query.py prepare workbook.xlsx --sheet EDB_raw --anchor A4 `
    --codes M0001427 --start 2024-01-01 --end 2024-12-31 --output query.json

# 4. write and refresh natively, then verify
pwsh -File scripts/edb_write.ps1 -CommentFile cmt.txt -Anchor A4 -Workbook workbook.xlsx -Sheet EDB_raw -Save
```

`cmt.txt` is the output of `edb_tools.pack_config(config)` — that Base64 comment.

## Traps worth knowing

| Symptom | What is actually happening |
|---|---|
| The formula returns `Wind` | That is the **header row**, not an error code |
| `=EDB("code","start","end")` returns `#VALUE!` | This add-in build does not accept the parameterised form at all; it says nothing about whether the code is valid |
| You changed the code but the old indicator name still shows | The "indicator name" row echoes the `name` **you passed in**. It cannot validate a code — only the magnitude of the values can |
| The comment is there but refresh returns nothing | `WindFunc.xla` was never opened as an add-in, so the macro never ran |
| The connection vanished after editing with openpyxl | Generic spreadsheet writers drop comments and external links; you must go through Excel COM |

## Finding indicator codes

Indicator codes are **not** in the Wind installation directory. `etc/edb/*.chs` is a token index, `Indicator.xml` is encrypted, `wind_IndicatorTree.dat` is compressed, `formulaBasicIndicator.json` only holds fundamentals, and `WindFunc.xla` is a bare add-in. None of them yield macro EDB codes.

Your own historical workbooks are the useful source: one `pool_metrics.py` pass over a few of them yields hundreds of indicators in seconds. Only what is left after that is worth looking up in Wind's indicator browser.

## Layout

```
wind-edb/
├── SKILL.md                      main skill instructions
├── agents/openai.yaml            platform interface descriptor
├── references/
│   ├── native-workflow.md        connection format, write order, failure table
│   ├── specific-data.md          dated queries, date semantics, report-cell lookup
│   └── metric-pool.md            metric pool, code reference, probing, cross-checks
└── scripts/
    ├── edb_tools.py              offline inspect / config codec / CRC checks
    ├── pool_metrics.py           build a metric pool across workbooks
    ├── prepare_query.py          build a dated query from a verified template
    ├── workbook_query.py         search a workbook, then build a query
    ├── probe_excel.py            fresh-workbook probe (needs pywin32)
    └── edb_write.ps1             equivalent write path in PowerShell
```

## Scope and limits

- The connection format is an **empirically compatible** format, not a documented, stable Wind protocol. Re-probe after an add-in upgrade or on a different machine.
- No Wind licence, commercial data or user workbook is included.
- No local paths, accounts or machine identifiers are included; example paths use `<metrics-dir>` and `<Wind-install-dir>` placeholders.
- Business delivery still requires you to check the statistical period, unit, frequency, indicator definition and the latest published vintage yourself.

## License

[MIT](LICENSE) © 2026 MlittleFriend

Anyone may use, modify and redistribute this project, including commercially, provided the copyright and permission notices are retained in copies or substantial portions. The Chinese note in `README.md` is for convenience only; the English text in [LICENSE](LICENSE) is authoritative.

Wind is a trademark of Wind Information Co., Ltd. This project is not affiliated with it.
