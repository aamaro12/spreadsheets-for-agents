# Spreadsheets for agents

`sheetscope` lets an AI agent read an Excel workbook the way an analyst does: structure first, then search, then exact cells. Every value comes with the cell it came from, its formula and the labels around it, so an answer can cite `[model.xlsx]Summary!B4` and a reviewer can check it.

Pasting a workbook into a prompt loses what makes it a model: the formulas, the number formats, the row a figure sits on and the year at the top of its column. Reading it one cell at a time takes too many calls. `sheetscope` gives an agent 5 steps instead, as a command line tool and as an MCP server.

| Step | What it returns |
|---|---|
| `outline` | Sheets, with hidden ones flagged; blocks of cells with their header rows and row labels; defined names with their values; tables; warnings |
| `find` | Cells that match some text or a number, allowing for rounding, for how a cell shows it and for the units its table states: `1.4` finds 1.38, `12.5%` finds 0.1249, `1,597` finds 1,597,000 shown in thousands and `60.9bn` finds 60,922 in a table in millions |
| `read` | A range as a grid, with the labels and headers just outside it, and its formulas grouped by the ranges they were filled across |
| `cell` | One cell in full: shown and stored value, number format, formula, the cells it reads and the formulas that use it |
| `trace` | The tree of cells behind a figure, down to its inputs, or the formulas that depend on it |

## Install

Python 3.11 or later.

```
git clone https://github.com/aamaro12/spreadsheets-for-agents
cd spreadsheets-for-agents
pip install -e ".[mcp]"
```

Leave out `[mcp]` if you only need the command line.

## Try it

`examples/northwind.xlsx` is a small operating model for a made-up company, with a summary, a model, inputs, a table of customers and a hidden scratch sheet.

```
$ sheetscope outline examples/northwind.xlsx
**northwind.xlsx** · sha256 3aad242a1758 · 5 sheets
Cite cells like [northwind.xlsx]Summary!A1.

## Model · A1:E12 · 45 cells · 31 formulas
- A1:E12 · Northwind Trading: operating model · header row 3: A £k · B FY2024 · C FY2025 · D FY2026 · E FY2027
  rows: 4 Revenue · 5 Growth · 6 Gross profit · 7 Operating costs · 8 EBIT · 9 EBIT margin · 11 Tax · 12 Net income
...
## Scratch · hidden · A1:B2 · 3 cells
...
Names: Growth = Inputs!$B$4 (10.0%) · GrossMargin = Inputs!$B$5 (42.0%) · Inflation = Inputs!$B$7 (3.0%) · TaxRate = Inputs!$B$8 (25%)

Tables: Customers = Customers!A1:C6 (Customer, Region, Revenue FY2024)
```

Find a figure quoted in a report, even when the cell holds more decimals than the page shows:

```
$ sheetscope find examples/northwind.xlsx 1,597
| Cell | Shows | Row | Column | Formula |
|---|---|---|---|---|
| Summary!B4 | 1,597 | Revenue | FY2027 | =Model!E4 |
| Model!E4 | 1,597 | Revenue | FY2027 | =D4*(1+Growth) |
```

Or as a report writes it, in the units the table states (the model's header says £k):

```
$ sheetscope find examples/northwind.xlsx 1.6m
| Cell | Shows | Row | Column | Formula |
|---|---|---|---|---|
| Model!E4 | 1,597 (in thousands) | Revenue | FY2027 | =D4*(1+Growth) |
```

Read a slice without losing its headers. A formula filled across a row is shown once:

```
$ sheetscope read examples/northwind.xlsx Model!C4:E9
|   | C · FY2025 | D · FY2026 | E · FY2027 |
|---|---|---|---|
| 4 · Revenue | 1,320 | 1,452 | 1,597 |
| 5 · Growth | 10.0% | 10.0% | 10.0% |
| 6 · Gross profit | 554 | 610 | 671 |
| 7 · Operating costs | 319 | 329 | 339 |
| 8 · EBIT | 235 | 281 | 332 |
| 9 · EBIT margin | 17.8% | 19.3% | 20.8% |

Formulas, each written for the first cell of its range:
- C4:E4: =B4*(1+Growth)
- C5:E5: =C4/B4-1
...
```

Look at one cell, and follow a figure back to its inputs:

```
$ sheetscope cell examples/northwind.xlsx Model!D4
Model!D4 shows 1,452

- Value: 1452.0000000000002 (number, format #,##0;(#,##0))
- Formula: =C4*(1+Growth)
- Row: Revenue · Column: FY2026
- Reads: Model!C4 (Revenue, FY2025) = 1,320; Growth → Inputs!B4 (Revenue growth, Value) = 10.0%
- Used by: Summary!D4, Model!E4, Model!D5, Model!E5, Model!D6, Model!D9
- Cite as: [northwind.xlsx]Model!D4

$ sheetscope trace examples/northwind.xlsx Summary!B4
- Summary!B4 (Revenue, FY2027) = 1,597 · =Model!E4
  - Model!E4 (Revenue, FY2027) = 1,597 · =D4*(1+Growth)
    - Model!D4 (Revenue, FY2026) = 1,452 · =C4*(1+Growth)
      - Model!C4 (Revenue, FY2025) = 1,320 · =B4*(1+Growth)
        - Model!B4 (Revenue, FY2024) = 1,200 · =SUM(Customers[Revenue FY2024])
          - Customers[Revenue FY2024] → Customers!C2:C6 (5 cells, 0 formulas): 310, 285, 240, 205, 160
        - Growth → Inputs!B4 (Revenue growth, Value) = 10.0%
...
```

References can be ranges (`Model!B3:F12`), whole sheets (`Model`), defined names (`Growth`), table references (`Customers[Revenue FY2024]`) or citations (`[northwind.xlsx]Model!D4`, or `'[model.xlsx]P&L 2025'!B4` and `[model.xlsx]'P&L 2025'!B4` for a sheet name with spaces). Add `--json` to any step for the full result as JSON.

## Use it from an agent

`sheetscope mcp --root FOLDER` serves the steps as read-only MCP tools, plus `workbooks` to list the files. The server reads workbooks under that folder and nowhere else. Its instructions tell the agent to start with `outline` and to cite every figure by its cell.

For Codex, in `~/.codex/config.toml`:

```toml
[mcp_servers.sheetscope]
command = "sheetscope"
args = ["mcp", "--root", "/path/to/workbooks"]
```

For clients configured in JSON:

```json
{"mcpServers": {"sheetscope": {"command": "sheetscope", "args": ["mcp", "--root", "/path/to/workbooks"]}}}
```

Use the full path to `sheetscope` if it is not on the client's `PATH`.

From Python, each step returns a plain dict:

```python
from sheetscope import load

book = load("examples/northwind.xlsx")
book.find("1,597")["matches"][0]["citation"]  # '[northwind.xlsx]Summary!B4'
```

## What it reads, and what it does not

- **Values are the ones Excel saved.** Nothing is recalculated. A formula saved without its result, for example by a script that wrote the file, is flagged, and the outline warns about it.
- **Shown values follow the number format**: thousands separators, decimals, percentages, scaling to thousands or millions, currencies, negatives in brackets, dates and durations. Fractions are shown as decimals.
- **Formulas show as Excel shows them**, without the prefixes the file format adds (`_xlfn.XLOOKUP` is `XLOOKUP`), and are followed across sheets, through 3D references, defined names, table references and whole columns. A trace follows up to 12 formulas inside each range it meets. Cells filled by an array formula or a data table, including a spilled dynamic array, show that formula. INDIRECT, OFFSET and references to a spill (`A1#`) are flagged, because their cells are only known when Excel calculates, and so are references to deleted ranges (`#REF!`). References to other workbooks name the file but are not followed, and a formula that cannot be parsed is noted without stopping the rest.
- **Labels are reading aids.** A row label is the text in the table's leading text columns, such as `UK / Revenue`, after the date or year that leads each row of a table such as a price history or a list of sales by month (`Mar 2024 / North`). A label that repeats in its table starts with the heading of its section, so a filing's `Revenue` rows read `Taiwan / Revenue` and `China (including Hong Kong) / Revenue`. A column header comes from the nearest header row above: text, dates or consecutive years with figures under them, joined with up to 2 header rows stacked over it, such as `FY2025 / Budget`. The header of a list of labels and values stops at a blank row. A timeline at the top of a sheet heads the sections below it. Placeholders such as NM, n/a and dashes are values, not headers. Cite cells, not labels.
- **Hidden sheets, rows and columns are read and flagged**, so nothing is missed and nothing hidden is mistaken for what a reader sees.
- **Large workbooks are read in 1 pass.** A 31 MB model with 4.6 million cells and 565,000 formulas loads in 20 s and outlines in 6 s. The outline shows fewer blocks and labels per sheet as a workbook grows, so it stays a few thousand words, and the MCP server keeps the workbooks it read last in memory, up to 8 of them or 10 million cells. Calls that arrive while a workbook is being read wait for that read rather than starting another.
- **Untrusted files are handled with care.** Only .xlsx and .xlsm files are read. Old .xls files, password-protected files and Excel's lock files are refused with a reason. So is any XML the reader parses that declares a document type, or that is in an encoding which could disguise one, which rules out entity expansion; and so is a file that unpacks to more than 512 MB or holds more than 10 million cells.
- **Not read:** charts, pivot table definitions, data validation and conditional formats.

## Development

```
python -m unittest discover -s tests
ruff check src tests examples && ruff format --check src tests examples
```

`examples/northwind.py` builds the example workbook. openpyxl saves formulas without their results, so the script stores each result itself, as Excel does when it saves.

## Licence

MIT. See [LICENSE](LICENSE).
