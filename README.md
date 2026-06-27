# ECU BOM Delta

A PyQt5 desktop app for comparing two Bill of Materials (BOM) Excel files side-by-side. Highlights added, removed, and changed components and exports the delta to a formatted Excel report.

Designed for hardware/PCB engineers who need to audit component changes between ECU BOM revisions without manually diffing spreadsheets.

## Features

- Select two BOM `.xlsx` files via file picker
- Four configurable filters:
  - Remove zero-quantity items in both BOMs
  - Remove new zero-quantity items (BOM2 only)
  - Remove deprecated zero-quantity items (BOM1 only)
  - Remove unchanged items
- Exports a two-sheet `.xlsx` report:
  - **Side by Side Comparison** — both BOMs aligned by RefDes, with column headers and auto-filter
  - **BOM Delta** — outer-joined diff with `_BOM1` / `_BOM2` suffixes
- Exports a `.txt` file of all differing RefDes identifiers
- Cross-platform: opens the result file automatically on Windows, macOS, and Linux

## Requirements

- Python 3.8+

## Installation

```sh
pip install -r requirements.txt
```

## Usage

```sh
python bom_delta.py
```

1. Click **Select First BOM File** and **Select Second BOM File**.
2. Toggle filters as needed.
3. Click **Run Comparison**.
4. The output opens automatically in `Result/`.

> **Note:** The tool expects BOM files with 15 header rows and 7 footer rows (standard Bosch BOM export format). Adjust `skiprows` / `skipfooter` in `_read_excel()` if your format differs.

## Tech Stack

`PyQt5` · `pandas` · `openpyxl`

## License

MIT
