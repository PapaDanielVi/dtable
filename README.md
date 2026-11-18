# DTable

A Manifest V3 Chrome extension that extracts HTML tables from any webpage and exports them automatically as Excel (`.xlsx`) workbooks or RFC 4180 CSV (`.csv`) files.

[![License: MIT](https://img.shields.io/badge/License-MIT-4F46E5.svg)](LICENSE)
[![Manifest V3](https://img.shields.io/badge/Chrome-Manifest%20V3-4F46E5.svg)](manifest.json)

## Features

- **Export All Tables:** Scans the active page for visible, non-empty `<table>` elements across `<thead>`, `<tbody>`, and `<tfoot>`. Handles `rowspan` and `colspan` alignment, merges cells in `.xlsx` output, strips `<script>` and `<style>` tags, and normalizes cell whitespace.
- **Excel and CSV Modes:** Exports all tables as separate sheets inside a single `.xlsx` workbook via bundled SheetJS (`0.20.3`), or exports each table as an individual RFC 4180 compliant `.csv` file.
- **Smart Naming:** Resolves table and worksheet names automatically in priority order, enforces Excel's 31-character limit, strips forbidden sheet characters, and deduplicates colliding names.
- **Interactive Page Picker:** Activates an isolated Shadow DOM overlay so you can hover over any table on the page to inspect its title and dimensions, then click once to export just that table (or press `Escape` to cancel).
- **Popup Mini-List:** Displays detected tables with row/column counts, hover-to-scroll-into-view highlighting, and one-click per-table export.
- **Persistent Preferences:** Remembers your chosen export format (`xlsx` or `csv`) in `chrome.storage.local`.

## Installation

### Load Unpacked (Local Development)

1. Clone this repository:
   ```sh
   git clone https://github.com/PapaDanielVi/dtable.git
   cd dtable
   ```
2. (Optional) Build a clean unpacked directory:
   ```sh
   make build
   ```
3. Open `chrome://extensions` in Google Chrome and enable **Developer mode**.
4. Click **Load unpacked** and select either the repository root or `dist/dtable/`.

## Usage

1. Navigate to any webpage containing HTML tables and click the **DTable** icon in the Chrome toolbar.
2. Choose **Excel (.xlsx)** or **CSV (.csv)** in the format switcher.
3. Click **Export All Tables** to download the full-page export, click the download button next to an individual table in the list, or click **Select Table on Page** to pick a table directly on the page.

### Output Filenames

- Full-page Excel export: `<Page_Title>_DTable.xlsx`
- Single-table export: `<Page_Title>_<Table_Name>.xlsx` or `<Page_Title>_<Table_Name>.csv`
- Full-page CSV export: one `<Page_Title>_<Table_Name>.csv` per detected table

### Table Naming Priority

Each table name is resolved in this order:

1. Text inside the table's `<caption>` element
2. `aria-label` or `title` attribute on `<table>`
3. Text of element(s) referenced by `aria-labelledby`
4. Nearest preceding heading (`<h1>` through `<h6>`) within the same parent section or container
5. Readable words parsed from the `<table>` `id` attribute
6. Fallback name (`Table 1`, `Table 2`, ...)

For Excel worksheets, characters `\ / ? * [ ] :` are stripped, names are capped at 31 characters, and duplicates are suffixed with `(2)`, `(3)`, etc. within the 31-character limit.

## Development Commands

Common tasks are wrapped in the [Makefile](Makefile):

| Command | Description |
| --- | --- |
| `make help` | List available targets |
| `make check` | Validate `manifest.json`, JS syntax, and required assets |
| `make test` | Run syntax checks and unit tests (`scripts/verify.js`) |
| `make build` | Copy runtime extension files into `dist/dtable/` |
| `make package` | Run tests, build, and create `dist/dtable-v<version>.zip` |
| `make update-sheetjs` | Download `xlsx.full.min.js` into `lib/` |
| `make open-chrome` | Open `chrome://extensions` in Google Chrome (macOS) |
| `make clean` | Remove `dist/` build artifacts |

## Permissions

DTable requests minimal Manifest V3 permissions with no broad host permissions:

- `activeTab`: Grants temporary access to the current tab when you open the popup.
- `scripting`: Injects `content.js` and `content.css` on demand into the active tab.
- `storage`: Persists your preferred export format in `chrome.storage.local`.
- `downloads`: Saves generated `.xlsx` and `.csv` files directly to your Downloads folder.

## Project Structure

```text
dtable/
├── manifest.json         # Extension manifest (MV3)
├── background.js         # Service worker for XLSX/CSV generation and downloads
├── content.js            # DOM scanner, table extractor, and Shadow DOM picker
├── content.css           # Scroll-into-view highlight styles
├── popup.html            # Extension popup markup
├── popup.css             # Adaptive light/dark popup styles
├── popup.js              # Popup state, tab messaging, and format switcher
├── icons/                # Extension PNG icons (16, 32, 48, 128)
├── lib/
│   └── xlsx.full.min.js  # Bundled SheetJS library (CSP compliant)
├── scripts/
│   └── verify.js         # Automated unit verification script
├── Makefile              # Build, test, and packaging commands
└── LICENSE               # MIT License
```

## Contributing

Issues and pull requests are welcome. Before opening a pull request, run `make test` and `make package` to verify syntax, export tests, and bundle assembly.

## License

Released under the [MIT License](LICENSE).
