# bb-plugin-sheet-preview

Read-only spreadsheet viewer for BB's file panel. Opening an `.xlsx`, `.xlsm`,
`.xlsb`, `.xls`, or `.ods` file shows a grid with sheet tabs instead of
"Preview not available". Every color comes from BB's theme tokens, so the grid
follows the active palette and light/dark mode without a reload.

## What it shows

- One tab per visible sheet; hidden and very-hidden sheets are skipped.
- Sticky column letters and row numbers.
- Values formatted the way Excel saved them (currency, percent, dates).
  Numbers are right-aligned, errors use the theme's destructive color.
- Merged cells, column widths from the file, hidden rows and columns.
- Hover a cell for its formula and any cell note.
- Right-click a column letter to freeze every column up to and including it,
  or to unfreeze. Each sheet keeps its own frozen columns while the file is
  open. A merged cell that crosses the frozen edge scrolls with the unfrozen
  columns.
- A formula with no saved result (a file written by a script and never opened
  in Excel) shows its formula text in muted italics instead of a blank cell.

Cell fills, fonts, borders, charts, and images are not drawn. Formulas are not
recalculated and macros never run.

## Limits

- Files over 25 MB show a notice instead of a preview: parsing runs on the UI
  thread.
- Each sheet renders its first 2,000 rows and 200 columns; a footer states how
  much was left out.

## Default opener

BB has no registration option for a default opener; with no per-extension
choice it uses the first matching opener in plugin-id order, and another
plugin may claim spreadsheet extensions first. A content script
(`default-opener.ts`) seeds the choice that Settings → File openers stores
(the browser-local `bb.fileOpenerByExtension` key):

- It runs once per browser and pins `xlsx`, `xlsm`, `xlsb`, `xls`, and `ods`
  to Sheet Preview only where that extension has no choice yet, then writes
  the marker `bb-plugin:sheet-preview:default-opener-seeded:v1`.
- An existing choice is never overridden. Picking **Automatic** later removes
  the pin, and the marker keeps the script from re-adding it.
- After uninstalling, the pin points to a missing opener and BB falls back to
  its built-in preview; choose Automatic under Settings → File openers to undo.

A file already open in a panel tab keeps the viewer it was opened with; close
and reopen it after installing.

## How it works

The plugin is frontend-only (the server entry is empty). `load.ts` fetches the
file's bytes from BB's own HTTP API using the same source → endpoint mapping as
BB's built-in PDF preview. `sheet.ts` parses the workbook with
[SheetJS](https://sheetjs.com) Community Edition and builds the grid model;
`app.tsx` renders it.

SheetJS 0.20.3 is vendored in `vendor/` from the official CDN
(`https://cdn.sheetjs.com/xlsx-0.20.3/xlsx-0.20.3.tgz`, sha256
`8dc73fc3b00203e72d176e85b50938627c7b086e607c682e8d3c22c02bb99fe8`). The `xlsx`
package on the npm registry stopped at 0.18.5, which has known
prototype-pollution and ReDoS advisories.

## Develop

```sh
npm install
npx tsc --noEmit -p .
bb plugin build
bb plugin install .
```

## License

MIT. SheetJS Community Edition is Apache-2.0.
