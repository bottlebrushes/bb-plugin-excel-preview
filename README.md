# bb-plugin-excel-preview

Spreadsheet viewer and light editor for BB's file panel. Opening an `.xlsx`,
`.xlsm`, `.xlsb`, `.xls`, or `.ods` file shows a grid with sheet tabs instead of
"Preview not available". Every color comes from BB's theme tokens, so the grid
follows the active palette and light/dark mode without a reload. With
LibreOffice on the bb server machine, formulas show their results and cells can
be edited.

## What it shows

- One tab per visible sheet; hidden and very-hidden sheets are skipped.
- Sticky column letters and row numbers.
- Values formatted the way Excel saved them (currency, percent, dates).
  Numbers are right-aligned, errors use the theme's destructive color.
- Merged cells, column widths from the file, hidden rows and columns.
- Hover a cell for its formula and any cell note.
- Freeze panes from the right-click menu (a long-press on touch screens). On a
  cell: freeze the rows and columns up to and including it, or just its rows
  or columns. On a row number or column letter: freeze the rows or columns up
  to it. Every menu can also freeze only that one row or column: it scrolls
  until it reaches the edge, then stays there while the ones before it scroll
  away. Each menu also unfreezes. Each sheet keeps its own frozen panes while
  the file is open. A merged cell that crosses a frozen edge scrolls with the
  unfrozen cells.
- A formula with no saved result (a file written by a script and never opened
  in Excel) is recalculated by LibreOffice and shows its result; a footer says
  so. The file itself is not changed. Without LibreOffice the cell shows its
  formula text in muted italics.

Cell fills, fonts, borders, charts, and images are not drawn. Macros never run.

## Editing

Click a cell to select it; arrow keys and Tab move the selection. Type to
replace the cell, or press Enter, F2, or double-click to edit what it holds.
Enter or Tab saves, Escape cancels, and Delete or Backspace clears the cell.
Input is read as if typed into Excel: `=` starts a formula (Excel syntax, e.g.
`='Other sheet'!B2*2`), `12.5%` is a percentage, other text is a number or
text.

Each edit is saved straight to the file. LibreOffice opens a copy, applies the
edit, recalculates every formula, and saves it in the file's own format; the
result replaces the file only if nobody changed it since it was opened
(sha256 compare-and-swap). Otherwise the footer offers a reload and nothing is
overwritten.

LibreOffice rewrites the whole file on save. Values, formulas, number formats,
merges, column widths, and sheets survive; Excel-only features LibreOffice
does not support may be lost. `.xlsb` files stay read-only because
LibreOffice cannot write them.

## Requirements

- BB 0.49 or later.
- For formula results and editing: LibreOffice Calc and its Python bridge on
  the machine running the bb server (Debian/Ubuntu:
  `apt install libreoffice-calc-nogui python3-uno`; macOS: the LibreOffice app,
  which bundles both). Without them the viewer is read-only, and the footer
  says why. The file may live on any connected host.

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
  to Excel Preview only where that extension has no choice yet, then writes
  the marker `bb-plugin:excel-preview:default-opener-seeded:v1`.
- An existing choice is never overridden. Picking **Automatic** later removes
  the pin, and the marker keeps the script from re-adding it.
- After uninstalling, the pin points to a missing opener and BB falls back to
  its built-in preview; choose Automatic under Settings → File openers to undo.
- A pin left by the plugin's earlier id, `sheet-preview`, counts as no
  choice and is re-pointed to Excel Preview.

A file already open in a panel tab keeps the viewer it was opened with; close
and reopen it after installing.

## How it works

`server.ts` finds the opened file on its host from the opener source (the
environment's directory, the project checkout, the thread's machine, or thread
storage) and reads and writes it through `bb.sdk.files`. `sheet.ts` parses the
workbook with [SheetJS](https://sheetjs.com) Community Edition and builds the
grid model; `app.tsx` renders it and handles selection and editing.

`libreoffice.ts` runs one headless LibreOffice per bb server, started on first
use with a private profile and stopped after five idle minutes. It is driven
over the UNO bridge by `libreoffice_worker.py`, which the build inlines into
the server bundle. Requests run one at a time.

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
