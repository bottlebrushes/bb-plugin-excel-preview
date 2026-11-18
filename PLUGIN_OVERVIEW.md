## What you get

- One tab per visible sheet, with column letters and row numbers that stay in view while you scroll.
- Values appear the way Excel saved them: currency, percentages, and dates keep their number formats, and numbers are right-aligned.
- Merged cells, column widths from the file, and hidden rows, columns, and sheets are respected.
- Hover a cell to see its formula and any cell note. A formula with no saved result, common in files written by scripts and agents, is recalculated by LibreOffice so you see its result.
- Edit cells in place: click a cell and type, or press Enter, F2, or double-click to change what it holds, including formulas. Each edit is saved to the file and every formula is recalculated. If the file changed on disk since you opened it, nothing is overwritten and you are offered a reload.
- Right-click (or long-press) a cell to freeze the rows above and columns left of it, including its own, so headers and labels stay in view while you scroll. Right-click a row number or column letter to freeze just rows or just columns. You can also freeze a single row or column, such as column C alone: it stays at the edge once you scroll past it. Each sheet keeps its own frozen panes while the file is open.
- The grid uses your BB theme's colors and switches with light and dark mode. The file's own cell fills are not drawn.

## Formats and limits

Opens `.xlsx`, `.xlsm`, `.xlsb`, `.xls`, and `.ods`. Each sheet shows its first 2,000 rows and 200 columns, and a footer says how much was left out. Files over 25 MB show a notice instead of a preview. Macros never run, and charts and images are not shown. LibreOffice rewrites the whole file when it saves an edit, so Excel-only features it does not support may be lost; `.xlsb` files cannot be edited.

## Default opener

On install, Excel Preview becomes the opener for those five extensions in each browser, the first time BB loads there, for any of them you have not already assigned. It never replaces a choice you made. Change it under Settings → File openers, or right-click a file link and pick Open with.

## Requirements

BB 0.49 or later. Parsing uses SheetJS Community Edition, bundled with the plugin. Formula results and editing need LibreOffice Calc and its Python bridge on the machine running the bb server (Debian/Ubuntu: `libreoffice-calc-nogui` and `python3-uno`; macOS: the LibreOffice app). Without them the viewer is read-only and says why.
