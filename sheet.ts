// Workbook parsing and the grid model the viewer renders. Pure (no DOM, no
// React) so it runs the same in the browser bundle and in Node smoke checks.
import { read, utils } from "xlsx";
import type { CellObject, WorkBook, WorkSheet } from "xlsx";

/** Rows parsed per sheet; the rest stay in the file and are counted from `!fullref`. */
export const MAX_ROWS = 2000;
/** Columns rendered per sheet. */
export const MAX_COLS = 200;

const DEFAULT_COLUMN_PX = 88;
const MIN_COLUMN_PX = 32;
const MAX_COLUMN_PX = 448;

export type CellKind = "empty" | "text" | "number" | "bool" | "error" | "formula";

export interface GridCell {
  /** Display text: Excel's formatted value, or `=FORMULA` when no result was saved. */
  text: string;
  kind: CellKind;
  /** Hover text: the formula and any cell note. */
  title: string | undefined;
  rowSpan: number;
  colSpan: number;
}

export interface GridColumn {
  label: string;
  widthPx: number;
}

export interface GridRow {
  /** One-based row number as Excel shows it. */
  number: number;
  /** `null` marks a slot covered by a merged cell to its upper-left. */
  cells: (GridCell | null)[];
}

export interface SheetGrid {
  columns: GridColumn[];
  rows: GridRow[];
  /** Rows and columns in the sheet's full used range. */
  totalRows: number;
  totalCols: number;
  /** Rows and columns of that range inside the rendered window (hidden ones included). */
  shownRows: number;
  shownCols: number;
}

export interface SheetTab {
  name: string;
}

export function parseWorkbook(bytes: ArrayBuffer): WorkBook {
  return read(new Uint8Array(bytes), {
    type: "array",
    dense: true,
    cellFormula: true,
    // Needed for column widths and hidden rows/columns; colors are ignored.
    cellStyles: true,
    cellHTML: false,
    sheetRows: MAX_ROWS,
  });
}

/** Sheets a reader would see in Excel: hidden and very-hidden ones are skipped. */
export function visibleSheets(workbook: WorkBook): SheetTab[] {
  const meta = workbook.Workbook?.Sheets;
  const visible = workbook.SheetNames.filter(
    (_name, index) => (meta?.[index]?.Hidden ?? 0) === 0,
  );
  return (visible.length > 0 ? visible : workbook.SheetNames).map((name) => ({ name }));
}

function columnWidthPx(sheet: WorkSheet, column: number): number {
  const info = sheet["!cols"]?.[column];
  const px =
    info?.wpx ??
    (info?.wch !== undefined ? info.wch * 7 + 5 : undefined) ??
    (info?.width !== undefined ? info.width * 6 : undefined) ??
    DEFAULT_COLUMN_PX;
  return Math.min(MAX_COLUMN_PX, Math.max(MIN_COLUMN_PX, Math.round(px)));
}

function cellTitle(cell: CellObject): string | undefined {
  const parts: string[] = [];
  if (cell.f !== undefined) parts.push(`=${cell.f}`);
  for (const comment of cell.c ?? []) {
    if (comment.t) parts.push(comment.t);
  }
  return parts.length > 0 ? parts.join("\n") : undefined;
}

function toGridCell(cell: CellObject | undefined): GridCell {
  if (cell === undefined) {
    return { text: "", kind: "empty", title: undefined, rowSpan: 1, colSpan: 1 };
  }
  const title = cellTitle(cell);
  // A formula written by a script and never recalculated has no saved result.
  if (cell.f !== undefined && (cell.v === undefined || cell.t === "z")) {
    return { text: `=${cell.f}`, kind: "formula", title, rowSpan: 1, colSpan: 1 };
  }
  if (cell.t === "z") {
    return { text: "", kind: "empty", title, rowSpan: 1, colSpan: 1 };
  }
  const kind: CellKind =
    cell.t === "n" || cell.t === "d"
      ? "number"
      : cell.t === "b"
        ? "bool"
        : cell.t === "e"
          ? "error"
          : "text";
  return { text: utils.format_cell(cell), kind, title, rowSpan: 1, colSpan: 1 };
}

export function buildGrid(sheet: WorkSheet): SheetGrid {
  const ref = sheet["!ref"];
  if (ref === undefined) {
    return { columns: [], rows: [], totalRows: 0, totalCols: 0, shownRows: 0, shownCols: 0 };
  }
  const range = utils.decode_range(ref);
  const full = utils.decode_range(sheet["!fullref"] ?? ref);
  const firstRow = range.s.r;
  const lastRow = range.e.r;
  const firstCol = range.s.c;
  const lastCol = Math.min(range.e.c, firstCol + MAX_COLS - 1);

  const visibleCols: number[] = [];
  for (let c = firstCol; c <= lastCol; c += 1) {
    if (sheet["!cols"]?.[c]?.hidden !== true) visibleCols.push(c);
  }
  const visibleRows: number[] = [];
  for (let r = firstRow; r <= lastRow; r += 1) {
    if (sheet["!rows"]?.[r]?.hidden !== true) visibleRows.push(r);
  }
  // Merge anchors get spans counted over visible rows/columns; the slots they
  // cover render nothing.
  const spans = new Map<string, { rowSpan: number; colSpan: number }>();
  const covered = new Set<string>();
  for (const merge of sheet["!merges"] ?? []) {
    const rows = visibleRows.filter((r) => r >= merge.s.r && r <= merge.e.r);
    const cols = visibleCols.filter((c) => c >= merge.s.c && c <= merge.e.c);
    if (rows.length === 0 || cols.length === 0) continue;
    const anchorRow = rows[0];
    const anchorCol = cols[0];
    spans.set(`${anchorRow}:${anchorCol}`, { rowSpan: rows.length, colSpan: cols.length });
    for (const r of rows) {
      for (const c of cols) {
        if (r !== anchorRow || c !== anchorCol) covered.add(`${r}:${c}`);
      }
    }
  }

  const data = sheet["!data"] ?? [];
  const rows: GridRow[] = visibleRows.map((r) => ({
    number: r + 1,
    cells: visibleCols.map((c) => {
      const key = `${r}:${c}`;
      if (covered.has(key)) return null;
      const cell = toGridCell(data[r]?.[c]);
      const span = spans.get(key);
      if (span !== undefined) {
        cell.rowSpan = span.rowSpan;
        cell.colSpan = span.colSpan;
      }
      return cell;
    }),
  }));

  return {
    columns: visibleCols.map((c) => ({
      label: utils.encode_col(c),
      widthPx: columnWidthPx(sheet, c),
    })),
    rows,
    totalRows: full.e.r - full.s.r + 1,
    totalCols: full.e.c - full.s.c + 1,
    shownRows: lastRow - firstRow + 1,
    shownCols: lastCol - firstCol + 1,
  };
}
