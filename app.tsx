// Read-only spreadsheet viewer registered as a BB file opener. Every color
// comes from BB's theme tokens so the grid follows the active palette.
import { useEffect, useLayoutEffect, useMemo, useRef, useState } from "react";
import type { RefObject } from "react";
import { definePluginApp } from "@get-bb/plugin-sdk/app";
import type { PluginFileOpenerProps } from "@get-bb/plugin-sdk/app";
import type { WorkBook } from "xlsx";
import {
  ContextMenu,
  ContextMenuContent,
  ContextMenuItem,
  ContextMenuTrigger,
} from "@/components/ui/context-menu";
import { cn } from "@/lib/utils";
import { FileTooLargeError, fileUrl, loadFileBytes } from "./load";
import { OPENER_ID, SPREADSHEET_EXTENSIONS, seedDefaultOpener } from "./default-opener";
import { buildGrid, parseWorkbook, visibleSheets } from "./sheet";
import type { CellKind, SheetGrid, SheetTab } from "./sheet";

type LoadState =
  | { status: "loading" }
  | { status: "error"; message: string; retryable: boolean }
  | { status: "ready"; workbook: WorkBook; tabs: SheetTab[] };

const CELL_CLASS: Record<CellKind, string> = {
  empty: "",
  text: "",
  number: "text-right tabular-nums",
  bool: "text-center",
  error: "text-destructive",
  formula: "italic text-muted-foreground",
};

/** Rows and columns pinned at the top and left edges, counted over the grid. */
interface FrozenPanes {
  rows: number;
  cols: number;
}

/** What a right-click landed on: a cell, a row number, or a column letter. */
interface MenuTarget {
  row: number | null;
  col: number | null;
}

const FROZEN_EDGE_RIGHT = "border-r-2 border-r-muted-foreground/50";
const FROZEN_EDGE_BOTTOM = "border-b-2 border-b-muted-foreground/50";

/** Spreadsheet-style name of the first `count` columns, e.g. "columns A–C". */
function columnRangeLabel(grid: SheetGrid, count: number): string {
  const first = grid.columns[0].label;
  const last = grid.columns[count - 1].label;
  return count === 1 ? `column ${first}` : `columns ${first}–${last}`;
}

/** Name of the first `count` visible rows by their sheet numbers, e.g. "rows 1–5". */
function rowRangeLabel(grid: SheetGrid, count: number): string {
  const first = grid.rows[0].number;
  const last = grid.rows[count - 1].number;
  return count === 1 ? `row ${first}` : `rows ${first}–${last}`;
}

/**
 * Sticky offsets (px) of the first `count` columns or rows: the size of `lead`
 * (the header corner) plus the sizes of the items before each one. Measured
 * from the DOM so they track auto-sized widths and heights; empty until measured.
 */
function useStickyOffsets(
  count: number,
  axis: "width" | "height",
  grid: SheetGrid,
  tableRef: RefObject<HTMLTableElement | null>,
  leadRef: RefObject<Element | null>,
  itemRefs: RefObject<(Element | null)[]>,
): number[] {
  const [offsets, setOffsets] = useState<number[]>([]);
  useLayoutEffect(() => {
    const table = tableRef.current;
    const lead = leadRef.current;
    if (count === 0 || table === null || lead === null) {
      setOffsets([]);
      return;
    }
    const measure = () => {
      const next: number[] = [];
      // Sub-pixel sizes: rounding each one would drift the later offsets.
      let offset = lead.getBoundingClientRect()[axis];
      for (let index = 0; index < count; index += 1) {
        next.push(offset);
        offset += itemRefs.current[index]?.getBoundingClientRect()[axis] ?? 0;
      }
      setOffsets((previous) =>
        previous.length === next.length && previous.every((value, index) => value === next[index])
          ? previous
          : next,
      );
    };
    measure();
    const observer = new ResizeObserver(measure);
    observer.observe(table);
    return () => observer.disconnect();
  }, [count, axis, grid, tableRef, leadRef, itemRefs]);
  return offsets;
}

function FreezeMenuItems({
  grid,
  frozen,
  target,
  onFreeze,
}: {
  grid: SheetGrid;
  frozen: FrozenPanes;
  target: MenuTarget;
  onFreeze: (panes: FrozenPanes) => void;
}) {
  // Freezing through a row or column includes it, like the column letters.
  const rows = target.row === null ? null : target.row + 1;
  const cols = target.col === null ? null : target.col + 1;
  return (
    <>
      {rows !== null && cols !== null && (frozen.rows !== rows || frozen.cols !== cols) ? (
        <ContextMenuItem onSelect={() => onFreeze({ rows, cols })}>
          Freeze {rowRangeLabel(grid, rows)} and {columnRangeLabel(grid, cols)}
        </ContextMenuItem>
      ) : null}
      {rows !== null && frozen.rows !== rows ? (
        <ContextMenuItem onSelect={() => onFreeze({ rows, cols: frozen.cols })}>
          Freeze {rowRangeLabel(grid, rows)}
        </ContextMenuItem>
      ) : null}
      {cols !== null && frozen.cols !== cols ? (
        <ContextMenuItem onSelect={() => onFreeze({ rows: frozen.rows, cols })}>
          Freeze {columnRangeLabel(grid, cols)}
        </ContextMenuItem>
      ) : null}
      {rows !== null && cols !== null ? (
        <ContextMenuItem
          disabled={frozen.rows === 0 && frozen.cols === 0}
          onSelect={() => onFreeze({ rows: 0, cols: 0 })}
        >
          Unfreeze rows and columns
        </ContextMenuItem>
      ) : rows !== null ? (
        <ContextMenuItem
          disabled={frozen.rows === 0}
          onSelect={() => onFreeze({ rows: 0, cols: frozen.cols })}
        >
          Unfreeze rows
        </ContextMenuItem>
      ) : (
        <ContextMenuItem
          disabled={frozen.cols === 0}
          onSelect={() => onFreeze({ rows: frozen.rows, cols: 0 })}
        >
          Unfreeze columns
        </ContextMenuItem>
      )}
    </>
  );
}

function SheetTable({
  grid,
  frozen,
  onFreeze,
}: {
  grid: SheetGrid;
  frozen: FrozenPanes;
  onFreeze: (panes: FrozenPanes) => void;
}) {
  const tableRef = useRef<HTMLTableElement>(null);
  const cornerRef = useRef<HTMLTableCellElement>(null);
  const headerRefs = useRef<(HTMLTableCellElement | null)[]>([]);
  const rowRefs = useRef<(HTMLTableRowElement | null)[]>([]);
  const leftOffsets = useStickyOffsets(frozen.cols, "width", grid, tableRef, cornerRef, headerRefs);
  const topOffsets = useStickyOffsets(frozen.rows, "height", grid, tableRef, cornerRef, rowRefs);
  const [menuTarget, setMenuTarget] = useState<MenuTarget>({ row: null, col: 0 });

  // A cell sticks only when its whole span sits inside the frozen range, so a
  // merge that crosses the edge scrolls instead of covering other cells.
  const stickyLeft = (col: number, colSpan: number): number | undefined =>
    col + colSpan <= frozen.cols ? leftOffsets[col] : undefined;
  const stickyTop = (row: number, rowSpan: number): number | undefined =>
    row + rowSpan <= frozen.rows ? topOffsets[row] : undefined;

  // Layering, low to high: cells frozen on one axis pass beneath cells frozen
  // on both, and every body cell passes beneath the header row.
  return (
    <ContextMenu>
      <ContextMenuTrigger asChild>
        <table
          ref={tableRef}
          className="border-separate border-spacing-0 font-mono text-xs text-foreground"
        >
          <thead>
            <tr>
              <th
                ref={cornerRef}
                aria-hidden="true"
                // The corner has nothing to freeze: keep the browser's own menu.
                onContextMenu={(event) => event.stopPropagation()}
                className="sticky left-0 top-0 z-50 border-b border-r border-border bg-card"
              />
              {grid.columns.map((column, col) => {
                const left = stickyLeft(col, 1);
                return (
                  <th
                    key={column.label}
                    ref={(element) => {
                      headerRefs.current[col] = element;
                    }}
                    scope="col"
                    onContextMenu={() => setMenuTarget({ row: null, col })}
                    style={left === undefined ? undefined : { left }}
                    className={cn(
                      "sticky top-0 border-b border-r border-border bg-card px-2 py-1 text-center font-normal text-muted-foreground",
                      left === undefined ? "z-30" : "z-40",
                      col + 1 === frozen.cols && FROZEN_EDGE_RIGHT,
                    )}
                  >
                    {column.label}
                  </th>
                );
              })}
            </tr>
          </thead>
          <tbody>
            {grid.rows.map((gridRow, row) => {
              const rowTop = stickyTop(row, 1);
              return (
                <tr
                  key={gridRow.number}
                  ref={(element) => {
                    rowRefs.current[row] = element;
                  }}
                  className="group"
                >
                  <th
                    scope="row"
                    onContextMenu={() => setMenuTarget({ row, col: null })}
                    style={rowTop === undefined ? undefined : { top: rowTop }}
                    className={cn(
                      "sticky left-0 border-b border-r border-border bg-card px-2 text-right font-normal tabular-nums text-muted-foreground",
                      rowTop === undefined ? "z-10" : "z-30",
                      row + 1 === frozen.rows && FROZEN_EDGE_BOTTOM,
                    )}
                  >
                    {gridRow.number}
                  </th>
                  {gridRow.cells.map((cell, col) => {
                    if (cell === null) return null;
                    const left = stickyLeft(col, cell.colSpan);
                    const top = stickyTop(row, cell.rowSpan);
                    const sticky = left !== undefined || top !== undefined;
                    return (
                      <td
                        key={grid.columns[col].label}
                        rowSpan={cell.rowSpan}
                        colSpan={cell.colSpan}
                        title={cell.title}
                        onContextMenu={() => setMenuTarget({ row, col })}
                        style={sticky ? { left, top } : undefined}
                        className={cn(
                          "border-b border-r border-border px-2 py-1 align-top",
                          // Frozen cells stay opaque so scrolled cells pass beneath
                          // them; the hover tint is layered on top as an image.
                          sticky
                            ? "sticky bg-background group-hover:bg-linear-to-r group-hover:from-state-hover group-hover:to-state-hover"
                            : "group-hover:bg-state-hover",
                          left !== undefined && top !== undefined
                            ? "z-30"
                            : top !== undefined
                              ? "z-20"
                              : left !== undefined && "z-10",
                          col + cell.colSpan === frozen.cols && FROZEN_EDGE_RIGHT,
                          row + cell.rowSpan === frozen.rows && FROZEN_EDGE_BOTTOM,
                          CELL_CLASS[cell.kind],
                        )}
                      >
                        <div
                          className="max-w-[28rem] truncate"
                          style={{ minWidth: cell.colSpan === 1 ? grid.columns[col].widthPx : undefined }}
                        >
                          {cell.text}
                        </div>
                      </td>
                    );
                  })}
                </tr>
              );
            })}
          </tbody>
        </table>
      </ContextMenuTrigger>
      <ContextMenuContent>
        <FreezeMenuItems grid={grid} frozen={frozen} target={menuTarget} onFreeze={onFreeze} />
      </ContextMenuContent>
    </ContextMenu>
  );
}

function WorkbookView({ workbook, tabs }: { workbook: WorkBook; tabs: SheetTab[] }) {
  const [selectedName, setSelectedName] = useState(tabs[0].name);
  // Frozen rows and columns per sheet, kept while the file stays open.
  const [frozenBySheet, setFrozenBySheet] = useState<Record<string, FrozenPanes>>({});
  // A reload or another file can drop the selected sheet; fall back to the first.
  const activeName = tabs.some((tab) => tab.name === selectedName) ? selectedName : tabs[0].name;
  const grid = useMemo(() => buildGrid(workbook.Sheets[activeName]), [workbook, activeName]);
  const frozen: FrozenPanes = {
    rows: Math.min(frozenBySheet[activeName]?.rows ?? 0, grid.rows.length),
    cols: Math.min(frozenBySheet[activeName]?.cols ?? 0, grid.columns.length),
  };
  const notes: string[] = [];
  if (grid.totalRows > grid.shownRows) {
    notes.push(`first ${grid.shownRows.toLocaleString()} of ${grid.totalRows.toLocaleString()} rows`);
  }
  if (grid.totalCols > grid.shownCols) {
    notes.push(`first ${grid.shownCols.toLocaleString()} of ${grid.totalCols.toLocaleString()} columns`);
  }

  return (
    <div className="flex h-full min-h-0 flex-col bg-background text-foreground">
      <div className="min-h-0 flex-1 overflow-auto">
        {grid.rows.length === 0 ? (
          <div className="p-4">
            <div
              role="status"
              className="rounded-lg border border-dashed border-border px-4 py-6 text-center text-sm text-muted-foreground"
            >
              This sheet is empty.
            </div>
          </div>
        ) : (
          <SheetTable
            grid={grid}
            frozen={frozen}
            onFreeze={(panes) => setFrozenBySheet((previous) => ({ ...previous, [activeName]: panes }))}
          />
        )}
      </div>
      {notes.length === 0 ? null : (
        <p className="shrink-0 border-t border-border px-3 py-1 text-xs text-muted-foreground">
          Showing {notes.join(", ")}.
        </p>
      )}
      <div
        role="tablist"
        aria-label="Sheets"
        className="flex shrink-0 overflow-x-auto border-t border-border bg-card px-2"
      >
        {tabs.map((tab) => {
          const active = tab.name === activeName;
          return (
            <button
              key={tab.name}
              type="button"
              role="tab"
              aria-selected={active}
              onClick={() => setSelectedName(tab.name)}
              className={cn(
                "-mt-px whitespace-nowrap border-x border-t px-3 py-1.5 text-xs focus-visible:outline-none focus-visible:ring-1 focus-visible:ring-ring",
                active
                  ? "border-border bg-background font-medium text-foreground"
                  : "border-transparent text-muted-foreground hover:bg-state-hover hover:text-foreground",
              )}
            >
              {tab.name}
            </button>
          );
        })}
      </div>
    </div>
  );
}

function SpreadsheetViewer({ path, source, Original }: PluginFileOpenerProps) {
  const [attempt, setAttempt] = useState(0);
  const [state, setState] = useState<LoadState>({ status: "loading" });
  const url = useMemo(
    () => fileUrl(path, source),
    // `source` is a fresh object per render; its fields are the identity.
    [path, source.kind, source.threadId, source.environmentId, source.projectId, source.experimental_hostId],
  );

  useEffect(() => {
    if (url === null) return;
    const controller = new AbortController();
    setState({ status: "loading" });
    loadFileBytes(url, controller.signal)
      .then((bytes) => {
        if (controller.signal.aborted) return;
        const workbook = parseWorkbook(bytes);
        setState({ status: "ready", workbook, tabs: visibleSheets(workbook) });
      })
      .catch((cause: unknown) => {
        if (controller.signal.aborted) return;
        setState({
          status: "error",
          message: cause instanceof Error ? cause.message : String(cause),
          retryable: !(cause instanceof FileTooLargeError),
        });
      });
    return () => controller.abort();
  }, [url, attempt]);

  if (url === null) return <Original />;

  if (state.status === "loading") {
    return (
      <div
        role="status"
        aria-label={`Loading ${path}`}
        className="flex h-full min-h-0 items-center justify-center gap-2 bg-background text-sm text-muted-foreground"
      >
        <span className="size-4 animate-spin rounded-full border-2 border-border border-t-foreground" />
        Loading spreadsheet…
      </div>
    );
  }

  if (state.status === "error") {
    return (
      <div className="flex h-full min-h-0 items-center justify-center bg-background p-6">
        <div className="max-w-md space-y-3 text-center" role="alert">
          <p className={cn("text-sm", state.retryable ? "text-destructive" : "text-muted-foreground")}>
            {state.retryable ? `Failed to open spreadsheet: ${state.message}` : state.message}
          </p>
          {state.retryable ? (
            <button
              type="button"
              className="inline-flex h-8 items-center justify-center rounded-md border border-border bg-background px-3 text-sm font-medium text-foreground hover:bg-state-hover focus-visible:outline-none focus-visible:ring-1 focus-visible:ring-ring"
              onClick={() => setAttempt((value) => value + 1)}
            >
              Retry
            </button>
          ) : null}
        </div>
      </div>
    );
  }

  return <WorkbookView workbook={state.workbook} tabs={state.tabs} />;
}

export default definePluginApp((app) => {
  app.slots.fileOpener({
    id: OPENER_ID,
    title: "Excel Preview",
    extensions: SPREADSHEET_EXTENSIONS,
    component: SpreadsheetViewer,
  });
  // One-shot and synchronous: nothing stays mounted, so there is nothing to dispose.
  app.contentScripts.register({
    id: "default-spreadsheet-opener",
    mount: ({ pluginId }) => seedDefaultOpener(pluginId),
  });
});
