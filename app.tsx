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
import { FileTooLargeError, fileRequest, loadFileBytes } from "./load";
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

/** Spreadsheet-style name of the frozen range, e.g. "column A" or "columns A–C". */
function frozenRangeLabel(grid: SheetGrid, count: number): string {
  const first = grid.columns[0].label;
  const last = grid.columns[count - 1].label;
  return count === 1 ? `column ${first}` : `columns ${first}–${last}`;
}

/**
 * Left offsets (px) of the first `frozen` columns, measured from the header row
 * so they track the auto-sized column widths. Empty until measured.
 */
function useFrozenOffsets(
  frozen: number,
  grid: SheetGrid,
  tableRef: RefObject<HTMLTableElement | null>,
  cornerRef: RefObject<HTMLTableCellElement | null>,
  headerRefs: RefObject<(HTMLTableCellElement | null)[]>,
): number[] {
  const [offsets, setOffsets] = useState<number[]>([]);
  useLayoutEffect(() => {
    const table = tableRef.current;
    const corner = cornerRef.current;
    if (frozen === 0 || table === null || corner === null) {
      setOffsets([]);
      return;
    }
    const measure = () => {
      const next: number[] = [];
      // Sub-pixel widths: rounding each one would drift the later offsets.
      let left = corner.getBoundingClientRect().width;
      for (let index = 0; index < frozen; index += 1) {
        next.push(left);
        left += headerRefs.current[index]?.getBoundingClientRect().width ?? 0;
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
  }, [frozen, grid, tableRef, cornerRef, headerRefs]);
  return offsets;
}

function SheetTable({
  grid,
  frozen,
  onFreeze,
}: {
  grid: SheetGrid;
  /** Columns frozen at the left edge, counted over `grid.columns`. */
  frozen: number;
  onFreeze: (count: number) => void;
}) {
  const tableRef = useRef<HTMLTableElement>(null);
  const cornerRef = useRef<HTMLTableCellElement>(null);
  const headerRefs = useRef<(HTMLTableCellElement | null)[]>([]);
  const offsets = useFrozenOffsets(frozen, grid, tableRef, cornerRef, headerRefs);
  // The column whose header was right-clicked; the menu acts on it.
  const [menuColumn, setMenuColumn] = useState(0);

  // A cell sticks only when its whole span sits inside the frozen range, so a
  // merge that crosses the boundary scrolls instead of covering other columns.
  const stickyLeft = (index: number, colSpan: number): number | undefined =>
    index + colSpan <= frozen ? offsets[index] : undefined;

  return (
    <table
      ref={tableRef}
      className="border-separate border-spacing-0 font-mono text-xs text-foreground"
    >
      <ContextMenu>
        <ContextMenuTrigger asChild>
          <thead>
            <tr>
              <th
                ref={cornerRef}
                aria-hidden="true"
                // The corner has no column to act on: keep the browser's own menu.
                onContextMenu={(event) => event.stopPropagation()}
                className="sticky left-0 top-0 z-40 border-b border-r border-border bg-card"
              />
              {grid.columns.map((column, index) => {
                const left = stickyLeft(index, 1);
                return (
                  <th
                    key={column.label}
                    ref={(element) => {
                      headerRefs.current[index] = element;
                    }}
                    scope="col"
                    onContextMenu={() => setMenuColumn(index)}
                    style={left === undefined ? undefined : { left }}
                    className={cn(
                      "sticky top-0 border-b border-r border-border bg-card px-2 py-1 text-center font-normal text-muted-foreground",
                      left === undefined ? "z-20" : "z-30",
                      index + 1 === frozen && "border-r-2 border-r-muted-foreground/50",
                    )}
                  >
                    {column.label}
                  </th>
                );
              })}
            </tr>
          </thead>
        </ContextMenuTrigger>
        <ContextMenuContent>
          {frozen === menuColumn + 1 ? null : (
            <ContextMenuItem onSelect={() => onFreeze(menuColumn + 1)}>
              Freeze {frozenRangeLabel(grid, menuColumn + 1)}
            </ContextMenuItem>
          )}
          <ContextMenuItem disabled={frozen === 0} onSelect={() => onFreeze(0)}>
            Unfreeze columns
          </ContextMenuItem>
        </ContextMenuContent>
      </ContextMenu>
      <tbody>
        {grid.rows.map((row) => (
          <tr key={row.number} className="group">
            <th
              scope="row"
              className="sticky left-0 z-10 border-b border-r border-border bg-card px-2 text-right font-normal tabular-nums text-muted-foreground"
            >
              {row.number}
            </th>
            {row.cells.map((cell, index) => {
              if (cell === null) return null;
              const left = stickyLeft(index, cell.colSpan);
              return (
                <td
                  key={grid.columns[index].label}
                  rowSpan={cell.rowSpan}
                  colSpan={cell.colSpan}
                  title={cell.title}
                  style={left === undefined ? undefined : { left }}
                  className={cn(
                    "border-b border-r border-border px-2 py-1 align-top",
                    // Frozen cells stay opaque so scrolled cells pass beneath
                    // them; the hover tint is layered on top as an image.
                    left === undefined
                      ? "group-hover:bg-state-hover"
                      : "sticky z-10 bg-background group-hover:bg-linear-to-r group-hover:from-state-hover group-hover:to-state-hover",
                    index + cell.colSpan === frozen && "border-r-2 border-r-muted-foreground/50",
                    CELL_CLASS[cell.kind],
                  )}
                >
                  <div
                    className="max-w-[28rem] truncate"
                    style={{ minWidth: cell.colSpan === 1 ? grid.columns[index].widthPx : undefined }}
                  >
                    {cell.text}
                  </div>
                </td>
              );
            })}
          </tr>
        ))}
      </tbody>
    </table>
  );
}

function WorkbookView({ workbook, tabs }: { workbook: WorkBook; tabs: SheetTab[] }) {
  const [selectedName, setSelectedName] = useState(tabs[0].name);
  // Frozen column count per sheet, kept while the file stays open.
  const [frozenBySheet, setFrozenBySheet] = useState<Record<string, number>>({});
  // A reload or another file can drop the selected sheet; fall back to the first.
  const activeName = tabs.some((tab) => tab.name === selectedName) ? selectedName : tabs[0].name;
  const grid = useMemo(() => buildGrid(workbook.Sheets[activeName]), [workbook, activeName]);
  const frozen = Math.min(frozenBySheet[activeName] ?? 0, grid.columns.length);
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
            onFreeze={(count) => setFrozenBySheet((previous) => ({ ...previous, [activeName]: count }))}
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
  const request = useMemo(
    () => fileRequest(path, source),
    // `source` is a fresh object per render; its fields are the identity.
    [path, source.kind, source.threadId, source.environmentId, source.projectId, source.experimental_hostId],
  );

  useEffect(() => {
    if (request === null) return;
    const controller = new AbortController();
    setState({ status: "loading" });
    loadFileBytes(request, controller.signal)
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
  }, [request, attempt]);

  if (request === null) return <Original />;

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
    title: "Spreadsheet viewer",
    extensions: SPREADSHEET_EXTENSIONS,
    component: SpreadsheetViewer,
  });
  // One-shot and synchronous: nothing stays mounted, so there is nothing to dispose.
  app.contentScripts.register({
    id: "default-spreadsheet-opener",
    mount: ({ pluginId }) => seedDefaultOpener(pluginId),
  });
});
