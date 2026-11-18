// Spreadsheet viewer and light editor registered as a BB file opener. Every
// color comes from BB's theme tokens so the grid follows the active palette.
// Formula results and cell edits come from LibreOffice on the bb server.
import { useEffect, useLayoutEffect, useMemo, useRef, useState } from "react";
import type { KeyboardEvent, RefObject } from "react";
import { definePluginApp, useRpc } from "@get-bb/plugin-sdk/app";
import type { PluginFileOpenerProps } from "@get-bb/plugin-sdk/app";
import type { WorkBook } from "xlsx";
import {
  ContextMenu,
  ContextMenuContent,
  ContextMenuItem,
  ContextMenuTrigger,
} from "@/components/ui/context-menu";
import { cn } from "@/lib/utils";
import { OPENER_ID, SPREADSHEET_EXTENSIONS, seedDefaultOpener } from "./default-opener";
import { buildGrid, needsRecalculation, parseWorkbook, visibleSheets } from "./sheet";
import type { CellKind, GridCell, SheetGrid, SheetTab } from "./sheet";
import type { rpcContract } from "./server";

type LoadState =
  | { status: "loading" }
  | { status: "error"; message: string; retryable: boolean }
  | ({ status: "ready" } & Opened);

/** One line under the grid: progress, a result, or a failure (with a reload offer). */
interface ViewerStatus {
  tone: "muted" | "error";
  text: string;
  reload?: boolean;
}

/** Present when cells can be edited. */
interface Editing {
  /** Address of the cell being saved; no other edit starts until it lands. */
  saving: string | null;
  onCommit: (sheet: string, cell: string, input: string) => void;
}

const BUTTON_CLASS =
  "inline-flex h-8 items-center justify-center rounded-md border border-border bg-background px-3 text-sm font-medium text-foreground hover:bg-state-hover focus-visible:outline-none focus-visible:ring-1 focus-visible:ring-ring";

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

/** A slot in the rendered grid: indices into `grid.rows` and a row's `cells`. */
interface CellPosition {
  row: number;
  col: number;
}

function SheetTable({
  grid,
  frozen,
  onFreeze,
  sheet,
  editing,
}: {
  grid: SheetGrid;
  frozen: FrozenPanes;
  onFreeze: (panes: FrozenPanes) => void;
  sheet: string;
  editing: Editing | null;
}) {
  const tableRef = useRef<HTMLTableElement>(null);
  const cornerRef = useRef<HTMLTableCellElement>(null);
  const headerRefs = useRef<(HTMLTableCellElement | null)[]>([]);
  const rowRefs = useRef<(HTMLTableRowElement | null)[]>([]);
  const leftOffsets = useStickyOffsets(frozen.cols, "width", grid, tableRef, cornerRef, headerRefs);
  const topOffsets = useStickyOffsets(frozen.rows, "height", grid, tableRef, cornerRef, rowRefs);
  const [menuTarget, setMenuTarget] = useState<MenuTarget>({ row: null, col: 0 });
  const [selected, setSelected] = useState<CellPosition | null>(null);
  const [draft, setDraft] = useState<{ position: CellPosition; value: string } | null>(null);
  // Enter and blur both finish an edit; the ref lets only the first one commit.
  const draftRef = useRef(draft);
  draftRef.current = draft;

  const cellAt = (position: CellPosition): GridCell | null =>
    grid.rows[position.row]?.cells[position.col] ?? null;

  /** The next cell in a direction, skipping slots a merge covers; stays put at the edge. */
  function step(from: CellPosition, rows: number, cols: number): CellPosition {
    let next = { row: from.row + rows, col: from.col + cols };
    while (next.row >= 0 && next.row < grid.rows.length && next.col >= 0 && next.col < grid.columns.length) {
      if (cellAt(next) !== null) return next;
      next = { row: next.row + rows, col: next.col + cols };
    }
    return from;
  }

  useEffect(() => {
    if (selected === null) return;
    tableRef.current
      ?.querySelector(`[data-slot="${selected.row}:${selected.col}"]`)
      ?.scrollIntoView({ block: "nearest", inline: "nearest" });
  }, [selected]);

  function startEditing(position: CellPosition, value: string): void {
    if (editing === null || editing.saving !== null || cellAt(position) === null) return;
    setSelected(position);
    setDraft({ position, value });
  }

  function finishEditing(save: boolean, next?: CellPosition): void {
    const current = draftRef.current;
    if (current === null) return;
    draftRef.current = null;
    setDraft(null);
    const cell = cellAt(current.position);
    if (save && editing !== null && cell !== null && current.value !== cell.input) {
      editing.onCommit(sheet, cell.address, current.value);
    }
    if (next !== undefined) setSelected(next);
    tableRef.current?.focus({ preventScroll: true });
  }

  function onGridKeyDown(event: KeyboardEvent<HTMLTableElement>): void {
    if (selected === null || draftRef.current !== null) return;
    const cell = cellAt(selected);
    const moves: Record<string, [number, number]> = {
      ArrowUp: [-1, 0],
      ArrowDown: [1, 0],
      ArrowLeft: [0, -1],
      ArrowRight: [0, 1],
      Tab: [0, event.shiftKey ? -1 : 1],
    };
    const move = moves[event.key];
    if (move !== undefined) {
      event.preventDefault();
      setSelected(step(selected, move[0], move[1]));
    } else if (event.key === "Enter" || event.key === "F2") {
      event.preventDefault();
      startEditing(selected, cell?.input ?? "");
    } else if (event.key === "Delete" || event.key === "Backspace") {
      event.preventDefault();
      if (editing !== null && editing.saving === null && cell !== null && cell.input !== "") {
        editing.onCommit(sheet, cell.address, "");
      }
    } else if (event.key.length === 1 && !event.ctrlKey && !event.metaKey && !event.altKey) {
      // Typing over a cell replaces it, as in Excel.
      event.preventDefault();
      startEditing(selected, event.key);
    } else if (event.key === "Escape") {
      // Lets Tab leave the grid again.
      setSelected(null);
    }
  }

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
          tabIndex={0}
          aria-label={editing === null ? "Sheet" : "Sheet; press Enter or type to edit the selected cell"}
          onKeyDown={onGridKeyDown}
          className="border-separate border-spacing-0 font-mono text-xs text-foreground focus:outline-none"
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
                    const isSelected = selected?.row === row && selected.col === col;
                    const isDraft = draft?.position.row === row && draft.position.col === col;
                    return (
                      <td
                        key={grid.columns[col].label}
                        data-slot={`${row}:${col}`}
                        rowSpan={cell.rowSpan}
                        colSpan={cell.colSpan}
                        title={cell.title}
                        aria-selected={isSelected}
                        onClick={() => setSelected({ row, col })}
                        onDoubleClick={() => startEditing({ row, col }, cell.input)}
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
                          isSelected && "outline-2 -outline-offset-2 outline-ring",
                          editing?.saving === cell.address && "animate-pulse",
                        )}
                      >
                        <div
                          className="max-w-[28rem] truncate"
                          style={{ minWidth: cell.colSpan === 1 ? grid.columns[col].widthPx : undefined }}
                        >
                          {isDraft ? (
                            <input
                              autoFocus
                              aria-label={`Edit ${cell.address}`}
                              value={draft.value}
                              spellCheck={false}
                              onChange={(event) => setDraft({ position: { row, col }, value: event.target.value })}
                              onBlur={() => finishEditing(true)}
                              onKeyDown={(event) => {
                                event.stopPropagation();
                                if (event.key === "Enter") {
                                  event.preventDefault();
                                  finishEditing(true, step({ row, col }, event.shiftKey ? -1 : 1, 0));
                                } else if (event.key === "Tab") {
                                  event.preventDefault();
                                  finishEditing(true, step({ row, col }, 0, event.shiftKey ? -1 : 1));
                                } else if (event.key === "Escape") {
                                  event.preventDefault();
                                  finishEditing(false);
                                }
                              }}
                              className="-my-px w-full bg-transparent p-0 font-mono text-xs not-italic text-foreground outline-none"
                            />
                          ) : (
                            cell.text
                          )}
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

function WorkbookView({
  workbook,
  tabs,
  editing,
  readOnlyReason,
  status,
  onReload,
}: {
  workbook: WorkBook;
  tabs: SheetTab[];
  editing: Editing | null;
  readOnlyReason: string | null;
  status: ViewerStatus | null;
  onReload: () => void;
}) {
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
  const footer = [
    notes.length === 0 ? null : `Showing ${notes.join(", ")}.`,
    readOnlyReason === null ? null : `Read-only: ${readOnlyReason}`,
  ].filter((line) => line !== null);

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
            key={activeName}
            grid={grid}
            frozen={frozen}
            onFreeze={(panes) => setFrozenBySheet((previous) => ({ ...previous, [activeName]: panes }))}
            sheet={activeName}
            editing={editing}
          />
        )}
      </div>
      {status === null ? null : (
        <div
          role="status"
          className={cn(
            "flex shrink-0 items-center gap-3 border-t border-border px-3 py-1 text-xs",
            status.tone === "error" ? "text-destructive" : "text-muted-foreground",
          )}
        >
          <span className="min-w-0 truncate">{status.text}</span>
          {status.reload ? (
            <button type="button" className="shrink-0 underline hover:text-foreground" onClick={onReload}>
              Reload
            </button>
          ) : null}
        </div>
      )}
      {footer.length === 0 ? null : (
        <p className="shrink-0 border-t border-border px-3 py-1 text-xs text-muted-foreground">{footer.join(" ")}</p>
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

/** The file as last read from or written to disk; edits are made against `sha256`. */
interface Opened {
  workbook: WorkBook;
  tabs: SheetTab[];
  sha256: string;
  readOnlyReason: string | null;
}

function decodeBase64(content: string): Uint8Array {
  const binary = atob(content);
  const bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index += 1) bytes[index] = binary.charCodeAt(index);
  return bytes;
}

function errorMessage(cause: unknown): string {
  return cause instanceof Error ? cause.message : String(cause);
}

function SpreadsheetViewer({ path, source }: PluginFileOpenerProps) {
  const rpc = useRpc<typeof rpcContract>();
  const [attempt, setAttempt] = useState(0);
  const [state, setState] = useState<LoadState>({ status: "loading" });
  const [status, setStatus] = useState<ViewerStatus | null>(null);
  const [saving, setSaving] = useState<string | null>(null);
  // `source` is a fresh object per render; its fields are the identity.
  const file = useMemo(
    () => ({ path, source: { ...source } }),
    [path, source.kind, source.threadId, source.environmentId, source.projectId, source.experimental_hostId],
  );

  useEffect(() => {
    let cancelled = false;
    setState({ status: "loading" });
    setStatus(null);
    void (async () => {
      try {
        const opened = await rpc.call("open", file);
        if (cancelled) return;
        if (opened.status === "too-large") {
          const mb = (bytes: number) => `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
          setState({
            status: "error",
            message: `This file is ${mb(opened.sizeBytes)}; previews are limited to 25 MB.`,
            retryable: false,
          });
          return;
        }
        const workbook = parseWorkbook(decodeBase64(opened.content));
        const ready: Opened = {
          workbook,
          tabs: visibleSheets(workbook),
          sha256: opened.sha256,
          readOnlyReason: opened.readOnlyReason,
        };
        setState({ status: "ready", ...ready });
        if (!needsRecalculation(workbook) || opened.recalculateUnavailableReason !== null) return;

        setStatus({ tone: "muted", text: "Calculating formulas with LibreOffice…" });
        const recalculated = await rpc.call("recalculate", file);
        if (cancelled) return;
        if (recalculated.sha256 !== opened.sha256) {
          setStatus({ tone: "error", text: "The file changed while its formulas were calculated.", reload: true });
          return;
        }
        const calculated = parseWorkbook(decodeBase64(recalculated.content));
        // A cell saved meanwhile came back recalculated and newer; keep it.
        setState((previous) =>
          previous.status === "ready" && previous.sha256 === opened.sha256
            ? { ...previous, workbook: calculated }
            : previous,
        );
        setStatus({ tone: "muted", text: "Formula results calculated by LibreOffice; the file has none saved." });
      } catch (cause) {
        if (cancelled) return;
        setState((previous) =>
          previous.status === "ready" ? previous : { status: "error", message: errorMessage(cause), retryable: true },
        );
        setStatus({ tone: "error", text: `Couldn't calculate formulas: ${errorMessage(cause)}` });
      }
    })();
    return () => {
      cancelled = true;
    };
  }, [file, attempt, rpc]);

  async function commit(sheet: string, cell: string, input: string): Promise<void> {
    if (state.status !== "ready" || saving !== null) return;
    setSaving(cell);
    setStatus({ tone: "muted", text: `Saving ${cell}…` });
    try {
      const result = await rpc.call("edit", { ...file, expectedSha256: state.sha256, sheet, cell, input });
      if (result.outcome === "conflict") {
        setStatus({ tone: "error", text: `${cell} was not saved: the file changed on disk.`, reload: true });
        return;
      }
      const workbook = parseWorkbook(decodeBase64(result.content));
      setState({ ...state, workbook, tabs: visibleSheets(workbook), sha256: result.sha256 });
      setStatus({ tone: "muted", text: `Saved ${cell}.` });
    } catch (cause) {
      setStatus({ tone: "error", text: `Couldn't save ${cell}: ${errorMessage(cause)}` });
    } finally {
      setSaving(null);
    }
  }

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
            <button type="button" className={BUTTON_CLASS} onClick={() => setAttempt((value) => value + 1)}>
              Retry
            </button>
          ) : null}
        </div>
      </div>
    );
  }

  return (
    <WorkbookView
      workbook={state.workbook}
      tabs={state.tabs}
      editing={
        state.readOnlyReason === null
          ? { saving, onCommit: (sheet, cell, input) => void commit(sheet, cell, input) }
          : null
      }
      readOnlyReason={state.readOnlyReason}
      status={status}
      onReload={() => setAttempt((value) => value + 1)}
    />
  );
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
