// Read-only spreadsheet viewer registered as a BB file opener. Every color
// comes from BB's theme tokens so the grid follows the active palette.
import { useEffect, useMemo, useState } from "react";
import { definePluginApp } from "@get-bb/plugin-sdk/app";
import type { PluginFileOpenerProps } from "@get-bb/plugin-sdk/app";
import type { WorkBook } from "xlsx";
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

function SheetTable({ grid }: { grid: SheetGrid }) {
  return (
    <table className="border-separate border-spacing-0 font-mono text-xs text-foreground">
      <thead>
        <tr>
          <th
            aria-hidden="true"
            className="sticky left-0 top-0 z-20 border-b border-r border-border bg-card"
          />
          {grid.columns.map((column) => (
            <th
              key={column.label}
              scope="col"
              className="sticky top-0 z-10 border-b border-r border-border bg-card px-2 py-1 text-center font-normal text-muted-foreground"
            >
              {column.label}
            </th>
          ))}
        </tr>
      </thead>
      <tbody>
        {grid.rows.map((row) => (
          <tr key={row.number} className="hover:bg-state-hover">
            <th
              scope="row"
              className="sticky left-0 z-10 border-b border-r border-border bg-card px-2 text-right font-normal tabular-nums text-muted-foreground"
            >
              {row.number}
            </th>
            {row.cells.map((cell, index) =>
              cell === null ? null : (
                <td
                  key={grid.columns[index].label}
                  rowSpan={cell.rowSpan}
                  colSpan={cell.colSpan}
                  title={cell.title}
                  className={cn(
                    "border-b border-r border-border px-2 py-1 align-top",
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
              ),
            )}
          </tr>
        ))}
      </tbody>
    </table>
  );
}

function WorkbookView({ workbook, tabs }: { workbook: WorkBook; tabs: SheetTab[] }) {
  const [selectedName, setSelectedName] = useState(tabs[0].name);
  // A reload or another file can drop the selected sheet; fall back to the first.
  const activeName = tabs.some((tab) => tab.name === selectedName) ? selectedName : tabs[0].name;
  const grid = useMemo(() => buildGrid(workbook.Sheets[activeName]), [workbook, activeName]);
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
          <SheetTable grid={grid} />
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
