// One headless LibreOffice per plugin server, driven over the Python UNO bridge
// (`libreoffice_worker.py`). Started on first use, stopped after idle time.
import { spawn, type ChildProcessWithoutNullStreams } from "node:child_process";
import { constants } from "node:fs";
import { access } from "node:fs/promises";
import path from "node:path";
import { createInterface } from "node:readline";
import workerSource from "./libreoffice_worker.py" with { type: "text" };

const IDLE_STOP_MS = 5 * 60 * 1000;
const START_TIMEOUT_MS = 60 * 1000;
const REQUEST_TIMEOUT_MS = 2 * 60 * 1000;
/** A failed probe is retried after this long, so installing LibreOffice needs no reload. */
const PROBE_RETRY_MS = 30 * 1000;

export interface CellEdit {
  sheet: string;
  /** A1 address, e.g. "B5". */
  cell: string;
  /** What the user typed; "" clears the cell. */
  input: string;
}

export type WorkerRequest =
  | { op: "recalculate"; input: string; output: string; filter: string }
  | { op: "edit"; input: string; output: string; filter: string; edits: CellEdit[] };

interface Toolchain {
  soffice: string;
  python: string;
}

type Probe = { ok: true; toolchain: Toolchain } | { ok: false; reason: string; atMs: number };

const SOFFICE_CANDIDATES = [
  "/usr/bin/soffice",
  "/usr/lib/libreoffice/program/soffice",
  "/opt/libreoffice/program/soffice",
  "/snap/bin/libreoffice",
  "/Applications/LibreOffice.app/Contents/MacOS/soffice",
];
// LibreOffice on macOS ships its own Python with the UNO bridge.
const PYTHON_CANDIDATES = ["/Applications/LibreOffice.app/Contents/Resources/python", "python3"];

async function executable(file: string): Promise<boolean> {
  try {
    await access(file, constants.X_OK);
    return true;
  } catch {
    return false;
  }
}

async function onPath(name: string): Promise<string | null> {
  for (const directory of (process.env.PATH ?? "").split(path.delimiter)) {
    if (directory === "") continue;
    const candidate = path.join(directory, name);
    if (await executable(candidate)) return candidate;
  }
  return null;
}

function run(command: string, args: string[]): Promise<boolean> {
  const { promise, resolve } = Promise.withResolvers<boolean>();
  const child = spawn(command, args, { stdio: "ignore" });
  child.on("error", () => resolve(false));
  child.on("exit", (code) => resolve(code === 0));
  return promise;
}

async function probeToolchain(): Promise<Probe> {
  let soffice = (await onPath("soffice")) ?? (await onPath("libreoffice"));
  for (const candidate of SOFFICE_CANDIDATES) {
    if (soffice !== null) break;
    if (await executable(candidate)) soffice = candidate;
  }
  if (soffice === null) {
    return { ok: false, reason: "LibreOffice is not installed on the bb server machine.", atMs: Date.now() };
  }
  for (const python of PYTHON_CANDIDATES) {
    if (await run(python, ["-c", "import uno"])) return { ok: true, toolchain: { soffice, python } };
  }
  return {
    ok: false,
    reason: "LibreOffice's Python bridge is missing (on Debian/Ubuntu: install python3-uno).",
    atMs: Date.now(),
  };
}

interface Pending {
  resolve: () => void;
  reject: (error: Error) => void;
  timer: NodeJS.Timeout;
}

class Worker {
  private nextId = 1;
  private readonly pending = new Map<number, Pending>();
  readonly ready: Promise<void>;
  private exited = false;

  constructor(
    private readonly child: ChildProcessWithoutNullStreams,
    private readonly onExit: (worker: Worker) => void,
    log: (message: string) => void,
  ) {
    const lines = createInterface({ input: child.stdout });
    const started = Promise.withResolvers<void>();
    const markReady = started.resolve;
    const failStart = started.reject;
    this.ready = started.promise;
    const startTimer = setTimeout(() => {
      failStart(new Error("LibreOffice took too long to start."));
      this.stop();
    }, START_TIMEOUT_MS);

    lines.on("line", (line) => {
      let message: { ready?: boolean; id?: number; ok?: boolean; error?: string };
      try {
        message = JSON.parse(line);
      } catch {
        log(`libreoffice: ${line}`);
        return;
      }
      if (message.ready === true) {
        clearTimeout(startTimer);
        markReady();
        return;
      }
      const waiter = message.id === undefined ? undefined : this.pending.get(message.id);
      if (waiter === undefined) return;
      this.pending.delete(message.id!);
      clearTimeout(waiter.timer);
      if (message.ok === true) waiter.resolve();
      else waiter.reject(new Error(message.error ?? "LibreOffice failed."));
    });

    let stderr = "";
    child.stderr.on("data", (chunk: Buffer) => {
      stderr = (stderr + chunk.toString()).slice(-4000);
    });
    child.on("exit", (code, signal) => {
      clearTimeout(startTimer);
      this.exited = true;
      const detail = stderr.trim().split("\n").at(-1) ?? "";
      const error = new Error(
        `LibreOffice stopped (${signal ?? `exit ${code}`})${detail === "" ? "" : `: ${detail}`}`,
      );
      failStart(error);
      for (const waiter of this.pending.values()) {
        clearTimeout(waiter.timer);
        waiter.reject(error);
      }
      this.pending.clear();
      this.onExit(this);
    });
  }

  request(request: WorkerRequest): Promise<void> {
    if (this.exited) return Promise.reject(new Error("LibreOffice is not running."));
    const id = this.nextId++;
    const { promise, resolve, reject } = Promise.withResolvers<void>();
    const timer = setTimeout(() => {
      this.pending.delete(id);
      reject(new Error("LibreOffice took too long; it was restarted."));
      this.child.kill("SIGKILL");
    }, REQUEST_TIMEOUT_MS);
    this.pending.set(id, { resolve, reject, timer });
    this.child.stdin.write(`${JSON.stringify({ id, ...request })}\n`);
    return promise;
  }

  /** Closing stdin makes the worker shut LibreOffice down and remove its profile. */
  stop(): void {
    if (this.exited) return;
    this.child.stdin.end();
    const killTimer = setTimeout(() => this.child.kill("SIGKILL"), 15_000);
    this.child.once("exit", () => clearTimeout(killTimer));
  }
}

export class LibreOffice {
  private probe: Probe | null = null;
  private worker: Worker | null = null;
  private idleTimer: NodeJS.Timeout | undefined;
  /** Requests run one at a time: one document open in one office. */
  private queue: Promise<unknown> = Promise.resolve();

  constructor(private readonly log: (message: string) => void) {}

  /** `null` when LibreOffice can be used, otherwise why not. */
  async unavailableReason(): Promise<string | null> {
    const probe = await this.currentProbe();
    return probe.ok ? null : probe.reason;
  }

  run(request: WorkerRequest): Promise<void> {
    const result = this.queue.then(() => this.runNow(request));
    this.queue = result.catch(() => undefined);
    return result;
  }

  dispose(): void {
    clearTimeout(this.idleTimer);
    this.worker?.stop();
    this.worker = null;
  }

  private async currentProbe(): Promise<Probe> {
    if (this.probe === null || (!this.probe.ok && Date.now() - this.probe.atMs > PROBE_RETRY_MS)) {
      this.probe = await probeToolchain();
    }
    return this.probe;
  }

  private async runNow(request: WorkerRequest): Promise<void> {
    const probe = await this.currentProbe();
    if (!probe.ok) throw new Error(probe.reason);
    clearTimeout(this.idleTimer);
    try {
      const worker = this.worker ?? this.start(probe.toolchain);
      await worker.ready;
      await worker.request(request);
    } finally {
      this.idleTimer = setTimeout(() => this.dispose(), IDLE_STOP_MS);
    }
  }

  private start(toolchain: Toolchain): Worker {
    const child = spawn(toolchain.python, ["-c", workerSource, toolchain.soffice], {
      stdio: ["pipe", "pipe", "pipe"],
    });
    const worker = new Worker(
      child,
      (exited) => {
        if (this.worker === exited) this.worker = null;
      },
      this.log,
    );
    this.worker = worker;
    this.log("libreoffice: started");
    return worker;
  }
}
