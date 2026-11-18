// Reads spreadsheets for the viewer, recalculates formulas and saves cell edits
// through LibreOffice. Files move through `bb.sdk.files`, so a workbook on any
// connected host works; LibreOffice itself runs on the bb server machine.
import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { defineRpcContract, experimental_PluginRpcRefusal, type BbPluginApi } from "@get-bb/plugin-sdk";
import { z } from "zod";
import { LibreOffice, type WorkerRequest } from "./libreoffice";

/** Larger files are not opened: SheetJS parses on the UI thread and bb writes cap at 25 MB. */
const MAX_FILE_BYTES = 25 * 1024 * 1024;

/** LibreOffice export filter per format it can write back. `.xlsb` is read-only. */
const SAVE_FILTERS: Record<string, string> = {
  ".xlsx": "Calc MS Excel 2007 XML",
  ".xlsm": "Calc MS Excel 2007 VBA XML",
  ".xls": "MS Excel 97",
  ".ods": "calc8",
};
const DISPLAY_FILTER = "Calc MS Excel 2007 XML";

const sourceSchema = z.object({
  kind: z.enum(["workspace", "host", "thread-storage"]),
  threadId: z.string().nullable(),
  environmentId: z.string().nullable(),
  projectId: z.string().nullable(),
  experimental_hostId: z.string().nullable().optional(),
});

const fileSchema = { path: z.string().min(1), source: sourceSchema };

export const rpcContract = defineRpcContract({
  open: {
    input: z.object(fileSchema).strict(),
    output: z.discriminatedUnion("status", [
      z.object({ status: z.literal("too-large"), sizeBytes: z.number() }),
      z.object({
        status: z.literal("ok"),
        /** The file's bytes, base64. */
        content: z.string(),
        sha256: z.string(),
        /** Why cells cannot be edited, or `null` when they can. */
        readOnlyReason: z.string().nullable(),
        /** Why formulas cannot be recalculated, or `null` when they can. */
        recalculateUnavailableReason: z.string().nullable(),
      }),
    ]),
  },
  recalculate: {
    input: z.object(fileSchema).strict(),
    output: z.object({
      /** The workbook with every formula recalculated, as `.xlsx`, base64. */
      content: z.string(),
      /** Of the file on disk that was recalculated; it is not written. */
      sha256: z.string(),
    }),
  },
  edit: {
    input: z
      .object({
        ...fileSchema,
        /** The version the edit was made against; a different file on disk is a conflict. */
        expectedSha256: z.string(),
        sheet: z.string().min(1),
        cell: z.string().regex(/^[A-Z]{1,3}[1-9][0-9]{0,6}$/u),
        /** What the user typed, as in a Calc cell; "" clears it. */
        input: z.string().max(32_000),
      })
      .strict(),
    output: z.discriminatedUnion("outcome", [
      z.object({ outcome: z.literal("conflict") }),
      z.object({
        outcome: z.literal("written"),
        /** The saved file with formulas recalculated, base64. */
        content: z.string(),
        sha256: z.string(),
      }),
    ]),
  },
});

type FileSource = z.infer<typeof sourceSchema>;

interface Located {
  hostId: string;
  absolutePath: string;
  rootPath: string;
}

export default function plugin(bb: BbPluginApi): void {
  const libreOffice = new LibreOffice((message) => bb.log.info(message));
  bb.onDispose(() => libreOffice.dispose());

  /** `projects.get` 404s for the personal project; the list includes it. */
  async function project(projectId: string) {
    try {
      return await bb.sdk.projects.get({ projectId });
    } catch {
      const projects = await bb.sdk.projects.list({ includePersonal: true });
      return projects.find((entry) => entry.id === projectId) ?? null;
    }
  }

  function within(rootPath: string, relativePath: string): string {
    const absolutePath = path.posix.resolve(rootPath, relativePath);
    if (!absolutePath.startsWith(rootPath.endsWith("/") ? rootPath : `${rootPath}/`)) {
      throw new experimental_PluginRpcRefusal({ code: "invalid", message: "That path leaves its workspace." });
    }
    return absolutePath;
  }

  /** The host and absolute path BB's file routes would serve for this opener source. */
  async function locate(filePath: string, source: FileSource): Promise<Located> {
    const missing = (message: string) => new experimental_PluginRpcRefusal({ code: "not_found", message });
    switch (source.kind) {
      case "workspace": {
        if (source.environmentId !== null) {
          const environment = await bb.sdk.environments.get({ environmentId: source.environmentId });
          if (environment.path === null) throw missing("This workspace has no directory yet.");
          return {
            hostId: environment.hostId,
            absolutePath: within(environment.path, filePath),
            rootPath: environment.path,
          };
        }
        if (source.projectId !== null) {
          const found = await project(source.projectId);
          const checkout =
            found?.sources.find((entry) =>
              source.experimental_hostId ? entry.hostId === source.experimental_hostId : entry.isDefault,
            ) ?? found?.sources[0];
          if (checkout === undefined) throw missing("This project has no checkout.");
          return {
            hostId: checkout.hostId,
            absolutePath: within(checkout.path, filePath),
            rootPath: checkout.path,
          };
        }
        throw missing("This file has no workspace.");
      }
      case "host": {
        if (source.threadId === null) throw missing("This file has no thread.");
        const thread = await bb.sdk.threads.get({ threadId: source.threadId });
        if (thread.environmentId === null) throw missing("This thread has no machine yet.");
        const environment = await bb.sdk.environments.get({ environmentId: thread.environmentId });
        const absolutePath = path.posix.resolve("/", filePath);
        return { hostId: environment.hostId, absolutePath, rootPath: path.posix.dirname(absolutePath) };
      }
      case "thread-storage": {
        if (source.threadId === null) throw missing("This file has no thread.");
        const location = await bb.sdk.threads.storageLocation({ threadId: source.threadId });
        return {
          hostId: location.hostId,
          absolutePath: within(location.storageRootPath, filePath),
          rootPath: location.storageRootPath,
        };
      }
    }
  }

  async function readBytes(located: Located): Promise<{ bytes: Buffer; sha256: string }> {
    const file = await bb.sdk.files.read({
      hostId: located.hostId,
      path: located.absolutePath,
      rootPath: located.rootPath,
    });
    if (file.sizeBytes > MAX_FILE_BYTES) {
      throw new experimental_PluginRpcRefusal({ code: "invalid", message: "This file is too large to edit here." });
    }
    return { bytes: Buffer.from(file.content, file.contentEncoding), sha256: file.sha256 };
  }

  /** Runs one LibreOffice job on a private copy and returns what it wrote. */
  async function convert(
    bytes: Buffer,
    extensions: { input: string; output: string },
    job: (input: string, output: string) => WorkerRequest,
  ): Promise<Buffer> {
    const directory = await mkdtemp(path.join(tmpdir(), "excel-preview-"));
    try {
      const input = path.join(directory, `in${extensions.input}`);
      const output = path.join(directory, `out${extensions.output}`);
      await writeFile(input, bytes, { mode: 0o600 });
      await libreOffice.run(job(input, output));
      return await readFile(output);
    } finally {
      await rm(directory, { recursive: true, force: true });
    }
  }

  bb.rpc.register(rpcContract, {
    async open({ path: filePath, source }) {
      const located = await locate(filePath, source);
      const file = await bb.sdk.files.read({
        hostId: located.hostId,
        path: located.absolutePath,
        rootPath: located.rootPath,
      });
      if (file.sizeBytes > MAX_FILE_BYTES) return { status: "too-large" as const, sizeBytes: file.sizeBytes };
      const unavailable = await libreOffice.unavailableReason();
      const extension = path.extname(located.absolutePath).toLowerCase();
      return {
        status: "ok" as const,
        content: Buffer.from(file.content, file.contentEncoding).toString("base64"),
        sha256: file.sha256,
        readOnlyReason:
          unavailable ?? (SAVE_FILTERS[extension] === undefined ? "LibreOffice cannot save this format." : null),
        recalculateUnavailableReason: unavailable,
      };
    },

    async recalculate({ path: filePath, source }) {
      const located = await locate(filePath, source);
      const { bytes, sha256 } = await readBytes(located);
      const input = path.extname(located.absolutePath).toLowerCase();
      const recalculated = await convert(bytes, { input, output: ".xlsx" }, (inputPath, outputPath) => ({
        op: "recalculate",
        input: inputPath,
        output: outputPath,
        filter: DISPLAY_FILTER,
      }));
      return { content: recalculated.toString("base64"), sha256 };
    },

    async edit({ path: filePath, source, expectedSha256, sheet, cell, input: typed }) {
      const located = await locate(filePath, source);
      const extension = path.extname(located.absolutePath).toLowerCase();
      const filter = SAVE_FILTERS[extension];
      if (filter === undefined) {
        throw new experimental_PluginRpcRefusal({ code: "invalid", message: "LibreOffice cannot save this format." });
      }
      const { bytes, sha256 } = await readBytes(located);
      if (sha256 !== expectedSha256) return { outcome: "conflict" as const };

      const edited = await convert(bytes, { input: extension, output: extension }, (inputPath, outputPath) => ({
        op: "edit",
        input: inputPath,
        output: outputPath,
        filter,
        edits: [{ sheet, cell, input: typed }],
      }));
      const content = edited.toString("base64");
      const saved = await bb.sdk.files.write({
        hostId: located.hostId,
        path: located.absolutePath,
        rootPath: located.rootPath,
        content,
        contentEncoding: "base64",
        expectedSha256,
      });
      if (saved.outcome === "conflict") return { outcome: "conflict" as const };
      bb.log.info(`edited ${sheet}!${cell} in ${located.absolutePath}`);
      return { outcome: "written" as const, content, sha256: saved.sha256 };
    },
  });
}
