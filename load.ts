// Fetches an opened file's bytes from BB's own HTTP API, following the same
// source → endpoint mapping as BB's built-in PDF preview.
import type { PluginFileOpenerSource } from "@get-bb/plugin-sdk/app";

/** Larger files are not parsed: SheetJS works on the UI thread. */
export const MAX_FILE_BYTES = 25 * 1024 * 1024;

export type FileRequest =
  | { kind: "raw"; url: string }
  | { kind: "workspace-json"; url: string };

export class FileTooLargeError extends Error {
  constructor(readonly sizeBytes: number) {
    const mb = (bytes: number) => `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
    super(`This file is ${mb(sizeBytes)}; previews are limited to ${mb(MAX_FILE_BYTES)}.`);
  }
}

function withQuery(url: string, query: Record<string, string>): string {
  return `${url}?${new URLSearchParams(query).toString()}`;
}

/** `null` when the source carries no identity BB can serve the file from. */
export function fileRequest(path: string, source: PluginFileOpenerSource): FileRequest | null {
  const encodedPath = path.split("/").map(encodeURIComponent).join("/");
  switch (source.kind) {
    case "workspace":
      if (source.threadId !== null) {
        return {
          kind: "raw",
          url: `/api/v1/threads/${encodeURIComponent(source.threadId)}/worktree/files/${encodedPath}`,
        };
      }
      if (source.environmentId !== null) {
        return {
          kind: "workspace-json",
          url: withQuery(
            `/api/v1/environments/${encodeURIComponent(source.environmentId)}/diff/file`,
            { target: "uncommitted", path, side: "new" },
          ),
        };
      }
      if (source.projectId !== null) {
        return {
          kind: "raw",
          url: withQuery(`/api/v1/projects/${encodeURIComponent(source.projectId)}/files/content`, {
            path,
            ...(source.experimental_hostId ? { hostId: source.experimental_hostId } : {}),
          }),
        };
      }
      return null;
    case "host":
      return source.threadId === null
        ? null
        : {
            kind: "raw",
            url: withQuery(
              `/api/v1/threads/${encodeURIComponent(source.threadId)}/host-files/content`,
              { path },
            ),
          };
    case "thread-storage":
      return source.threadId === null
        ? null
        : {
            kind: "raw",
            url: `/api/v1/threads/${encodeURIComponent(source.threadId)}/thread-storage/files/${encodedPath}`,
          };
  }
}

export async function loadFileBytes(request: FileRequest, signal: AbortSignal): Promise<ArrayBuffer> {
  const response = await fetch(request.url, { credentials: "same-origin", signal });
  if (!response.ok) {
    throw new Error(`File request failed with status ${response.status}.`);
  }
  if (request.kind === "raw") {
    const declaredSize = Number(response.headers.get("content-length"));
    if (declaredSize > MAX_FILE_BYTES) {
      await response.body?.cancel();
      throw new FileTooLargeError(declaredSize);
    }
    const buffer = await response.arrayBuffer();
    if (buffer.byteLength > MAX_FILE_BYTES) throw new FileTooLargeError(buffer.byteLength);
    return buffer;
  }
  // The environment diff endpoint wraps binary content as base64 JSON.
  const body: unknown = await response.json();
  const { content, contentEncoding } =
    typeof body === "object" && body !== null ? (body as Record<string, unknown>) : {};
  if (typeof content !== "string" || contentEncoding !== "base64") {
    throw new Error("The workspace returned an invalid file response.");
  }
  const binary = atob(content);
  if (binary.length > MAX_FILE_BYTES) throw new FileTooLargeError(binary.length);
  const bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index += 1) {
    bytes[index] = binary.charCodeAt(index);
  }
  return bytes.buffer;
}
