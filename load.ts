// Fetches an opened file's bytes from BB's own HTTP API, following the same
// source → endpoint mapping as BB's built-in PDF preview.
import type { PluginFileOpenerSource } from "@get-bb/plugin-sdk/app";

/** Larger files are not parsed: SheetJS works on the UI thread. */
export const MAX_FILE_BYTES = 25 * 1024 * 1024;

export class FileTooLargeError extends Error {
  constructor(readonly sizeBytes: number) {
    const mb = (bytes: number) => `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
    super(`This file is ${mb(sizeBytes)}; previews are limited to ${mb(MAX_FILE_BYTES)}.`);
  }
}

/** `null` when the source carries no identity BB can serve the file from. */
export function fileUrl(path: string, source: PluginFileOpenerSource): string | null {
  // Routes take the path as URL segments; absolute host paths drop the leading slash.
  const encodedPath = path.replace(/^\/+/u, "").split("/").map(encodeURIComponent).join("/");
  switch (source.kind) {
    case "workspace":
      if (source.environmentId !== null) {
        return `/api/v1/environments/${encodeURIComponent(source.environmentId)}/files/${encodedPath}`;
      }
      if (source.projectId !== null) {
        const project = `/api/v1/projects/${encodeURIComponent(source.projectId)}`;
        return source.experimental_hostId
          ? `${project}/hosts/${encodeURIComponent(source.experimental_hostId)}/files/${encodedPath}`
          : `${project}/files/${encodedPath}`;
      }
      return null;
    case "host":
      return source.threadId === null
        ? null
        : `/api/v1/threads/${encodeURIComponent(source.threadId)}/host-files/${encodedPath}`;
    case "thread-storage":
      return source.threadId === null
        ? null
        : `/api/v1/threads/${encodeURIComponent(source.threadId)}/thread-storage/files/${encodedPath}`;
  }
}

export async function loadFileBytes(url: string, signal: AbortSignal): Promise<ArrayBuffer> {
  const response = await fetch(url, { credentials: "same-origin", signal });
  if (!response.ok) {
    throw new Error(`File request failed with status ${response.status}.`);
  }
  const declaredSize = Number(response.headers.get("content-length"));
  if (declaredSize > MAX_FILE_BYTES) {
    await response.body?.cancel();
    throw new FileTooLargeError(declaredSize);
  }
  const buffer = await response.arrayBuffer();
  if (buffer.byteLength > MAX_FILE_BYTES) throw new FileTooLargeError(buffer.byteLength);
  return buffer;
}
