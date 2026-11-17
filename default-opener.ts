// Makes Excel Preview the default opener for spreadsheet extensions on first
// use, without ever overriding a choice the user made.
//
// bb has no registration option for a default opener: with no per-extension
// pin it uses the first matching opener in plugin-id order, so any plugin
// whose id sorts first and claims spreadsheet extensions would win. The pin
// bb reads is the browser-local preference behind Settings → File openers
// (`bb.fileOpenerByExtension`). On its first mount in a browser this script
// pins each spreadsheet extension that has no pin yet, then records that it
// did so and never touches the preference again — so a later "Automatic"
// (which removes the pin) or "Built-in preview" sticks.
// Same scheme as the live-markdown plugin's default-opener.ts.

/** The preference map BB's Settings → File openers reads and writes. */
const PREFERENCE_KEY = "bb.fileOpenerByExtension";

export const SPREADSHEET_EXTENSIONS = ["xlsx", "xlsm", "xlsb", "xls", "ods"] as const;

export const OPENER_ID = "spreadsheet";

/**
 * This plugin was published as `sheet-preview` before it was renamed. A pin to
 * that id now names a missing opener, which bb resolves to its built-in
 * preview, so it counts as unpinned and is re-pointed here.
 */
const PRE_RENAME_OPENER_REF = `sheet-preview:${OPENER_ID}`;
const PRE_RENAME_MARKER_KEY = "bb-plugin:sheet-preview:default-opener-seeded:v1";

/** Pins every unpinned spreadsheet extension once per browser, and records that it ran. */
export function seedDefaultOpener(pluginId: string, storage: Storage = window.localStorage): void {
  const markerKey = `bb-plugin:${pluginId}:default-opener-seeded:v1`;
  if (storage.getItem(markerKey) !== null) return;

  const raw = storage.getItem(PREFERENCE_KEY);
  let preference: Record<string, string> = {};
  if (raw !== null) {
    let parsed: unknown;
    try {
      parsed = JSON.parse(raw);
    } catch {
      return; // Not ours to repair; leave the user's storage exactly as it is.
    }
    if (parsed === null || typeof parsed !== "object" || Array.isArray(parsed)) return;
    // Extension → opener ref; only key presence is read, values are kept as-is.
    preference = parsed as Record<string, string>;
  }

  const seeded = SPREADSHEET_EXTENSIONS.filter(
    (extension) =>
      preference[extension] === undefined || preference[extension] === PRE_RENAME_OPENER_REF,
  );
  if (seeded.length > 0) {
    const next = { ...preference };
    for (const extension of seeded) next[extension] = `${pluginId}:${OPENER_ID}`;
    const value = JSON.stringify(next);
    storage.setItem(PREFERENCE_KEY, value);
    // BB keeps the preference in memory and refreshes it from `storage` events; a
    // write in this same tab fires none, so announce it the way another tab would.
    window.dispatchEvent(
      new StorageEvent("storage", { key: PREFERENCE_KEY, newValue: value, oldValue: raw, storageArea: storage }),
    );
  }
  storage.setItem(markerKey, JSON.stringify({ seeded, at: new Date().toISOString() }));
  storage.removeItem(PRE_RENAME_MARKER_KEY);
}
