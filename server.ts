// The viewer is frontend-only: it reads files through BB's own HTTP API.
// BB requires a server entry, so this factory does nothing.
import type { BbPluginApi } from "@get-bb/plugin-sdk";

export default function plugin(bb: BbPluginApi): void {
  void bb;
}
