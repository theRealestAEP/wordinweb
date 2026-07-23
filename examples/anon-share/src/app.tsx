import { CollabEditor } from "wordinweb/collab";

/**
 * Minimal anon-share demo app (browser-run). The whole collab loop — join,
 * render the live reconciled document, forward local edits, and re-render on
 * broadcasts — is composed by CollabEditor; the app supplies only the
 * connection params.
 *
 * A route like /d/:docId provides the doc id (the magic-link capability) and a
 * per-tab clientId; a party route calls the server's PartyPool.assign()
 * instead. In a real deployment the token comes from the app's own
 * /api/session route (plan doc 07); the dev server runs auth-off.
 */
export function App({ url, docId, clientId }: { url: string; docId: string; clientId: string }) {
  // The main `wordinweb` bundle never imports the collab engine — it arrives
  // only through this `wordinweb/collab` entry (plan doc 07 tree-shaking).
  return <CollabEditor url={url} docId={docId} clientId={clientId} editable />;
}
