import { useCollab } from "wordinweb/collab";
import { DocxView } from "wordinweb";

/**
 * Minimal anon-share demo app (browser-run). Shows the collab wiring only —
 * the visual editor is browser-verified; the protocol/convergence/binding it
 * rides on are covered by the headless test suites.
 *
 * A route like /d/:docId provides the doc id (the magic-link capability) and
 * a per-tab clientId; a party route calls the server's PartyPool.assign()
 * instead. The token would come from the app's own /api/session route in a
 * real deployment (plan doc 07); the dev server runs auth-off.
 */
export function App({ url, docId, clientId }: { url: string; docId: string; clientId: string }) {
  const session = useCollab({ url, docId, clientId });

  if (session.refused) {
    return <div className="notice">Please refresh — {session.refused}.</div>;
  }
  if (!session.ready) {
    return <div className="notice">Connecting…</div>;
  }

  return (
    <div className="editor">
      {/* DocxView renders the collab session's live document and forwards
          local edits as intents (the collab prop is injected, so the main
          `wordinweb` bundle never imports the collab engine — doc 07). */}
      <DocxView collab={session} editable />
      <RemoteCursors presence={session.presence} />
    </div>
  );
}

function RemoteCursors({ presence }: { presence: Record<string, unknown> }) {
  const n = Object.values(presence).filter((p) => p != null).length;
  return <div className="presence">{n} other {n === 1 ? "person" : "people"} here</div>;
}
