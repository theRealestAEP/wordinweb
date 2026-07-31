import { DocxDocument } from "@wordinweb/core";
import { CollabConnection } from "./connection.js";
import { Intent } from "./intents.js";
import { PresencePosition } from "./protocol.js";

/**
 * The seam between an editor view and a CollabConnection (plan doc 06/07).
 * An editor implementation (the DOM DocxEditor, or a test double) provides
 * this; `bindEditor` wires it to the connection so local edits become intents
 * and remote broadcasts/presence drive the view. Framework-agnostic and
 * DOM-free, so the wiring is unit-testable without a browser — only the
 * concrete DocxEditor implementation of EditorBridge needs a browser.
 */
export interface EditorBridge {
  /** Point the editor at the connection's live document (called on connect
   * and after each remote reconciliation) and re-render. */
  setDocument(doc: DocxDocument): void;
  /** Register a handler the editor calls when the user performs an edit,
   * supplying the intent minus its wire bookkeeping (clientId/clientSeq/base
   * — the connection fills those). Return unsubscribe. */
  onLocalIntent(handler: (intent: Omit<Intent, "clientId" | "clientSeq" | "base">) => void): () => void;
  /** Register a handler for the local cursor/selection changing, so presence
   * can be broadcast. Return unsubscribe. Optional. */
  onLocalPresence?(handler: (pos: PresencePosition | null) => void): () => void;
  /** Render a remote participant's cursor/selection (or clear it: null).
   * Optional (an editor may not draw remote cursors). */
  setRemotePresence?(participant: string, pos: PresencePosition | null): void;
}

export interface EditorBinding {
  /** Tear down the binding (unsubscribes editor handlers). */
  dispose(): void;
}

/**
 * Wire an EditorBridge to a CollabConnection: the editor's local intents are
 * submitted (optimistically applied + sent), and the connection's document
 * changes / presence updates drive the editor. The connection must already be
 * constructed; `join` is the caller's responsibility.
 */
export function bindEditor(connection: CollabConnection, bridge: EditorBridge): EditorBinding {
  const unsubscribers: Array<() => void> = [];

  // Push document changes (welcome + reconciled broadcasts) into the editor.
  const pushDoc = (): void => {
    const doc = connection.doc;
    if (doc) bridge.setDocument(doc);
  };
  connection.setCallbacks({
    onChange: pushDoc,
    onPresence: (participant, pos) => bridge.setRemotePresence?.(participant, pos),
  });
  pushDoc();

  // Local edits → intents.
  unsubscribers.push(
    bridge.onLocalIntent((intent) => {
      connection.submit(intent);
      pushDoc();
    }),
  );

  // Local presence → broadcast.
  if (bridge.onLocalPresence) {
    unsubscribers.push(bridge.onLocalPresence((pos) => connection.setPresence(pos)));
  }

  return {
    dispose(): void {
      for (const u of unsubscribers) u();
    },
  };
}
