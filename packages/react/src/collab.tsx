import { useEffect, useMemo, useRef, useState, createElement, type ReactNode } from "react";
import type { DocxDocument } from "@wordinweb/core";
import { DocxView } from "./index.js";
import {
  CollabConnection,
  createWebSocketTransport,
  type Intent,
  type PresencePosition,
} from "@wordinweb/collab/client";

/**
 * React binding for wordinweb collaboration. Imported from the SEPARATE
 * `wordinweb/collab` entry so that non-collab apps never pull the collab
 * engine into their bundle (plan doc 07: unreachable beats shakeable — a
 * local-only `import { DocxView } from "wordinweb"` has no path to this file).
 */

export interface UseCollabOptions {
  /** WebSocket URL of the collab server (ws://… / wss://…). */
  url: string;
  /** Document id to join (the magic-link id, plan doc 11). */
  docId: string;
  /** This client's stable id (persist per browser/tab). */
  clientId: string;
  /** Optional auth token (JWT minted by the app's backend, plan doc 07). */
  token?: string;
  /** Construct the socket. Defaults to `new WebSocket(url)`; injectable for
   * tests / custom transports. */
  createSocket?: (url: string) => WebSocket;
}

export interface CollabSession {
  /** The live document to render (null until the welcome arrives). */
  doc: DocxDocument | null;
  /** Monotonically increases on every reconciled change — a cheap re-render
   * signal (the doc object may be reloaded by reconciliation). */
  version: number;
  /** Increases only when reconciliation RELOADED the document (a true
   * conflict). The editor re-mounts on this; between reloads it updates in
   * place (no flash for the common non-conflicting edits). */
  docEpoch: number;
  /** True once joined and the snapshot is loaded. */
  ready: boolean;
  /** Submit a local edit (bookkeeping filled by the connection). */
  submit: (intent: Omit<Intent, "clientId" | "clientSeq" | "base">) => void;
  /** Broadcast this client's cursor/selection. */
  setPresence: (pos: PresencePosition | null) => void;
  /** Allocate carried node ids (sub-range format / split / insert). */
  allocIds: (n: number) => number[];
  /** Remote participants' latest cursor/selection positions. */
  presence: Record<string, PresencePosition | null>;
  /** Set if the server refused the connection (e.g. version mismatch). */
  refused: string | null;
}

/**
 * Connect to a collab document and track its live state. Returns a
 * CollabSession the app injects into `<DocxView collab={session} />` (the
 * DocxView `collab` prop is typed structurally, so the react package needs no
 * runtime dependency on this module).
 */
export function useCollab(opts: UseCollabOptions): CollabSession {
  const { url, docId, clientId, token, createSocket } = opts;
  const connRef = useRef<CollabConnection | null>(null);
  const [doc, setDoc] = useState<DocxDocument | null>(null);
  const [version, setVersion] = useState(0);
  const [docEpoch, setDocEpoch] = useState(0);
  const [ready, setReady] = useState(false);
  const [presence, setPresence] = useState<Record<string, PresencePosition | null>>({});
  const [refused, setRefused] = useState<string | null>(null);

  useEffect(() => {
    const socket = (createSocket ?? ((u: string) => new WebSocket(u)))(url);
    const transport = createWebSocketTransport(socket);
    const conn = new CollabConnection(transport, clientId, {
      onChange: () => {
        setDoc(conn.doc);
        setReady(conn.ready);
        setVersion((v) => v + 1); // signal a re-render on every reconciled change
        setDocEpoch(conn.docEpoch); // bumps only on a reload (true conflict)
      },
      onPresence: (participant, pos) =>
        setPresence((prev) => ({ ...prev, [participant]: pos })),
      onRefused: (reason) => setRefused(reason),
    });
    connRef.current = conn;
    conn.join(docId, token);
    return () => {
      connRef.current = null;
    };
    // Reconnect when the target document or endpoint changes.
  }, [url, docId, clientId, token, createSocket]);

  return {
    doc,
    version,
    docEpoch,
    ready,
    submit: (intent) => connRef.current?.submit(intent),
    setPresence: (pos) => connRef.current?.setPresence(pos),
    allocIds: (n) => connRef.current?.allocIds(n) ?? [],
    presence,
    refused,
  };
}

/**
 * A complete collaborative editor: joins a document over the collab server and
 * renders its live, reconciled state, forwarding local edits as intents. This
 * is the inbound + outbound integration composed for you — the app supplies
 * only the connection params.
 *
 * Rendering strategy: DocxView renders the replica's LIVE document object
 * directly (`collab.doc`). Broadcasts mutate that same instance in place and
 * bump `version`, which DocxView receives as `renderSignal` and turns into a
 * single in-place repaint — no per-broadcast serialize/parse round-trip. The
 * `source` bytes are computed only on a true-conflict reload (keyed on
 * `docEpoch`) as the placeholder/fallback. The visual rendering runs in the
 * browser (as all of DocxView does); the protocol/convergence/binding it rides
 * on are covered by the headless test suites.
 */
export function CollabEditor(opts: UseCollabOptions & { editable?: boolean }): ReactNode {
  const session = useCollab(opts);
  // Placeholder bytes only — DocxView renders session.doc directly (below), so
  // this is re-serialized ONLY on a reload (docEpoch), never per broadcast.
  const bytes = useMemo(
    () => (session.doc ? session.doc.save() : null),
    [session.doc, session.docEpoch],
  );

  if (session.refused) return createElement("div", { className: "dxw-collab-refused" }, `Please refresh — ${session.refused}.`);
  if (!session.ready || !bytes || !session.doc) return createElement("div", { className: "dxw-collab-connecting" }, "Connecting…");

  return createElement(DocxView, {
    source: bytes,
    // Render the live doc object directly; repaint in place on each version
    // bump. submit + presence + id allocator flow out; DocxView draws carets.
    collab: {
      submit: session.submit,
      presence: session.presence,
      allocIds: session.allocIds,
      doc: session.doc,
      renderSignal: session.version,
    },
    editable: opts.editable ?? true,
    // Re-key only on docEpoch (a true-conflict reload) — NOT on every change.
    // Between reloads the live doc mutates in place and the key stays stable,
    // so DocxView repaints in place instead of re-mounting (no flash/jump).
    key: session.docEpoch,
  });
}

/** Structural type for the DocxView `collab` prop — mirrors CollabSession so
 * DocxView can accept an injected session via a type-only import (no runtime
 * dependency on the collab engine). */
export type { CollabSession as InjectedCollabSession };
