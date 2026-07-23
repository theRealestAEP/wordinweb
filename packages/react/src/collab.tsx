import { useEffect, useRef, useState } from "react";
import type { DocxDocument } from "@wordinweb/core";
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
  /** True once joined and the snapshot is loaded. */
  ready: boolean;
  /** Submit a local edit (bookkeeping filled by the connection). */
  submit: (intent: Omit<Intent, "clientId" | "clientSeq" | "base">) => void;
  /** Broadcast this client's cursor/selection. */
  setPresence: (pos: PresencePosition | null) => void;
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
    ready,
    submit: (intent) => connRef.current?.submit(intent),
    setPresence: (pos) => connRef.current?.setPresence(pos),
    presence,
    refused,
  };
}

/** Structural type for the DocxView `collab` prop — mirrors CollabSession so
 * DocxView can accept an injected session via a type-only import (no runtime
 * dependency on the collab engine). */
export type { CollabSession as InjectedCollabSession };
