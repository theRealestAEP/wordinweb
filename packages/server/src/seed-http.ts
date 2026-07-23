import { CollabHub } from "./hub.js";
import { makeDocId } from "./demo.js";
import type { IdSidecar } from "@wordinweb/collab/server";

/**
 * Transport-free core of the go-live / bring-it-back HTTP endpoints (plan
 * doc 12 §3):
 *
 *   POST /docs            { docx, sidecar? }  → 201 { docId, genesisId }
 *   PUT  /docs/:docId     { docx, sidecar? }  → 200 { docId, genesisId }
 *                                             | 409 { reason, genesisId }  (first-wins)
 *
 * POST is "go live": the server mints an unguessable docId (the magic link IS
 * the capability, doc 11) and seeds a fresh session from the uploaded bytes.
 * PUT is "bring it back live": a bundle-holder revives a KNOWN docId so the
 * link in everyone's chat history keeps working; if someone else revived it
 * first, 409 returns the incumbent epoch and the caller joins it instead
 * (resume case 2). Bytes are held in RAM only — this handler never touches a
 * storage driver (zero custody).
 *
 * Kept free of node:http so it is unit-testable and reusable by any HTTP
 * layer; `wireSeedEndpoints` in cli.ts does the actual routing.
 */

export interface SeedHttpRequest {
  method: "POST" | "PUT";
  /** Required for PUT (the docId being revived); ignored for POST. */
  docId?: string;
  /** Parsed JSON body. `docx` is base64 (JSON has no binary type).
   * `codeVerifier` (doc 13 §7) registers a share-code gate for the epoch —
   * a PBKDF2 output computed client-side; the code itself never crosses
   * the wire. */
  body: { docx?: string; sidecar?: IdSidecar; codeVerifier?: string };
}

export interface SeedHttpResponse {
  status: number;
  body: Record<string, unknown>;
}

export interface SeedHttpOptions {
  /**
   * Cap on the DECODED docx size (doc 11 go-live amendment: size caps apply
   * server-side in every mode; the zip/XML input caps inside
   * DocxDocument.load are the deeper content validation). Default 10 MB.
   */
  maxDocxBytes?: number;
  /** Injectable docId minting (tests); production default is crypto-random
   * ≥128-bit via makeDocId (doc 11 — unguessable IS the access control). */
  mintDocId?: () => string;
}

export function handleSeedRequest(
  hub: CollabHub,
  req: SeedHttpRequest,
  opts: SeedHttpOptions = {},
): SeedHttpResponse {
  const maxBytes = opts.maxDocxBytes ?? 10 * 1024 * 1024;
  if (!req.body?.docx || typeof req.body.docx !== "string") {
    return { status: 400, body: { error: "missing-docx" } };
  }
  let docx: Uint8Array;
  try {
    docx = base64ToBytes(req.body.docx);
  } catch {
    return { status: 400, body: { error: "bad-base64" } };
  }
  if (docx.length > maxBytes) {
    return { status: 413, body: { error: "too-large", maxBytes } };
  }

  const docId = req.method === "POST" ? (opts.mintDocId ?? defaultMintDocId)() : req.docId;
  if (!docId) return { status: 400, body: { error: "missing-doc-id" } };

  // seed() parses the bytes (DocxDocument.load — which enforces the core
  // zip/XML input caps, security gate 3). Unparseable/hostile input is the
  // CALLER's error, reported as 400, and no room is created.
  let result: ReturnType<CollabHub["seed"]>;
  try {
    result = hub.seed(docId, docx, req.body.sidecar, req.body.codeVerifier);
  } catch {
    return { status: 400, body: { error: "invalid-docx" } };
  }

  if (!result.ok) {
    // First-wins (doc 12 §5.3): the incumbent's epoch is returned so the
    // losing client can join it and make its lineage decision (case 2).
    return { status: 409, body: { error: "exists", genesisId: result.genesisId } };
  }
  return { status: req.method === "POST" ? 201 : 200, body: { docId, genesisId: result.genesisId } };
}

function defaultMintDocId(): string {
  return makeDocId((n) => {
    const b = new Uint8Array(n);
    crypto.getRandomValues(b);
    return b;
  });
}

function base64ToBytes(b64: string): Uint8Array {
  // Reject non-base64 up front — atob/Buffer are lenient about whitespace
  // but throw (or produce garbage) on arbitrary text; normalize to a throw.
  if (!/^[A-Za-z0-9+/]*={0,2}$/.test(b64)) throw new Error("bad base64");
  return new Uint8Array(Buffer.from(b64, "base64"));
}
