import { CollabHub } from "./hub.js";
import { makeDocId } from "./demo.js";
import { blankDocxBytes } from "./blank.js";
import type { IdSidecar, SealedCheckpoint, LineageHead } from "@wordinweb/collab/server";

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
  body: {
    docx?: string;
    sidecar?: IdSidecar;
    codeVerifier?: string;
    blank?: boolean;
    /** The seeder's lineage chain (doc 15) — echoed in welcomes for
     * client-side fast-forward/fork decisions. */
    lineage?: LineageHead[];
    /** Encrypted seed (doc 13): the client-sealed checkpoint + the epoch id
     * it was sealed under (client-minted for encrypted epochs — the keys
     * derive from (K_doc, genesisId), so the id must exist before sealing).
     * The server stores the blob opaquely; size caps still apply. */
    encrypted?: { genesisId: string; checkpoint: SealedCheckpoint };
  };
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
  /**
   * Refuse plaintext seeds, so the only rooms this server can hold are ones
   * it cannot read. Plaintext rooms stay fully supported by the library — a
   * trusted-network deployment may not want key management — but a server
   * that calls itself zero-custody must not hold a readable document, and
   * before this flag the UI was the only thing keeping it from doing so.
   * Nothing stops a crafted request from posting plain bytes to the same
   * route the demo uses.
   *
   * `startZeroCustodyServer` sets this. `startDevServer` does not.
   */
  encryptedOnly?: boolean;
  /**
   * Normalized creator address for the per-IP seed limits (see ip-guard.ts).
   * Supplied by the HTTP layer, which is the only place an address exists.
   * Omitted ⇒ no per-IP accounting, which is what every transport-free test
   * and embedder wants.
   */
  creatorIp?: string;
}

/**
 * Map a hub seed refusal to an HTTP status. `exists` is 409 first-wins and
 * carries the incumbent epoch so the loser can join it; the per-IP throttles
 * are 429 with a Retry-After-shaped story and NO epoch, because no epoch was
 * consulted — the request was refused before any room was looked at.
 */
function seedRefusalResponse(result: { reason: string; genesisId?: string }): SeedHttpResponse {
  if (result.reason === "exists") {
    return { status: 409, body: { error: "exists", genesisId: result.genesisId } };
  }
  // 503, not 429: the global room ceiling is not the caller's fault and not
  // something backing off harder fixes for them specifically. The distinction
  // matters to the client ("try again shortly" vs "you are doing too much")
  // and to whoever is paged — one says add capacity, the other says find the
  // abuser.
  if (result.reason === "server-full") {
    return { status: 503, body: { error: "server-full" } };
  }
  return { status: 429, body: { error: result.reason } };
}

export function handleSeedRequest(
  hub: CollabHub,
  req: SeedHttpRequest,
  opts: SeedHttpOptions = {},
): SeedHttpResponse {
  const maxBytes = opts.maxDocxBytes ?? 10 * 1024 * 1024;

  // Encrypted go-live / revival (doc 13): opaque checkpoint, no content
  // validation possible server-side (clients validate on decrypt — doc 11
  // E2EE amendment); size cap enforced on the ciphertext.
  if (req.body?.encrypted) {
    const { genesisId, checkpoint } = req.body.encrypted;
    if (!genesisId || !checkpoint?.ciphertext || typeof checkpoint.seq !== "number") {
      return { status: 400, body: { error: "bad-encrypted-seed" } };
    }
    if (checkpoint.ciphertext.length > maxBytes) {
      return { status: 413, body: { error: "too-large", maxBytes } };
    }
    const docId = req.method === "POST" ? (opts.mintDocId ?? defaultMintDocId)() : req.docId;
    if (!docId) return { status: 400, body: { error: "missing-doc-id" } };
    const result = hub.seedEncrypted(docId, genesisId, checkpoint, req.body.codeVerifier, opts.creatorIp);
    if (!result.ok) return seedRefusalResponse(result);
    return { status: req.method === "POST" ? 201 : 200, body: { docId, genesisId: result.genesisId, ownerToken: result.ownerToken } };
  }

  // Everything below seeds a PLAINTEXT room, which this server may refuse.
  // Checked before the body is decoded so a refused request never costs a
  // base64 decode or a docx parse.
  if (opts.encryptedOnly) {
    return { status: 400, body: { error: "encrypted-only" } };
  }

  let docx: Uint8Array;
  if (req.body?.blank && req.method === "POST") {
    // "New document" go-live (doc 12 §1): the landing page starts a blank
    // doc without shipping template bytes to the browser first.
    docx = blankDocxBytes();
  } else if (!req.body?.docx || typeof req.body.docx !== "string") {
    return { status: 400, body: { error: "missing-docx" } };
  } else {
    try {
      docx = base64ToBytes(req.body.docx);
    } catch {
      return { status: 400, body: { error: "bad-base64" } };
    }
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
    result = hub.seed(docId, docx, req.body.sidecar, req.body.codeVerifier, req.body.lineage, opts.creatorIp);
  } catch {
    return { status: 400, body: { error: "invalid-docx" } };
  }

  if (!result.ok) {
    // First-wins (doc 12 §5.3): the incumbent's epoch is returned so the
    // losing client can join it and make its lineage decision (case 2).
    // Per-IP throttles come back through the same door as 429s.
    return seedRefusalResponse(result);
  }
  // The owner token (doc 14 §2.5) goes ONLY to the seeder in this response;
  // it is never in the shared link and never broadcast.
  return { status: req.method === "POST" ? 201 : 200, body: { docId, genesisId: result.genesisId, ownerToken: result.ownerToken } };
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
