import type { Intent } from "./intents.js";

/**
 * E2EE primitives for the blind-sequencer mode (plan doc 13). Pure WebCrypto
 * (available in browsers and Node ≥16 globals) — no dependencies, usable
 * from both the client and tests.
 *
 * Key model (doc 13 §1): a per-DOCUMENT master key `K_doc` minted at
 * go-live, carried only in the share link's URL fragment. Per-epoch subkeys
 * are DERIVED — `K_epoch = HKDF(K_doc, genesisId [+ share code])` — then
 * domain-separated into K_content (intents/checkpoints) and K_media
 * (blobs). Epoch derivation kills cross-epoch replay (round-4 F5a) and
 * resets the GCM IV budget per epoch (F21); nobody ever redistributes keys,
 * because anyone with the link + the (public) epoch id derives the same
 * values.
 *
 * AAD discipline (round-4 F5): every ciphertext is bound to its exact
 * position in the protocol — intents to (docId, genesisId, clientId,
 * clientSeq, base); checkpoints to (docId, genesisId, seq). `base` is
 * authenticated because it is a TRANSFORM INPUT: an unauthenticated base
 * would let a keyless server shift where an edit applies.
 */

/** The plaintext bookkeeping + opaque body the blind sequencer handles. */
export interface IntentEnvelope {
  clientId: string;
  clientSeq: number;
  base: number;
  /** base64 12-byte GCM IV (fresh per envelope). */
  iv: string;
  /** base64 AES-256-GCM ciphertext (+tag) of the JSON-serialized intent. */
  ciphertext: string;
}

export interface EpochKeys {
  kContent: CryptoKey;
  kMedia: CryptoKey;
}

const te = new TextEncoder();
const td = new TextDecoder();

export function bytesToB64(bytes: Uint8Array): string {
  let bin = "";
  for (let i = 0; i < bytes.length; i++) bin += String.fromCharCode(bytes[i]);
  return typeof btoa === "function" ? btoa(bin) : Buffer.from(bytes).toString("base64");
}
export function b64ToBytes(b64: string): Uint8Array {
  const bin = typeof atob === "function" ? atob(b64) : Buffer.from(b64, "base64").toString("binary");
  const out = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
  return out;
}

/** Mint a fresh 256-bit document master key (go-live, doc 13 §1). Returned
 * base64url — the exact string that rides the URL fragment (`#k=`). */
export function mintDocKey(): string {
  const bytes = new Uint8Array(32);
  crypto.getRandomValues(bytes);
  return bytesToB64(bytes).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
}

/** Parse `#k=<key>` out of a URL fragment; null if absent. The PRESENCE of
 * this is what fixes the doc's mode client-side (doc 13 §6: `#k` present ⇒
 * encrypted, always — a welcome claiming otherwise is hard-refused). */
export function docKeyFromFragment(fragment: string): string | null {
  const m = /(?:^|[#&])k=([A-Za-z0-9_-]+)/.exec(fragment);
  return m ? m[1] : null;
}

/**
 * Derive this epoch's working keys. Share code (doc 13 §7) mixes in when
 * set: a leaked link without the code cannot decrypt anything. The code is
 * stretched by the CALLER via stretchShareCode (PBKDF2) — this function
 * just mixes; keeping stretching separate lets the UI do it once.
 */
export async function deriveEpochKeys(
  docKeyB64url: string,
  genesisId: string,
  stretchedCode?: Uint8Array,
): Promise<EpochKeys> {
  const raw = b64ToBytes(docKeyB64url.replace(/-/g, "+").replace(/_/g, "/"));
  const master = await crypto.subtle.importKey("raw", raw as BufferSource, "HKDF", false, ["deriveKey"]);
  const salt = te.encode(`wordinweb-epoch:${genesisId}`);
  const mix = stretchedCode ?? new Uint8Array(0);
  const derive = (info: string) =>
    crypto.subtle.deriveKey(
      { name: "HKDF", hash: "SHA-256", salt: salt as BufferSource, info: concat(te.encode(info), mix) as BufferSource },
      master,
      { name: "AES-GCM", length: 256 },
      false,
      ["encrypt", "decrypt"],
    );
  return { kContent: await derive("content"), kMedia: await derive("media") };
}

/** PBKDF2-SHA256 stretch of the share code (doc 13 §7): 600k iterations,
 * WebCrypto-native (no WASM dep), salt bound to the docId so a precomputed
 * table for one doc is useless for another. Deliberately NOT epoch-salted:
 * the hello must carry the proof BEFORE the welcome reveals the epoch id
 * (chicken-and-egg otherwise); epoch-binding of the KEYS happens in
 * deriveEpochKeys' HKDF salt instead, which is where it matters. */
export async function stretchShareCode(code: string, docId: string): Promise<Uint8Array> {
  const key = await crypto.subtle.importKey("raw", te.encode(code) as BufferSource, "PBKDF2", false, ["deriveBits"]);
  const bits = await crypto.subtle.deriveBits(
    { name: "PBKDF2", hash: "SHA-256", salt: te.encode(`wordinweb-code:${docId}`) as BufferSource, iterations: 600_000 },
    key,
    256,
  );
  return new Uint8Array(bits);
}

/** Seal an intent into an envelope under K_content (doc 13 §2). The AAD
 * binds every plaintext bookkeeping field the transform depends on. */
export async function sealIntent(
  kContent: CryptoKey,
  docId: string,
  genesisId: string,
  intent: Intent,
): Promise<IntentEnvelope> {
  const iv = new Uint8Array(12);
  crypto.getRandomValues(iv);
  const aad = intentAad(docId, genesisId, intent.clientId, intent.clientSeq, intent.base);
  const ct = await crypto.subtle.encrypt(
    { name: "AES-GCM", iv: iv as BufferSource, additionalData: aad as BufferSource },
    kContent,
    te.encode(JSON.stringify(intent)) as BufferSource,
  );
  return {
    clientId: intent.clientId,
    clientSeq: intent.clientSeq,
    base: intent.base,
    iv: bytesToB64(iv),
    ciphertext: bytesToB64(new Uint8Array(ct)),
  };
}

/** Open an envelope; throws on ANY tamper — wrong epoch, altered base,
 * spliced clientSeq, flipped ciphertext bit. Callers treat a throw as a
 * deterministic no-op for that seq (doc 13 §2: garbage from a malicious
 * participant no-ops identically on every honest client). */
export async function openIntent(
  kContent: CryptoKey,
  docId: string,
  genesisId: string,
  env: IntentEnvelope,
): Promise<Intent> {
  const aad = intentAad(docId, genesisId, env.clientId, env.clientSeq, env.base);
  const pt = await crypto.subtle.decrypt(
    { name: "AES-GCM", iv: b64ToBytes(env.iv) as BufferSource, additionalData: aad as BufferSource },
    kContent,
    b64ToBytes(env.ciphertext) as BufferSource,
  );
  const intent = JSON.parse(td.decode(pt)) as Intent;
  // Belt-and-braces: the plaintext bookkeeping must MATCH the sealed body
  // (AAD already guarantees it — this catches a buggy honest client early).
  if (intent.clientId !== env.clientId || intent.clientSeq !== env.clientSeq || intent.base !== env.base) {
    throw new Error("envelope bookkeeping mismatch");
  }
  return intent;
}

/** Seal a checkpoint bundle (doc 13 §3): AAD binds (docId, genesisId, seq)
 * so a stored checkpoint cannot be replayed at another position (F5c). The
 * body carries the docx + sidecar + the canonical docHash for joiner
 * cross-checking (blocker-2 verification). */
export async function sealCheckpoint(
  kContent: CryptoKey,
  docId: string,
  genesisId: string,
  seq: number,
  body: { docx: string; sidecar: unknown; docHash: string },
): Promise<{ iv: string; ciphertext: string }> {
  const iv = new Uint8Array(12);
  crypto.getRandomValues(iv);
  const aad = te.encode(`cp:${docId}:${genesisId}:${seq}`);
  const ct = await crypto.subtle.encrypt(
    { name: "AES-GCM", iv: iv as BufferSource, additionalData: aad as BufferSource },
    kContent,
    te.encode(JSON.stringify(body)) as BufferSource,
  );
  return { iv: bytesToB64(iv), ciphertext: bytesToB64(new Uint8Array(ct)) };
}

export async function openCheckpoint(
  kContent: CryptoKey,
  docId: string,
  genesisId: string,
  seq: number,
  sealed: { iv: string; ciphertext: string },
): Promise<{ docx: string; sidecar: unknown; docHash: string }> {
  const aad = te.encode(`cp:${docId}:${genesisId}:${seq}`);
  const pt = await crypto.subtle.decrypt(
    { name: "AES-GCM", iv: b64ToBytes(sealed.iv) as BufferSource, additionalData: aad as BufferSource },
    kContent,
    b64ToBytes(sealed.ciphertext) as BufferSource,
  );
  return JSON.parse(td.decode(pt));
}

function intentAad(docId: string, genesisId: string, clientId: string, clientSeq: number, base: number): Uint8Array {
  return te.encode(`in:${docId}:${genesisId}:${clientId}:${clientSeq}:${base}`);
}

function concat(a: Uint8Array, b: Uint8Array): Uint8Array {
  const out = new Uint8Array(a.length + b.length);
  out.set(a, 0);
  out.set(b, a.length);
  return out;
}
