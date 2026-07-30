import {
  mintDocKey,
  deriveEpochKeys,
  sealCheckpoint,
  stretchShareCode,
  bytesToB64,
  docHash,
  mediaAddressesOf,
  type DocBundle,
} from "wordinweb/collab";
import { DocxDocument } from "wordinweb";

/**
 * The demo's encrypted go-live / bring-it-back flows. Both
 * seal CLIENT-side — the server cannot (it has no keys) — and both use PUT
 * with a CLIENT-minted docId: the checkpoint AAD binds the docId, so the id
 * must exist before sealing (same reason the epoch id is client-minted in
 * encrypted mode). Unguessability is preserved — the id is 128-bit browser
 * crypto, exactly like the server's own minting.
 */

function randHex(bytes: number): string {
  const b = new Uint8Array(bytes);
  crypto.getRandomValues(b);
  return [...b].map((x) => x.toString(16).padStart(2, "0")).join("");
}

/** Where the go-live pipeline currently is, for progress UI. */
export type GoLivePhase = "encrypt" | "upload";

/** Going live failed in a way worth different words in the UI. */
export class GoLiveError extends Error {
  constructor(
    message: string,
    /** HTTP status, when the server answered at all. */
    readonly status?: number,
    /** The server's ciphertext cap (413 responses carry it). */
    readonly maxBytes?: number,
  ) {
    super(message);
  }
}

/**
 * Let the browser PAINT between phases. Setting state and then doing
 * synchronous work in the same tick means the progress text never reaches
 * the screen (the fea7e44 lesson: the parse spinner needed the same two
 * frames). Two frames, not one — a single rAF callback runs BEFORE the paint
 * it is scheduled alongside.
 */
const paintYield = (): Promise<void> =>
  new Promise((resolve) => {
    const frame = typeof requestAnimationFrame === "function"
      ? requestAnimationFrame
      : (fn: () => void) => setTimeout(fn, 0);
    frame(() => frame(() => resolve()));
  });

async function sealSeed(
  httpBase: string,
  docId: string,
  docKey: string,
  docx: Uint8Array,
  sidecar: unknown,
  shareCode?: string,
  onPhase?: (phase: GoLivePhase) => void,
): Promise<{ status: number; genesisId: string; ownerToken?: string; maxBytes?: number }> {
  const genesisId = `g_${randHex(16)}`;
  const now = () => performance.now();
  const t: Record<string, number> = {};
  onPhase?.("encrypt");
  await paintYield();
  let t0 = now();
  const stretched = shareCode ? await stretchShareCode(shareCode, docId) : undefined;
  t.stretchMs = now() - t0;
  t0 = now();
  const keys = await deriveEpochKeys(docKey, genesisId, stretched);
  t.deriveMs = now() - t0;
  // The synchronous blocks below (re-parse, canonical serialise, base64) are
  // each yielded ahead of, so the phase text and spinner stay painted.
  await paintYield();
  t0 = now();
  const doc = DocxDocument.load(docx);
  t.parseMs = now() - t0;
  t0 = now();
  const hash = await docHash(doc as never); // duplicate d.ts identities across entries; same runtime class
  t.hashMs = now() - t0;
  await paintYield();
  t0 = now();
  const sealed = await sealCheckpoint(keys.kContent, docId, genesisId, 0, {
    docx: bytesToB64(docx),
    sidecar,
    docHash: hash,
    // Doc 16 §6 late-join: carry the addresses of any media the seeded
    // document already registers, so joiners can fetch bytes the seed bytes
    // themselves may not contain (a revival's bundle can hold registrations
    // whose pixels only participants have).
    mediaMeta: mediaAddressesOf(doc as never),
  });
  t.sealMs = now() - t0;
  const codeVerifier = stretched ? btoa(String.fromCharCode(...stretched)) : undefined;
  onPhase?.("upload");
  await paintYield();
  t0 = now();
  const res = await fetch(`${httpBase}/docs/${docId}`, {
    method: "PUT",
    headers: { "content-type": "application/json" },
    body: JSON.stringify({ encrypted: { genesisId, checkpoint: { seq: 0, ...sealed } }, codeVerifier }),
  });
  t.uploadMs = now() - t0;
  // One grep-able line per seal, in the shape scripts/bench-golive.mjs and
  // perf-report.mjs already speak — so "going live chugs" stays measurable
  // in the field rather than anecdotal.
  // eslint-disable-next-line no-console
  console.log(
    `STRESS-METRIC golive-seal docxBytes=${docx.byteLength} status=${res.status} ` +
      Object.entries(t).map(([k, v]) => `${k}=${v.toFixed(2)}`).join(" ") +
      ` totalMs=${Object.values(t).reduce((a, b) => a + b, 0).toFixed(2)}`,
  );
  const body = (await res.json().catch(() => ({}))) as { genesisId?: string; ownerToken?: string; maxBytes?: number };
  return { status: res.status, genesisId: body.genesisId ?? genesisId, ownerToken: body.ownerToken, maxBytes: body.maxBytes };
}

/** Encrypted go-live: seal the CURRENT local document's bytes (the ones the
 * user just edited in the local editor), mint id + key, PUT. Returns what the
 * URL needs: `?doc=<id>#k=<key>`. The bytes are the genesis checkpoint (seq 0),
 * so the id table IS this docx's parse order — no sidecar needed (null is
 * honest here; a REVIVAL below always carries the bundle's sidecar). Callers
 * that want a blank session pass the bytes from `GET {httpBase}/blank`.
 *
 * THROWS a GoLiveError on refusal instead of returning garbage: a 413 used
 * to fall through as `docKey` with no session behind it, and the app would
 * "join" a document the server had refused — landing on a misleading
 * "session has ended" screen. */
export async function goLiveEncrypted(
  httpBase: string,
  docx: Uint8Array,
  shareCode?: string,
  onPhase?: (phase: GoLivePhase) => void,
): Promise<{ docId: string; docKey: string; ownerToken?: string }> {
  const docId = `d_${randHex(16)}`;
  const docKey = mintDocKey();
  const seeded = await sealSeed(httpBase, docId, docKey, docx, null, shareCode, onPhase);
  if (seeded.status !== 200 && seeded.status !== 201) {
    if (seeded.status === 413) {
      const mb = seeded.maxBytes ? ` (the limit is about ${Math.round((seeded.maxBytes * 0.55) / 1024 / 1024)} MB of document)` : "";
      throw new GoLiveError(`This document is too large to share live${mb}. You can keep editing it here, and share a smaller copy.`, 413, seeded.maxBytes);
    }
    throw new GoLiveError(`The server refused to start the session (${seeded.status}). Try again in a moment.`, seeded.status);
  }
  return { docId, docKey, ownerToken: seeded.ownerToken };
}

/** Encrypted "Bring it back live": re-seed from this browser's bundle under
 * a FRESH epoch (doc 12 §5.3). First-wins: a 409 means someone else revived
 * first — the caller just joins (case 2 handles the rest). */
export async function reviveEncrypted(
  httpBase: string,
  bundle: DocBundle,
  docKey: string,
  shareCode?: string,
): Promise<void> {
  await sealSeed(httpBase, bundle.docId, docKey, bundle.confirmedBytes, bundle.confirmedSidecar, shareCode);
}
