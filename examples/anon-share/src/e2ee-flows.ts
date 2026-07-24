import {
  mintDocKey,
  deriveEpochKeys,
  sealCheckpoint,
  stretchShareCode,
  bytesToB64,
  docHash,
  type DocBundle,
} from "wordinweb/collab";
import { DocxDocument } from "wordinweb";

/**
 * The demo's encrypted go-live / bring-it-back flows (plan doc 13). Both
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

async function sealSeed(
  httpBase: string,
  docId: string,
  docKey: string,
  docx: Uint8Array,
  sidecar: unknown,
  shareCode?: string,
): Promise<{ status: number; genesisId: string; ownerToken?: string }> {
  const genesisId = `g_${randHex(16)}`;
  const stretched = shareCode ? await stretchShareCode(shareCode, docId) : undefined;
  const keys = await deriveEpochKeys(docKey, genesisId, stretched);
  const doc = DocxDocument.load(docx);
  const sealed = await sealCheckpoint(keys.kContent, docId, genesisId, 0, {
    docx: bytesToB64(docx),
    sidecar,
    docHash: await docHash(doc as never), // duplicate d.ts identities across entries; same runtime class
  });
  const codeVerifier = stretched ? btoa(String.fromCharCode(...stretched)) : undefined;
  const res = await fetch(`${httpBase}/docs/${docId}`, {
    method: "PUT",
    headers: { "content-type": "application/json" },
    body: JSON.stringify({ encrypted: { genesisId, checkpoint: { seq: 0, ...sealed } }, codeVerifier }),
  });
  const body = (await res.json()) as { genesisId: string; ownerToken?: string };
  return { status: res.status, genesisId: body.genesisId, ownerToken: body.ownerToken };
}

/** Encrypted go-live: seal the CURRENT local document's bytes (the ones the
 * user just edited in the local editor), mint id + key, PUT. Returns what the
 * URL needs: `?doc=<id>#k=<key>`. The bytes are the genesis checkpoint (seq 0),
 * so the id table IS this docx's parse order — no sidecar needed (null is
 * honest here; a REVIVAL below always carries the bundle's sidecar). Callers
 * that want a blank session pass the bytes from `GET {httpBase}/blank`. */
export async function goLiveEncrypted(
  httpBase: string,
  docx: Uint8Array,
  shareCode?: string,
): Promise<{ docId: string; docKey: string; ownerToken?: string }> {
  const docId = `d_${randHex(16)}`;
  const docKey = mintDocKey();
  const seeded = await sealSeed(httpBase, docId, docKey, docx, null, shareCode);
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
