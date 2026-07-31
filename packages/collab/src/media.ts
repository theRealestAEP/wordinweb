import type { DocxDocument } from "@wordinweb/core";

/**
 * Media transfer client (plan doc 16 §5) — the half that was specced and
 * left unbuilt while the server relay shipped.
 *
 * Bytes NEVER ride the WebSocket sequencer: the intent carries only an
 * address (`blobSha`) and the blob travels over HTTP. Everything here is
 * organised around one rule from doc 16 §1.1 — nobody is ever TOLD a hash.
 * The placer commits `blobSha` inside the sequenced, authenticated intent;
 * from then on every participant re-derives sha256 itself and compares
 * against that commitment. A deliverer's claim about what it is delivering
 * is never consulted, so a hostile relay or peer can at worst withhold
 * bytes, never substitute them.
 *
 * ENVELOPE (as SHIPPED, which differs from doc 16 §1's prose): the blob is
 * the BARE AES-GCM ciphertext and the 12-byte IV rides in the intent's `iv`
 * field — the shape doc 16 §2 specifies and `sealMediaBlob` implements.
 * §1/§5.1's "prepend IV to the blob" is stale text. The sha is over the
 * ciphertext either way, which is what lets the blind relay verify it.
 */

/** One registered media part's address, as it travels to a late joiner:
 * plaintext rooms in `welcome.media`, encrypted rooms inside the SEALED
 * checkpoint body (the server must not learn part structure). */
export interface MediaAddress {
  part: string;
  sha: string;
  iv?: string;
}

/** The late-join address map for a document: every registered part whose sha
 * is actually KNOWN. Parse-derived holes (sha === "") are omitted — an empty
 * address is not an address, and passing one on would only teach the joiner
 * the same nothing. Mirrors the hub's filter for plaintext welcomes. */
export function mediaAddressesOf(doc: DocxDocument): MediaAddress[] {
  const out: MediaAddress[] = [];
  for (const [part, meta] of doc.mediaMeta) {
    if (!meta.sha) continue;
    out.push({ part, sha: meta.sha, ...(meta.iv ? { iv: meta.iv } : {}) });
  }
  return out;
}

/**
 * Install late-join media addresses onto a freshly rehydrated document
 * (doc 16 §6).
 *
 * A joiner's snapshot contains the REGISTRATIONS — relationships pointing at
 * absent parts — but not the declared shas, because those live only in
 * intents already folded into the snapshot. `load()` therefore derives holes
 * with an empty sha, which render a skeleton but can never be fetched. This
 * fills those addresses in, turning "I can see an image belongs here" into
 * "I can go and get it".
 *
 * Only ever fills a part that is genuinely pending: an address must never
 * overwrite the metadata of a part whose bytes this replica already holds.
 */
export function applyMediaAddresses(
  doc: DocxDocument,
  addresses: MediaAddress[] | undefined,
): void {
  if (!addresses?.length) return;
  for (const { part, sha, iv } of addresses) {
    if (!sha || doc.mediaStatus(part) === "ready") continue;
    const meta = { sha, ...(iv ? { iv } : {}) };
    doc.pendingMedia.set(part, meta);
    doc.mediaMeta.set(part, meta);
  }
}

/** Lowercase hex sha256 — the address of a blob. */
export async function sha256Hex(bytes: Uint8Array): Promise<string> {
  const digest = await crypto.subtle.digest("SHA-256", bytes as BufferSource);
  let hex = "";
  for (const b of new Uint8Array(digest)) hex += b.toString(16).padStart(2, "0");
  return hex;
}

/** Injectable fetch so tests drive the relay without a socket or a server. */
export type FetchLike = (url: string, init?: { method?: string; body?: Uint8Array }) => Promise<{
  ok: boolean;
  status: number;
  arrayBuffer(): Promise<ArrayBuffer>;
}>;

export interface MediaTransportOptions {
  /** Origin serving the doc-16 §3 routes, e.g. "http://localhost:1234". */
  httpBase: string;
  /** Defaults to global fetch. */
  fetchImpl?: FetchLike;
}

function mediaUrl(httpBase: string, docId: string, sha: string): string {
  return `${httpBase.replace(/\/$/, "")}/docs/${encodeURIComponent(docId)}/media/${sha}`;
}

/** PUT a blob at its own hash. 200 already-present / 201 stored are both
 * success; anything else (413 too large, 507 room quota, 400 sha mismatch)
 * is a refusal the PLACER must see before it emits any intent. */
export async function putBlob(
  opts: MediaTransportOptions,
  docId: string,
  sha: string,
  blob: Uint8Array,
): Promise<{ ok: boolean; status: number }> {
  const f = opts.fetchImpl ?? (globalThis.fetch as unknown as FetchLike);
  try {
    const res = await f(mediaUrl(opts.httpBase, docId, sha), { method: "PUT", body: blob });
    return { ok: res.status === 200 || res.status === 201, status: res.status };
  } catch {
    return { ok: false, status: 0 }; // offline / CORS / DNS — same as a refusal
  }
}

/** GET a blob. null means "the relay does not have it" (404) — the caller
 * then asks the room via `media-need`. */
export async function getBlob(
  opts: MediaTransportOptions,
  docId: string,
  sha: string,
): Promise<Uint8Array | null> {
  const f = opts.fetchImpl ?? (globalThis.fetch as unknown as FetchLike);
  try {
    const res = await f(mediaUrl(opts.httpBase, docId, sha));
    if (!res.ok) return null;
    return new Uint8Array(await res.arrayBuffer());
  } catch {
    return null;
  }
}

/** How a pending part is faring, for the placeholder UI (doc 16 §5.2 step 4). */
export type MediaState = "fetching" | "waiting" | "unavailable";

/** Crypto seam: plaintext mode supplies neither, E2EE mode supplies both. */
export interface MediaCrypto {
  /** Seal plaintext for upload; reusing `iv` reproduces a blob byte-identically. */
  seal(plaintext: Uint8Array, iv?: string): Promise<{ blob: Uint8Array; iv: string }>;
  /** Open a downloaded blob; MUST throw on a GCM tag failure. */
  open(blob: Uint8Array, iv: string): Promise<Uint8Array>;
}

export interface MediaClientCallbacks {
  /** Bytes landed (or a state changed) — repaint. */
  onChange?: () => void;
  /** Placeholder state per part, for the "unavailable" affordance. */
  onState?: (part: string, state: MediaState) => void;
  /** Ask the room for a sha (the connection sends the WS frame). */
  need: (sha: string) => void;
  /** Volunteer holdings (the connection sends the WS frame). */
  have: (shas: string[]) => void;
}

/**
 * Per-replica media duties: eager fetch for holes, re-supply for holdings.
 *
 * Deliberately owns no transport and no keys of its own — the connection
 * hands it a doc getter, a crypto seam (absent in plaintext mode), and the
 * two WS senders. That keeps ONE implementation of the state machine behind
 * both the plaintext and encrypted connections, which had drifted to "fully
 * built" and "entirely absent" respectively.
 */
export class MediaClient {
  private inFlight = new Set<string>();
  private states = new Map<string, MediaState>();
  /** Shas we asked the room for, awaiting a media-ready. */
  private waitingFor = new Set<string>();

  constructor(
    private opts: MediaTransportOptions,
    private getDoc: () => DocxDocument | null,
    private cb: MediaClientCallbacks,
    private crypto?: MediaCrypto,
  ) {}

  /** Current placeholder state of a pending part (undefined once ready). */
  stateOf(part: string): MediaState | undefined {
    return this.states.get(part);
  }

  /**
   * PLACER (doc 16 §5.1). Seal (E2EE) → address → upload. Returns the intent
   * fields on success, or null when the relay refused — in which case the
   * caller MUST NOT emit an insertImage, or the room gets a reservation
   * pointing at a blob that does not exist.
   */
  async upload(docId: string, plaintext: Uint8Array): Promise<{ blobSha: string; bytesLen: number; iv?: string } | null> {
    const sealed = this.crypto ? await this.crypto.seal(plaintext) : null;
    const blob = sealed ? sealed.blob : plaintext;
    const blobSha = await sha256Hex(blob);
    const res = await putBlob(this.opts, docId, blobSha, blob);
    if (!res.ok) return null;
    // "Bytes install into the placer's own doc immediately" (doc 16 §5.1).
    // Keeping the plaintext here is what makes that true no matter WHEN the
    // reservation applies — in an encrypted room the intent is sealed and
    // applied on an async queue, so the part often doesn't exist yet at the
    // moment upload() returns, and the placer would otherwise download its
    // own image back from the relay just to look at it.
    this.ownBlobs.set(blobSha, plaintext);
    return { blobSha, bytesLen: blob.length, iv: sealed?.iv };
  }

  /** Plaintext of blobs THIS client uploaded, so its own images never take a
   * network round trip. Bounded: an entry is dropped once installed. */
  private ownBlobs = new Map<string, Uint8Array>();

  /**
   * RECEIVER (doc 16 §5.2). Eager-fetch every pending part the doc knows
   * about. Idempotent and safe to call after every applied broadcast: parts
   * already ready or already in flight are skipped.
   */
  async fetchPending(docId: string): Promise<void> {
    const doc = this.getDoc();
    if (!doc) return;
    for (const [part, meta] of [...doc.pendingMedia]) {
      // An empty sha means the hole was DERIVED from the package (a rel to a
      // missing part) rather than from an intent, so no address was ever
      // committed for it and there is nothing to ask the relay for. It still
      // renders a skeleton; see the late-join note in the arc report.
      if (!meta.sha || this.inFlight.has(part)) continue;
      this.inFlight.add(part);
      void this.fetchOne(docId, part, meta).finally(() => this.inFlight.delete(part));
    }
  }

  private setState(part: string, state: MediaState): void {
    this.states.set(part, state);
    // Mirror onto the doc so the renderer's skeleton can say which of
    // "still coming" and "nobody online has it" it is showing, without the
    // renderer knowing anything about transports.
    this.getDoc()?.mediaTransferState.set(part, state);
    this.cb.onState?.(part, state);
  }

  private async fetchOne(docId: string, part: string, meta: { sha: string; iv?: string }): Promise<void> {
    const doc = this.getDoc();
    if (!doc || doc.mediaStatus(part) === "ready") return;
    // Our own upload: install straight from memory (no relay, no decrypt —
    // we are holding the plaintext that produced this address).
    const own = this.ownBlobs.get(meta.sha);
    if (own) {
      doc.installMedia(part, own);
      this.states.delete(part);
      doc.mediaTransferState.delete(part);
      this.ownBlobs.delete(meta.sha);
      this.cb.onChange?.();
      return;
    }
    this.setState(part, "fetching");
    const blob = await getBlob(this.opts, docId, meta.sha);
    if (!blob) {
      // Relay miss: register as a waiter and let the re-supply flow run.
      this.setState(part, "waiting");
      this.waitingFor.add(meta.sha);
      this.cb.need(meta.sha);
      return;
    }
    await this.installVerified(part, meta, blob);
  }

  /**
   * Verify then install. The sha check is against the value COMMITTED IN THE
   * INTENT, never against anything the deliverer said — defense in depth
   * behind the relay's own check (doc 16 §5.2 step 3), and the only thing
   * standing between a compromised relay and arbitrary bytes in the document.
   */
  private async installVerified(part: string, meta: { sha: string; iv?: string }, blob: Uint8Array): Promise<boolean> {
    if ((await sha256Hex(blob)) !== meta.sha) {
      this.setState(part, "waiting"); // wrong bytes: keep the skeleton, keep waiting
      return false;
    }
    let bytes = blob;
    if (this.crypto) {
      if (!meta.iv) return false; // encrypted room, no IV recorded: cannot open
      try {
        bytes = await this.crypto.open(blob, meta.iv);
      } catch {
        this.setState(part, "waiting"); // GCM tag failure — reject, keep skeleton
        return false;
      }
    }
    const doc = this.getDoc();
    if (!doc) return false;
    doc.installMedia(part, bytes);
    this.states.delete(part);
    doc.mediaTransferState.delete(part);
    this.waitingFor.delete(meta.sha);
    this.cb.onChange?.();
    return true;
  }

  /** Server says the blob is fetchable now — retry the parts waiting on it. */
  async onReady(docId: string, sha: string): Promise<void> {
    this.waitingFor.delete(sha);
    const doc = this.getDoc();
    if (!doc) return;
    for (const [part, meta] of [...doc.pendingMedia]) {
      if (meta.sha !== sha || this.inFlight.has(part)) continue;
      this.inFlight.add(part);
      void this.fetchOne(docId, part, meta).finally(() => this.inFlight.delete(part));
    }
  }

  /** No holder is online. The registration stays; a later media-ready (or a
   * holder rejoining, §5.4) recovers it with no action from the user. */
  onUnavailable(sha: string): void {
    const doc = this.getDoc();
    if (!doc) return;
    for (const [part, meta] of doc.pendingMedia) {
      if (meta.sha === sha) this.setState(part, "unavailable");
    }
  }

  /** Shas this replica can serve: parts whose bytes are PRESENT. */
  heldShas(): string[] {
    const doc = this.getDoc();
    if (!doc) return [];
    const out: string[] = [];
    for (const [part, meta] of doc.mediaMeta) {
      if (doc.mediaStatus(part) === "ready") out.push(meta.sha);
    }
    return out;
  }

  /** §5.4: intersect the room's outstanding needs with local holdings and
   * volunteer — this is what makes evicted media reappear when a holder
   * comes back, with no polling anywhere. */
  volunteer(mediaNeeded: string[] | undefined): void {
    if (!mediaNeeded?.length) return;
    const held = new Set(this.heldShas());
    const intersection = mediaNeeded.filter((sha) => held.has(sha));
    if (intersection.length) this.cb.have(intersection);
  }

  /** Someone in the room needs a sha: answer if we hold it (§5.3). */
  answerRequest(sha: string): void {
    if (this.heldShas().includes(sha)) this.cb.have([sha]);
  }

  /**
   * HOLDER DUTY (doc 16 §5.3): chosen to re-supply — upload our copy.
   *
   * In an encrypted room this RE-SEALS with the IV recorded on the part, not
   * a fresh one: same key + same IV + same plaintext reproduces the exact
   * ciphertext, which is the only thing that still hashes to the address the
   * intent committed to. A fresh IV would produce a perfectly valid blob at
   * the WRONG address, and the upload would be rejected — do not "fix" it
   * that way (doc 16 §5.3 / doc 13 §4 both carry this warning).
   *
   * The local sha assertion before PUT is deliberate: failing it means THIS
   * replica's pixels are corrupt, and uploading them would just burn the
   * re-supply rotation on bytes the relay will reject anyway.
   */
  async resupply(docId: string, sha: string): Promise<boolean> {
    const doc = this.getDoc();
    if (!doc) return false;
    for (const [part, meta] of doc.mediaMeta) {
      if (meta.sha !== sha || doc.mediaStatus(part) !== "ready") continue;
      const pixels = doc.media(part);
      if (!pixels) continue;
      let blob = pixels;
      if (this.crypto) {
        if (!meta.iv) return false;
        blob = (await this.crypto.seal(pixels, meta.iv)).blob;
      }
      if ((await sha256Hex(blob)) !== sha) return false; // local corruption — never upload
      const res = await putBlob(this.opts, docId, sha, blob);
      return res.ok;
    }
    return false;
  }
}
