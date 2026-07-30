import { ReactNode } from 'react';
import { X as XmlElement, D as DocxDocument, C as ClientMessage, S as ServerMessage, P as PresencePosition, R as RosterEntry, a as ParticipantProfile, I as Intent, b as IdSidecar, L as LineageHead, T as ToolbarFeature, c as ToolbarMode, d as DocxViewApi, W as WriteStatus, U as UndoOutcome } from './index-PUk3MbmI.js';
export { e as bytesToB64, f as deriveEpochKeys, g as docKeyFromFragment, m as mintDocKey, s as sealCheckpoint, h as stretchShareCode } from './index-PUk3MbmI.js';

/**
 * How much of the document an applied intent disturbed — the input to
 * dirty-scoped reconciliation (perf B9/B10). Reconciling an edit used to cost
 * O(document) on every path (full model reparse + full id walk), so per-op
 * cost grew with the document; the scope lets the caller pay for THE EDIT.
 *
 * `doc` is the conservative default and means exactly today's behavior: full
 * refresh() + assignFromRoots(). Only intents whose blast radius is provably
 * one paragraph (or the two sides of a split) report a narrower scope; every
 * structural, document-level, or exotic intent keeps the document scope.
 */
type Scope = {
    kind: "doc";
}
/** Only these paragraphs' contents changed. */
 | {
    kind: "block";
    blocks: XmlElement[];
}
/** `before` was split; `after` is the new sibling that follows it. */
 | {
    kind: "split";
    before: XmlElement;
    after: XmlElement;
};

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
interface MediaAddress {
    part: string;
    sha: string;
    iv?: string;
}
/** The late-join address map for a document: every registered part whose sha
 * is actually KNOWN. Parse-derived holes (sha === "") are omitted — an empty
 * address is not an address, and passing one on would only teach the joiner
 * the same nothing. Mirrors the hub's filter for plaintext welcomes. */
declare function mediaAddressesOf(doc: DocxDocument): MediaAddress[];
/** Injectable fetch so tests drive the relay without a socket or a server. */
type FetchLike = (url: string, init?: {
    method?: string;
    body?: Uint8Array;
}) => Promise<{
    ok: boolean;
    status: number;
    arrayBuffer(): Promise<ArrayBuffer>;
}>;
interface MediaTransportOptions {
    /** Origin serving the doc-16 §3 routes, e.g. "http://localhost:1234". */
    httpBase: string;
    /** Defaults to global fetch. */
    fetchImpl?: FetchLike;
}
/** How a pending part is faring, for the placeholder UI (doc 16 §5.2 step 4). */
type MediaState = "fetching" | "waiting" | "unavailable";
/** Crypto seam: plaintext mode supplies neither, E2EE mode supplies both. */
interface MediaCrypto {
    /** Seal plaintext for upload; reusing `iv` reproduces a blob byte-identically. */
    seal(plaintext: Uint8Array, iv?: string): Promise<{
        blob: Uint8Array;
        iv: string;
    }>;
    /** Open a downloaded blob; MUST throw on a GCM tag failure. */
    open(blob: Uint8Array, iv: string): Promise<Uint8Array>;
}
interface MediaClientCallbacks {
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
declare class MediaClient {
    private opts;
    private getDoc;
    private cb;
    private crypto?;
    private inFlight;
    private states;
    /** Shas we asked the room for, awaiting a media-ready. */
    private waitingFor;
    constructor(opts: MediaTransportOptions, getDoc: () => DocxDocument | null, cb: MediaClientCallbacks, crypto?: MediaCrypto | undefined);
    /** Current placeholder state of a pending part (undefined once ready). */
    stateOf(part: string): MediaState | undefined;
    /**
     * PLACER (doc 16 §5.1). Seal (E2EE) → address → upload. Returns the intent
     * fields on success, or null when the relay refused — in which case the
     * caller MUST NOT emit an insertImage, or the room gets a reservation
     * pointing at a blob that does not exist.
     */
    upload(docId: string, plaintext: Uint8Array): Promise<{
        blobSha: string;
        bytesLen: number;
        iv?: string;
    } | null>;
    /** Plaintext of blobs THIS client uploaded, so its own images never take a
     * network round trip. Bounded: an entry is dropped once installed. */
    private ownBlobs;
    /**
     * RECEIVER (doc 16 §5.2). Eager-fetch every pending part the doc knows
     * about. Idempotent and safe to call after every applied broadcast: parts
     * already ready or already in flight are skipped.
     */
    fetchPending(docId: string): Promise<void>;
    private setState;
    private fetchOne;
    /**
     * Verify then install. The sha check is against the value COMMITTED IN THE
     * INTENT, never against anything the deliverer said — defense in depth
     * behind the relay's own check (doc 16 §5.2 step 3), and the only thing
     * standing between a compromised relay and arbitrary bytes in the document.
     */
    private installVerified;
    /** Server says the blob is fetchable now — retry the parts waiting on it. */
    onReady(docId: string, sha: string): Promise<void>;
    /** No holder is online. The registration stays; a later media-ready (or a
     * holder rejoining, §5.4) recovers it with no action from the user. */
    onUnavailable(sha: string): void;
    /** Shas this replica can serve: parts whose bytes are PRESENT. */
    heldShas(): string[];
    /** §5.4: intersect the room's outstanding needs with local holdings and
     * volunteer — this is what makes evicted media reappear when a holder
     * comes back, with no polling anywhere. */
    volunteer(mediaNeeded: string[] | undefined): void;
    /** Someone in the room needs a sha: answer if we hold it (§5.3). */
    answerRequest(sha: string): void;
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
    resupply(docId: string, sha: string): Promise<boolean>;
}

/**
 * A transport the connection drives: send a client message, and register a
 * handler for inbound server messages. A WebSocket adapter (or a fake, in
 * tests) implements it. The connection contains no socket code — the client
 * counterpart to the transport-free server hub (plan doc 07).
 */
interface ClientTransport {
    send(msg: ClientMessage): void;
    onMessage(cb: (msg: ServerMessage) => void): void;
}
interface ConnectionCallbacks {
    /** The document changed (welcome loaded, a broadcast applied). */
    onChange?: () => void;
    /** A remote participant's cursor/selection moved (or cleared: null). */
    onPresence?: (participant: string, position: PresencePosition | null) => void;
    /** The server refused the connection (e.g. version mismatch). */
    onRefused?: (reason: string) => void;
    /**
     * Resume landed in a DIFFERENT epoch than the bundle's (doc 12 §5 case 2):
     * someone re-seeded while this client was away. The connection has taken
     * the server's state and did NOT replay the bundle's pending intents
     * (they belong to the old epoch — replaying them would be a silent
     * cross-epoch merge, which the fork rule forbids). The consumer decides
     * what to do with the old bundle: doc-15 lineage fast-forward/draft, or
     * pre-lineage, "your offline copy is saved as a draft".
     */
    onEpochChange?: (storedGenesisId: string, currentGenesisId: string) => void;
    /** The session roster changed (join/leave/rename) — full snapshot. */
    onRoster?: (roster: RosterEntry[]) => void;
    /**
     * The connection detected that the local OPTIMISTIC replica drifted from
     * the canonical document at quiescence and rebuilt it in place (encrypted
     * mode: the mirror is local ground truth, so the check costs no network).
     * The doc object was swapped and docEpoch bumped — the renderer remounts
     * exactly like any true-conflict reload. Informational: the heal already
     * happened; consumers surface/telemeter it (the B6a typist-drift class).
     */
    onSelfHeal?: (info: {
        seq: number;
        liveHash: string;
        canonicalHash: string;
    }) => void;
    /**
     * A submit was DROPPED instead of sent, and the caller needs to know.
     *
     * Today the only reason is "not-ready": a submit arriving before the
     * welcome has nothing to apply against, no confirmed seq to base itself on,
     * and (encrypted) no key to seal with, so it cannot be honoured or queued.
     * It used to return silently — no error, no callback, no counter, no
     * replay — which made lost edits indistinguishable from edits that were
     * never made. B13 measured 403 submits, 172 envelopes on the wire and 172
     * in the document, with a completely clean server log; a channel that
     * consumes intents without a trace is what that requires.
     *
     * NOT `onRefused`: that one means the SERVER refused the connection and the
     * demo renders a "please refresh" screen on it. This is non-fatal and
     * per-intent — the connection stays perfectly usable.
     */
    onSubmitDropped?: (reason: "not-ready") => void;
    /**
     * Something threw on an async path that has nowhere to return an error —
     * a seal, a transport send, a persistence write. These used to die in a
     * serial chain's `.catch(() => {})`, which is how B13's lost edits and a
     * silently-failing durable copy both stayed invisible.
     *
     * DELIBERATELY SEPARATE from onRefused ("the server refused the
     * connection", which tears the session down) and onSubmitDropped ("an edit
     * was lost before it could be sent"). Each of the three now means exactly
     * one thing; conflating failure channels is precisely how the bugs this
     * arc is cleaning up managed to hide.
     *
     * `where` is a fixed code-site label, never interpolated with session data.
     */
    onError?: (info: {
        where: string;
        error: unknown;
    }) => void;
    /**
     * Resume landed in a different epoch that is a STRICT DESCENDANT of this
     * client's stored head (doc 15 §1 fast-forward): the seed's lineage
     * contains our (genesisId, hash) at ≥ our seq, so adopting the server's
     * state loses nothing — no draft, no banner noise. The consumer should
     * still bank the superseded bundle recoverable-not-gone (version ring /
     * a superseded slot) per doc 15's fabricated-lineage mitigation.
     */
    onFastForward?: (fromGenesisId: string, toGenesisId: string) => void;
    /**
     * The SERVER is going to end this session, and this is the grace period
     * before it does (server lifecycle arc). Render a countdown from `inMs`.
     *
     *  - `idle`     nobody has done anything for the idle window. Cancellable:
     *               any qualifying activity — an accepted edit, an admin
     *               action, a media transfer, someone joining, but NEVER
     *               presence — resets the clock and fires
     *               {@link onSessionWarningCleared}. Typing genuinely saves the
     *               session, so the countdown is worth showing prominently.
     *  - `lifetime` the room hit its absolute age cap. NOT cancellable by
     *               anything: the deadline was fixed when the room was created,
     *               which is the point (an activity-based clock is exactly what
     *               a keepalive script defeats). The document is not lost — the
     *               local bundle re-seeds it into a new epoch.
     *
     * `inMs` is the measured remainder at send time, not a configured constant,
     * so counting down from it lands on the real ending.
     *
     * NOT `onRefused`: nothing has been refused yet and the session is fully
     * usable during the grace. The kick that eventually follows arrives as
     * `onRefused("idle-timeout")` / `onRefused("session-expired")`.
     */
    onSessionWarning?: (info: {
        reason: "idle" | "lifetime";
        inMs: number;
    }) => void;
    /**
     * A previously announced deadline is no longer approaching — take the
     * countdown down. Only ever `idle`, because the lifetime cap cannot be
     * cancelled; the narrow type means a consumer cannot write a branch the
     * server is incapable of producing.
     */
    onSessionWarningCleared?: (info: {
        reason: "idle";
    }) => void;
    /**
     * ONE OF THIS CLIENT'S OWN INTENTS WAS REJECTED by canonical validation —
     * it will never apply, anywhere, and the edit is gone.
     *
     * This channel exists because of a bug that cost a user their photos. The
     * validator carried a stale image-size bound, so every large image was
     * rejected as "bad size" — and the failure was completely invisible. The
     * reason is worth stating exactly, because it inverts the usual assumption:
     * a rejection is sequenced, so EVERY REPLICA AGREES the intent is dead.
     * Nothing diverges, nothing throws, no upload fails. Our entire invariant
     * arsenal is built to detect disagreement, and this bug's signature was
     * agreement. The image simply never appeared.
     *
     * A wrong bound can happen again; next time somebody should SEE it. Fired
     * only for intents THIS client authored — another participant's rejection is
     * their business and not actionable here.
     *
     * `reason` is the validator's fixed vocabulary (`insertImage: bad size`,
     * `invalid base`, …) — never session content, so it is safe to surface.
     */
    onIntentRejected?: (info: {
        reason: string;
        clientSeq: number;
    }) => void;
    /** Media re-supply control (doc 16 §5): the room asks who holds `sha`;
     * a holder answers via `mediaHave` and uploads over HTTP when chosen. */
    onMediaRequest?: (sha: string) => void;
    /** A needed blob became fetchable / definitively unavailable. */
    onMediaReady?: (sha: string) => void;
    onMediaUnavailable?: (sha: string) => void;
    /** A pending part's placeholder state changed (fetching/waiting/
     * unavailable) — the renderer shows the "unavailable" affordance off this. */
    onMediaState?: (part: string, state: MediaState) => void;
}
/**
 * Client-side session: joins a document, applies canonical broadcasts to a
 * ClientReplica (optimistic apply + reconciliation), and relays presence.
 * The editor layer submits local intents and reads `replica.doc`.
 */
declare class CollabConnection {
    private transport;
    private clientId;
    private cb;
    private replica;
    private clientSeq;
    /** Bumps only when reconciliation RELOADED the document (a true conflict) —
     * the renderer re-mounts on this, but updates in place otherwise (no flash
     * for the common non-conflicting edits). */
    docEpoch: number;
    /** Per-client id allocator for carried node ids (split/format-range/insert):
     * a client-specific base block keeps concurrently-allocated ids disjoint
     * across clients so they never collide. */
    private idCounter;
    private idBase;
    /** The epoch id of the session this connection joined (from the welcome).
     * The resume layer compares it against the bundle's stored genesisId:
     * same ⇒ seamless rejoin; different ⇒ someone re-seeded while away
     * (doc 12 §5 case 2 — take server state, keep the local copy; the doc-15
     * lineage check decides fast-forward vs draft). Null until welcomed. */
    genesisId: string | null;
    /** Session encryption mode from the welcome (doc 13 §6). The E2EE layer
     * hard-refuses a value contradicting the link's `#k` fragment. */
    mode: "plaintext" | "encrypted" | null;
    /**
     * The relay's per-blob upload limit in bytes, from the welcome — so a file
     * can be checked BEFORE it is uploaded rather than after a 413.
     *
     * `null` means the server did not publish one (an older host). Callers must
     * treat that as "skip the pre-check", NOT as "no limit" and not as a
     * default they invent: the server still enforces its real limit either way,
     * and a client guessing a larger number would promise the user an upload
     * that cannot succeed.
     */
    mediaMaxBlobBytes: number | null;
    /** Out-of-band media transfer (doc 16 §5). Null until `httpBase` is
     * supplied — without a relay origin a client simply has no media duties
     * (every existing caller keeps working unchanged). */
    media: MediaClient | null;
    /** The room this connection joined; media URLs are per-doc. */
    private docId;
    constructor(transport: ClientTransport, clientId: string, cb?: ConnectionCallbacks, 
    /** Origin serving the doc-16 §3 media routes; enables media transfer. */
    mediaOpts?: MediaTransportOptions);
    /** Upload bytes and get the intent's media fields, or null if the relay
     * refused — the caller must not emit an insertImage in that case. */
    uploadMedia(plaintext: Uint8Array): Promise<{
        blobSha: string;
        bytesLen: number;
        iv?: string;
    } | null>;
    /** Replace the callbacks (used by bindEditor to attach after construction). */
    setCallbacks(cb: ConnectionCallbacks): void;
    /** Allocate `n` fresh carried node ids in this client's disjoint block. */
    allocIds(n: number): number[];
    /**
     * Join a document. The server replies with a welcome (snapshot + sidecar +
     * tail). The hello carries this connection's clientId — the hub binds it to
     * the socket and refuses submits under any other id (doc 11 decision 8).
     * `takeover: true` claims the identity from an existing live connection
     * (the doc-12 §7 "use here instead" path for a second same-profile tab);
     * without it, a duplicate join is refused `already-open`.
     */
    join(docId: string, token?: string, opts?: {
        takeover?: boolean;
        profile?: ParticipantProfile;
        codeProof?: string;
        ownerToken?: string;
    }): void;
    /** Owner admin op (doc 14 §2.5): honored only if this connection proved
     * the owner token at hello — otherwise the server refuses `not-owner`. */
    admin(action: {
        op: "kick";
        clientId: string;
    } | {
        op: "readOnly";
        on: boolean;
    } | {
        op: "setRole";
        clientId: string;
        role: "editor" | "viewer";
    }): void;
    /** The latest roster snapshot (doc 14 §2); empty until the first fan-out. */
    roster: RosterEntry[];
    /**
     * Attribution layer 1 (doc 14 §3): the canonical log IS the attribution
     * record — every applied entry carries its author's bound clientId. This
     * is the bounded client-side derivation an activity panel renders
     * (clientId → roster name/color). Newest last, capped, derived state:
     * never persisted, dies with the connection (zero custody).
     */
    activity: {
        seq: number;
        clientId: string;
        kind: string;
    }[];
    private static ACTIVITY_CAP;
    private recordActivity;
    /** Ask the room for a blob this replica needs (doc 16 §5.2). */
    mediaNeed(sha: string): void;
    /** Volunteer holdings (reply to a request, or after a welcome whose
     * mediaNeeded intersects local media — §5.4). */
    mediaHave(shas: string[]): void;
    /** Shas this replica can re-supply: metadata of parts whose bytes are
     * PRESENT (pending parts are exactly what we can't serve). */
    heldMediaShas(): string[];
    /** Rename/recolor this participant mid-session. The server sanitizes and
     * fans out the updated roster; the local copy updates on that echo (no
     * optimistic roster — a 1-RTT lag on your own rename is imperceptible and
     * keeps one code path). */
    setProfile(profile: ParticipantProfile): void;
    /** Bundle being resumed from (set by resume(), consumed at welcome). */
    private resuming;
    /**
     * Rejoin a document from a persisted bundle (doc 12 §5). Restores the
     * clientSeq watermark FIRST — a fresh counter would reuse already-
     * sequenced (clientId, clientSeq) keys and the server would dedup this
     * client's NEW edits as re-sends (silent edit loss). The welcome decides
     * the case: same epoch ⇒ replay pending (below); different ⇒
     * onEpochChange, pending withheld.
     */
    resume(bundle: DocBundle, token?: string, opts?: {
        profile?: ParticipantProfile;
        codeProof?: string;
        ownerToken?: string;
    }): void;
    /**
     * The current durable state as a doc-12 bundle, or null before welcome.
     * `savedAt` is stamped by the persister at write time (clock injection —
     * this module never reads Date.now itself).
     */
    exportBundle(docId: string): DocBundle | null;
    exportBundleAsync(docId: string): Promise<DocBundle | null>;
    /** The live document (null until welcome). The editor renders this. */
    get doc(): DocxDocument | null;
    /** Accumulates doc changes across replica rebuilds (welcome), so the getter
     * below stays monotonic per connection. */
    private docVersionBase;
    /**
     * Counts changes TO THE RENDERED DOCUMENT — the repaint signal, as opposed
     * to onChange (which also fires for pure bookkeeping such as a tracked own
     * echo). The editor-driven typing path mutates and paints the doc itself,
     * so its submit + echo leave this untouched; remote applies, optimistic
     * canonical applies (toolbar ops), reloads, and media installs advance it.
     * The react layer repaints on THIS, not on onChange — repainting on
     * onChange queued a whole-document relayout per keystroke, catastrophic
     * past the background-layout page threshold.
     */
    get docVersion(): number;
    /** docVersionBase at the last takeRenderScope — a base bump between takes
     * (welcome doc replacement, media install) has no narrow scope. */
    private takenDocVersionBase;
    /**
     * Drain the dirty scope behind the docVersion movement since the last take
     * — what the repaint answering `docVersion` must relayout. Replica applies
     * carry their per-intent scope; connection-level bumps (a replaced doc
     * object, media pixels landing) report document scope. Null means NOTHING
     * was recorded since the last take — the caller already painted everything
     * and may skip the repaint.
     */
    takeRenderScope(): Scope | null;
    get ready(): boolean;
    /** Un-confirmed local intents in flight — the drained-replay discipline
     * (doc 15 §2 / rebase.ts) polls this to submit one intent at a time. */
    get pendingCount(): number;
    /**
     * Submit a local edit. The caller supplies the intent minus its wire
     * bookkeeping (clientId/clientSeq/base) — the connection fills those from
     * the current confirmed seq and applies it optimistically before sending.
     */
    submit(intent: Omit<Intent, "clientId" | "clientSeq" | "base">): void;
    /** Submit an intent whose mutation the caller ALREADY performed on this
     * connection's live doc (the editor-driven path: DocxEditor applies the
     * command to `conn.doc` and then emits the intent). Skips the optimistic
     * re-apply — applying twice doubled every keystroke — but tracks pending and
     * sends identically, so echoes and reconciliation work unchanged. */
    submitPreApplied(intent: Omit<Intent, "clientId" | "clientSeq" | "base">): void;
    /** Submits dropped because the connection was not ready (see
     * onSubmitDropped). Exposed like `selfHeals` so a harness or a UI can read
     * it: when a run loses intents, this number says how many died HERE rather
     * than anywhere else on the path. */
    droppedPreReady: number;
    private submitFull;
    /** Broadcast this client's cursor/selection (ephemeral). Selection ranges
     * are clamped on the way OUT too, so an over-long selection costs the room
     * a bounded payload rather than every receiver a discard. */
    setPresence(position: PresencePosition | null): void;
    /**
     * Surface any of OUR OWN intents that canonical validation rejected.
     *
     * A rejection is sequenced and agreed by every replica, so it produces no
     * divergence to notice — which is exactly how a stale image-size bound
     * silently ate a user's photos. See ConnectionCallbacks.onIntentRejected.
     */
    private reportRejections;
    private onServer;
    private rlRedriveTimer;
    private rlRedriveBackoffMs;
}

/**
 * The client-side document bundle (plan doc 12 §4) — under zero custody this
 * is THE durable copy of a document anywhere: the server deletes everything
 * at session end, so each participating browser persists its own bundle and
 * any holder can re-seed the same link later.
 *
 * Contents are the replica's CONFIRMED state plus the pending queue:
 * `confirmedBytes` is a complete docx at `confirmedSeq`; `confirmedSidecar`
 * rides with it always (round-4 F10 — bytes alone cannot reproduce the id
 * table across split-created ids); `pending` is idempotency-keyed and safe
 * to replay after a crash (the server dedups by (clientId, clientSeq)).
 */
interface DocBundle {
    docId: string;
    /** Epoch of the session this state came from (doc 12 §5): on rejoin,
     * same ⇒ seamless resume; different ⇒ someone re-seeded while away. */
    genesisId: string;
    confirmedSeq: number;
    confirmedBytes: Uint8Array;
    confirmedSidecar: IdSidecar;
    /** Local intents not yet confirmed when the bundle was written — replayed
     * verbatim on resume; idempotent by (clientId, clientSeq). */
    pending: Intent[];
    /**
     * The clientSeq watermark at write time. Resume MUST restore this before
     * submitting anything new: a fresh connection restarting from 1 would
     * reuse already-sequenced (clientId, clientSeq) keys and the server would
     * dedup its NEW edits as re-sends — self-inflicted silent edit loss (the
     * same failure shape doc 12 §7's single-tab rule guards against).
     */
    clientSeq: number;
    savedAt: number;
    /**
     * The lineage chain (doc 15 §1): every epoch this copy has passed
     * through, newest last — `[..., {this bundle's genesisId, seq, hash}]`.
     * Rides with re-seeds so returning holders can prove ancestry and
     * fast-forward instead of forking. Capped; ancestry checks are O(chain).
     */
    lineage: LineageHead[];
    /**
     * Local intents recorded while the session was UNREACHABLE (doc 15 §2 /
     * doc 12 §5 "offline mode is first-class"): the user kept editing after
     * the socket died. On rejoin these are the offline tail the arrival
     * ladder reconciles — replayed as ordinary intents (same epoch), as
     * suggestions (diverged, small), or parked in a draft (diverged, large).
     * Distinct from `pending` (in-flight when a LIVE session was
     * interrupted): tail entries carry NO wire bookkeeping — clientId /
     * clientSeq / base are assigned at replay time by the ordinary submit
     * path, so the stored objects are the payload half of an Intent only.
     */
    offlineTail?: Intent[];
    /**
     * The genesisId of the epoch the offline tail was recorded AGAINST. This
     * is what lets a resume decide the arrival rung without guessing: same
     * epoch ⇒ the tail is morally a large pending queue and replays silently;
     * different (or absent, for a tail written by an older build) ⇒ true
     * divergence, offer suggest/draft — never silently bulldoze.
     */
    offlineTailEpoch?: string;
    /** Per-part media metadata (doc 16 §5.3): the sha (and E2EE iv/epoch)
     * of every media part this copy knows. The docx bytes carry the PIXELS;
     * this carries the ADDRESSES — without it a resumed holder couldn't
     * answer re-supply requests or verify fetches (metadata is in-memory on
     * the doc and would otherwise die with the session). */
    mediaMeta?: [string, {
        sha: string;
        iv?: string;
        genesisId?: string;
    }][];
}
/**
 * THE KEY CONVENTION, in one place. Everything a browser stores rides the
 * bundle store under a key derived from the base docId:
 *
 *  - `<docId>`                          — the live bundle
 *  - `<docId>#version-<ts>[-<label>]`   — a frozen restore point
 *  - `<docId>#draft-<genesis>`          — an offline copy parked on divergence
 *  - `<docId>#superseded-<genesis>`     — a banked copy after a fast-forward
 *  - `local:…`                          — a document that never went live
 *
 * The builders and the parser below are the only code that knows these
 * shapes; UIs render from `parseBundleKey` instead of re-deriving them.
 */
type StoredDocKind = "live" | "version" | "draft" | "superseded" | "local" | "unknown";
interface ParsedBundleKey {
    /** The base document id (the part before any `#` suffix). */
    docId: string;
    kind: StoredDocKind;
    /** A version's human label, when the key carries one. */
    label?: string;
    /** A version's freeze time (ms), parsed from the key. */
    versionSavedAt?: number;
}
declare function versionKey(docId: string, savedAt: number, label?: string): string;
declare function draftKey(docId: string, genesisId: string): string;
declare function supersededKey(docId: string, genesisId: string): string;
/** Parse any store key. An unrecognised `#` suffix degrades to `unknown`
 * rather than throwing — a future key shape must not break today's listing. */
declare function parseBundleKey(key: string): ParsedBundleKey;
/**
 * One stored entry, as METADATA ONLY — everything a list needs to render
 * (what it is, when, how big) without the document bytes. A 500-page doc is
 * megabytes; listing ten of them must not deserialise all ten, so `list()`
 * returns these and `get(key)` fetches bytes only when one is opened.
 */
interface StoredDocSummary extends ParsedBundleKey {
    /** The full store key — what `get()`/`delete()` take. */
    key: string;
    /** When this entry was written (bundle.savedAt, ms). */
    savedAt: number;
    /** Size of the stored document bytes. */
    byteLength: number;
}
/**
 * Storage seam for bundles. The browser implementation is IndexedDB (the
 * only storage that fits multi-MB binary docs); tests and Node use the
 * in-memory implementation — the doc-12 test plan's "in-memory IndexedDB
 * stub", which keeps every bundle test deterministic and dependency-free.
 */
interface BundleStore {
    get(docId: string): Promise<DocBundle | null>;
    put(bundle: DocBundle): Promise<void>;
    delete(docId: string): Promise<void>;
    /** Enumerate everything stored, metadata only (see StoredDocSummary). */
    list(): Promise<StoredDocSummary[]>;
}
declare class InMemoryBundleStore implements BundleStore {
    private bundles;
    /** Write count — lets tests assert the throttle's coalescing precisely. */
    writes: number;
    get(docId: string): Promise<DocBundle | null>;
    put(bundle: DocBundle): Promise<void>;
    delete(docId: string): Promise<void>;
    list(): Promise<StoredDocSummary[]>;
}
/**
 * Persists a connection's state to a BundleStore on a THROTTLE — not a
 * debounce (round-4 F8): a debounce resets per event, so sustained typing
 * defers the write for the whole burst and an OS kill loses the burst; a
 * throttle bounds data loss to the throttle window (default 1s), which is
 * the actual durability guarantee doc 12 §4 states ("RPO ≈ your last bundle
 * write ≈ the throttle window").
 *
 * Wiring: call `notify()` from the connection's onChange (every local edit,
 * echo, and remote broadcast); call `flush()` from pagehide/visibilitychange
 * — best-effort (IndexedDB has no synchronous API; the throttle is the real
 * mechanism, the flush just narrows the tail).
 *
 * The scheduler (setTimeout/clearTimeout) and clock are injectable so tests
 * drive time deterministically.
 */
declare class BundlePersister {
    private conn;
    private store;
    private docId;
    private opts;
    private lastWrite;
    private trailing;
    private stopped;
    /** Writes are async (digest + store I/O) and must not overlap on one
     * store; a serial chain guarantees ordering and lets flush() await
     * every in-flight write (round-4 F8: the flush is only meaningful if it
     * durably lands the latest state before pagehide). */
    private chain;
    constructor(conn: CollabConnection, store: BundleStore, docId: string, opts?: {
        throttleMs?: number;
        now?: () => number;
        setTimer?: (fn: () => void, ms: number) => unknown;
        clearTimer?: (t: unknown) => void;
        /**
         * A persistence write FAILED. Wire this — in a zero-custody design the
         * browser's stored bundle IS the durable copy, so a swallowed quota or
         * blocked-storage error means the user's only durable copy silently
         * stops updating while the editor looks perfectly healthy. That is the
         * most dangerous silence in this codebase, which is why it is now a
         * callback instead of an empty catch.
         */
        onError?: (err: unknown) => void;
        /**
         * The current offline tail (doc 15 §2), read at write time. The
         * connection's exportBundle cannot know it (the tail is recorded by
         * the layer ABOVE the connection, across connection rebuilds), so the
         * owner of the tail injects a getter. Returning an empty tail (or
         * omitting the getter) ERASES any stored tail — which is exactly
         * right once a replay has drained it.
         */
        offlineTail?: () => {
            tail: Intent[];
            epoch?: string;
        } | null;
    });
    private get throttleMs();
    private now;
    /** State changed — write now if the window allows, else arm ONE trailing
     * write for the window's end (never more than one timer in flight; N
     * notifies inside a window coalesce into a single trailing write). */
    notify(): void;
    /** Immediate best-effort write (pagehide/visibilitychange/session-end).
     * Awaits ALL in-flight writes plus this one so the latest state is
     * durably landed before the tab goes away. */
    flush(): Promise<void>;
    /** Append a write to the serial chain (never overlapping). */
    private enqueueWrite;
    /** Detach (unmount): cancel the trailing timer; no further writes. */
    stop(): void;
    private write;
}

/**
 * Whether this client's edits can currently reach the sequencer.
 *
 *  - `live`         the round trip completed recently; submits will land.
 *  - `reconnecting` the socket is gone and a retry is scheduled. TRANSIENT and
 *                   usually invisible — a laptop lid closed for ten seconds
 *                   passes through here and back to `live` without the user
 *                   noticing.
 *  - `lost`         the retries are exhausted. Needs a human: reconnect or
 *                   reload.
 *
 * A NON-LIVE STATE IS NOT A WRITE GATE (doc 12 §5: offline editing is
 * first-class). What consumers MUST guarantee on anything but `live` is that
 * nothing typed can be LOST: edits apply locally and are recorded to the
 * durable offline tail (doc 15 §2) for replay on reconnect — the react
 * layer's capture path does exactly this. The write gate (`writesBlocked`)
 * stays reserved for server REFUSALS, where accepting a keystroke really is
 * accepting-then-losing it. An editable surface whose disconnected edits are
 * neither kept nor gated remains the silent-loss bug this type was created
 * to surface — keep them either captured or blocked, never dropped.
 */
type ConnectionState = "live" | "reconnecting" | "lost";
interface LivenessOptions {
    /** Gap between probes once the previous one was answered. */
    intervalMs?: number;
    /** How long a probe may go unanswered before the socket is declared dead. */
    timeoutMs?: number;
    setTimer?: (fn: () => void, ms: number) => unknown;
    clearTimer?: (t: unknown) => void;
}

/**
 * The canonical document hash (plan doc 05): sha256 over the
 * deterministically-serialized editable roots in fixed part order.
 * `serializeXml` is a pure function of the tree (attribute order preserved
 * from parse, fixed escaping, no environment reads), so replicas that hold
 * the same document produce the same hash on any engine — which is what
 * makes the hash a convergence VERIFIER rather than a hope.
 *
 * Doc-05 scope note: the full recipe also folds in relationship roots,
 * content-types, and declared media shas (package state is
 * divergence-capable). Those parts are immutable in the demo's enabled
 * intent surface (media is gated off; no intent edits rels/content-types
 * outside media), so the editable roots are the complete divergence surface
 * today — the fold-in lands with the media gate (doc 16).
 */
declare function docHash(doc: DocxDocument): Promise<string>;

declare class IndexedDbBundleStore implements BundleStore {
    private dbName;
    private db;
    private persistRequested;
    constructor(dbName?: string);
    private open;
    /** One request in one transaction, as a promise. */
    private tx;
    get(docId: string): Promise<DocBundle | null>;
    put(bundle: DocBundle): Promise<void>;
    delete(docId: string): Promise<void>;
    list(): Promise<StoredDocSummary[]>;
}

/**
 * React binding for wordinweb collaboration. Imported from the SEPARATE
 * `wordinweb/collab` entry so that non-collab apps never pull the collab
 * engine into their bundle (plan doc 07: unreachable beats shakeable — a
 * local-only `import { DocxView } from "wordinweb"` has no path to this file).
 */
interface UseCollabOptions {
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
    /**
     * Bundle persistence + resume (plan doc 12 §4/§5). When set, the hook:
     * resumes from a stored bundle if one exists (replaying pending intents,
     * idempotently) instead of joining cold; persists the confirmed state on a
     * ~1s throttle after every change, plus a best-effort flush on pagehide/
     * hidden; and on an epoch change (someone re-seeded while away) preserves
     * the superseded bundle as a draft (`<docId>#draft-<oldGenesis>`) before
     * adopting the server's state — the fork rule made visible, never a silent
     * merge. Browser apps pass an `IndexedDbBundleStore` (from `wordinweb`);
     * tests inject the in-memory store. Without it, behavior is unchanged
     * (join cold, keep nothing).
     */
    store?: BundleStore;
    /** Claim the identity from an existing live connection (doc 12 §7 "use
     * here instead"): set after an `already-open` refusal and remount. */
    takeover?: boolean;
    /**
     * E2EE mode (doc 13): the document master key from the link's `#k=`
     * fragment. Its PRESENCE selects the encrypted connection — mode is
     * derived from the link, never the wire; a plaintext welcome is
     * hard-refused (`mode-downgrade`). */
    docKey?: string;
    /** Share code (doc 13 §7) when the doc has one — stretched client-side
     * and mixed into key derivation + sent as the hello proof. */
    shareCode?: string;
    /** Owner capability token (doc 14 §2.5): held by the seeder (returned by
     * go-live, kept in their bundle), never in the shared link. Presence of
     * a valid token unlocks the admin controls. */
    ownerToken?: string;
    /**
     * HTTP origin of the media relay (plan doc 16 §3), e.g.
     * "http://localhost:1234" — the same server the WebSocket points at.
     * Image bytes travel over these routes and NEVER over the sequencer.
     * Without it the room has no media duties and the image toolbar stays
     * inert, so an app that doesn't want media simply omits it.
     */
    httpBase?: string;
    /** Display profile sent at join (doc 14 §2) — self-asserted; persist it in
     * localStorage next to the clientId so identity is stable per browser. */
    profile?: ParticipantProfile;
    /**
     * Heartbeat + reconnect tuning. Defaults suit a real network (15s probe,
     * 10s answer deadline, 6 retries); tests inject short values and a fake
     * clock. `maxRetries` is how many backoff attempts run before the session
     * gives up and reports `lost` — the point at which a human has to act.
     */
    liveness?: LivenessOptions & {
        maxRetries?: number;
    };
    /**
     * Bound on the offline tail (doc 15 §4.3) — how many intents may be
     * recorded while disconnected before the editor gates read-only with an
     * explicit "offline limit" state. Defaults to OFFLINE_TAIL_CAP; injectable
     * so tests can reach the cap without recording thousands of edits.
     */
    offlineTailCap?: number;
}
interface CollabSession {
    /** The live document to render (null until the welcome arrives). */
    doc: DocxDocument | null;
    /** Monotonically increases on every reconciled change — a cheap re-render
     * signal (the doc object may be reloaded by reconciliation). */
    version: number;
    /**
     * Increases only when the RENDERED DOCUMENT changed outside the editor —
     * remote applies, optimistic toolbar ops, reloads, self-heals, media
     * installs. The editor-driven typing path (submit + own echo) leaves it
     * untouched: the editor already mutated AND painted the doc, so repainting
     * again on `version` queued a redundant whole-document relayout per
     * keystroke — invisible on 5 pages, catastrophic past DocxView's
     * background-layout threshold (a 500-page doc relaid out asynchronously
     * behind an inert container, per keystroke). DocxView's `renderSignal`
     * rides THIS; `version` remains the observer/bookkeeping signal.
     */
    renderVersion: number;
    /**
     * Drain the union of the dirty scopes behind `renderVersion` since the last
     * take — what the repaint answering it must relayout. A remote text edit
     * reports its one paragraph, so DocxView relayouts that paragraph
     * incrementally instead of the whole document (the far-page repaint stall);
     * structural or unverifiable intents, doc reloads, and media installs
     * report `doc`, which keeps today's whole-document path. Null means nothing
     * dirty was recorded since the last take (an earlier paint covered it) and
     * the repaint may be skipped. Consumed on the repaint itself so a coalesced
     * repaint sees every batched intent's scope.
     */
    takeRenderScope: () => Scope | null;
    /** Increases only when reconciliation RELOADED the document (a true
     * conflict). The editor re-mounts on this; between reloads it updates in
     * place (no flash for the common non-conflicting edits). */
    docEpoch: number;
    /** True once joined and the snapshot is loaded. */
    ready: boolean;
    /** Submit a local edit the editor ALREADY applied to the live doc
     * (bookkeeping filled by the connection; no local re-apply). */
    submit: (intent: Omit<Intent, "clientId" | "clientSeq" | "base">) => void;
    /** Submit an operation NOT yet applied locally: it is applied optimistically
     * through the same canonical applyIntent code the server runs, so the local
     * result is byte-identical to every replica by construction. Used for
     * toolbar/API commands (insert chart, set link, page layout, ...). */
    submitOp: (intent: Omit<Intent, "clientId" | "clientSeq" | "base">) => void;
    /** Broadcast this client's cursor/selection. */
    setPresence: (pos: PresencePosition | null) => void;
    /** Allocate carried node ids (sub-range format / split / insert). */
    allocIds: (n: number) => number[];
    /** Remote participants' latest cursor/selection positions. */
    presence: Record<string, PresencePosition | null>;
    /** Set if the server refused the connection (e.g. version mismatch). */
    refused: string | null;
    /**
     * WHETHER THIS CLIENT CAN REACH THE SEQUENCER AT ALL.
     *
     * Distinct from `refused`, which means the server answered and said no —
     * this means nobody answered. A refusal ends the session; this is a
     * transport fault that usually heals by itself.
     *
     *  - `live`         the last heartbeat round trip completed.
     *  - `reconnecting` the socket dropped; a backoff retry is scheduled. Most
     *                   drops pass through here and back to `live` unnoticed,
     *                   so UI should stay quiet or subtle here.
     *  - `lost`         retries exhausted. Show something modal: the document is
     *                   safe but nothing further will sync until a human acts.
     *
     * A NON-LIVE CONNECTION DOES NOT BLOCK WRITES (doc 12 §5: offline editing
     * is first-class). "I cannot reach the server" is a different fact from
     * "the server said no": nothing typed offline is doomed — it applies
     * locally, is recorded to the offline tail, persists with the bundle, and
     * replays on reconnect (doc 15's arrival ladder). {@link offline} is the
     * state a consumer renders for it; {@link writesBlocked} stays reserved
     * for genuine refusals, where accepting keystrokes would be a lie.
     */
    connection: ConnectionState;
    /**
     * OFFLINE EDITING IS ACTIVE (doc 15 §2 / doc 12 §5): the sequencer is
     * unreachable, a document is in hand, and nothing has refused this
     * client's writes — so edits apply locally and accumulate in the durable
     * offline tail instead of being dropped.
     *
     *  - `editsHeld` how many intents the tail currently holds. Show it: the
     *    difference between "editing is paused" and "your changes are being
     *    kept (N so far)" is the whole point of the state.
     *  - `capped` the tail hit its bound (doc 15 §4.3). Recording has STOPPED
     *    and {@link writesBlocked} is raised — loudly, with this flag saying
     *    why, never by silently discarding keystrokes. The escape hatches are
     *    reconnecting or downloading a copy.
     *
     * Null while the connection is live, and null when a refusal already
     * blocks writes (a viewer's link does not become writable by going
     * through a tunnel). On reconnect the tail replays: silently as ordinary
     * edits when the session is the same epoch, else via {@link arrival}.
     */
    offline: {
        editsHeld: number;
        capped: boolean;
    } | null;
    /**
     * Retry the connection immediately, resetting the backoff and the attempt
     * budget. Wired to a "Try again" control on the `lost` UI — from `lost`
     * nothing else will ever retry, because exhausting the budget is precisely
     * the statement that automatic recovery has stopped.
     */
    reconnect: () => void;
    /**
     * The server announced it will END this session, and this is the grace
     * period before it does. `inMs` is the measured remainder at the moment the
     * warning arrived — tick a countdown from it locally rather than expecting
     * refreshed messages.
     *
     *  - `idle`     cancellable: any qualifying activity (an accepted edit, an
     *               admin action, a media transfer, a join — never presence)
     *               resets the clock and this goes back to null on its own.
     *  - `lifetime` NOT cancellable by anything; it counts down to the ending.
     *
     * Null whenever no deadline is approaching — including after ANY session
     * end (kick, refusal), because a countdown that outlived its session would
     * keep promising an ending that already happened.
     */
    sessionWarning: {
        reason: "idle" | "lifetime";
        inMs: number;
    } | null;
    /** True after an edit was refused under the owner's read-only lock
     * (doc 14 §2.5). NON-FATAL: the live view keeps working.
     *
     * LEGACY/FALLBACK now: sticky until reload or `retryWrites`, because a
     * refusal can only ever say "blocked". Prefer {@link writesBlocked}, which
     * reads the server's positive roster status when one is published and only
     * falls back to this against an older server. */
    readOnlyBlocked: boolean;
    /**
     * The server has told this client it is NOT the owner of this document.
     *
     * Holding an `ownerToken` string is not the same as it being valid — a room
     * re-seeded into a new epoch mints a fresh one, leaving the old token
     * truthy in a browser that now has no admin rights. The welcome carries no
     * owner flag, so until an admin op is attempted this is genuinely unknown;
     * a `not-owner` refusal is the authoritative answer.
     *
     * Consumers should HIDE admin controls when this is true rather than
     * disable them: an offered control that cannot work is worse than an absent
     * one, and the refusal it produces is not a session failure.
     */
    notOwner: boolean;
    /**
     * THE SERVER WILL NOT ACCEPT THIS CLIENT'S WRITES. Consumers must render the
     * editor READ-ONLY on this, never merely annotate it: an editable surface
     * over a server that refuses the writes applies every keystroke locally and
     * then loses it, which is silent data loss that looks like it worked.
     *
     * One predicate rather than a per-reason check, so a new blocking state
     * inherits the gate instead of needing to remember it.
     *
     * NOW READS THE SERVER'S ROSTER STATUS (`RosterEntry.write`), which closed
     * the seam this comment used to describe. That matters in both directions:
     * the status is present AT JOIN, so edit 1 is gated rather than applied and
     * then healed away; and it is POSITIVE, so a lift arrives on its own instead
     * of leaving the client in viewer mode until reload. Against a server that
     * publishes no status it falls back to the refusal-derived flag — never to
     * "allowed", which would restore both bugs at once.
     *
     * The status also distinguishes the three server conditions the single
     * `read-only` refusal could not (owner lock, per-client demotion, viewer
     * token), so a UI can finally say which one applies rather than inventing a
     * distinction it could not see.
     *
     * A DEAD CONNECTION IS DELIBERATELY NOT A BLOCK — that is the refusal /
     * unreachable split (doc 12 §5). This predicate means "nothing you type
     * can ever land"; a dropped socket means "nothing you type can land YET",
     * and the offline tail is what keeps the two honest: disconnected edits
     * apply locally, persist durably, and replay on reconnect, so accepting
     * them is not a lie. The one offline condition that does fold in is the
     * tail CAP ({@link offline}.capped): past it an edit genuinely cannot be
     * kept, so the gate closes loudly rather than recording silently stopping.
     * An earlier revision folded every non-live connection in here; that
     * over-reached — it was written for server refusals and extended to
     * transport faults, which made first-class offline editing impossible.
     *
     * A REFUSED SESSION FOLDS IN TOO, by the same rule. A refusal ends the
     * session — no socket, nothing to sequence against — so any document still
     * on screen (e.g. behind a "session ended" dialog) must be read-only, and
     * it inherits that from this predicate rather than from a refusal-specific
     * gate someone could forget to compose.
     */
    writesBlocked: boolean;
    /**
     * WHY the server refuses this client's writes, straight from the roster.
     *
     * UNDEFINED means the server published nothing (an older build), NOT that
     * writing is allowed — the same contract as the media limit's null. A
     * consumer must keep its copy generic in that case rather than naming a
     * cause it was never told.
     *
     * The three causes are genuinely different to the person reading the banner:
     * `owner-lock` may lift at any moment, `demoted` is about them specifically,
     * and `viewer-role` is a property of their link that nothing the owner does
     * in-session will change. Telling that last group to wait for the owner is
     * telling them to wait for something that will never happen.
     */
    writeStatus?: WriteStatus;
    /**
     * Optimistically clear the FALLBACK block and let the user try again.
     *
     * BELT AND BRACES now, not the primary escape: where the server publishes a
     * write status, a lift arrives on its own and the editor becomes writable
     * without anyone clicking anything. This remains for the older-server
     * fallback path — where the block is sticky, no lift is announced, and
     * attempting a write is the only way to discover the lock is gone — and as a
     * manual override if a status update is ever missed.
     */
    retryWrites: () => void;
    /**
     * Set when a resume landed in a different epoch than the stored bundle
     * (doc 12 §5 case 2): the session took the server's state; the old copy
     * was saved as a draft. UI copy: "restored by another participant — your
     * offline copy is saved as a draft." Null otherwise.
     */
    epochChanged: {
        from: string;
        to: string;
    } | null;
    /** Session roster (doc 14 §2): everyone who joined this session, with
     * connection state — the identity keyspace presence/attribution share. */
    roster: RosterEntry[];
    /** Rename/recolor this participant (server sanitizes + fans out). */
    setProfile: (profile: ParticipantProfile) => void;
    /** Attribution layer 1 (doc 14 §3): recent applied entries
     * {seq, clientId, kind} — join clientId to `roster` for names/colors. */
    activity: {
        seq: number;
        clientId: string;
        kind: string;
    }[];
    /** Owner admin ops (doc 14 §2.5) — no-op unless this connection proved
     * the owner token; the server refuses `not-owner` otherwise. */
    admin: (action: {
        op: "kick";
        clientId: string;
    } | {
        op: "readOnly";
        on: boolean;
    } | {
        op: "setRole";
        clientId: string;
        role: "editor" | "viewer";
    }) => void;
    /**
     * Offline reconciliation (doc 15 arrival ladder). When a rejoin lands in
     * a DIFFERENT epoch than the tail was recorded against, the recommended
     * mode is `suggest` (replay as tracked changes; ≤ the suggest threshold)
     * or `draft` (large tails). `structural` counts the tail intents that
     * have no suggestion form (splits/merges/tables/format — rebase.ts):
     * they are NOT replayed by `suggest` and survive in the banked draft
     * slot instead — say so, or their absence reads as silent loss. A
     * same-epoch rejoin never surfaces here: its tail replays silently
     * through the ordinary submit path, like a large pending queue.
     * Null when there's nothing to reconcile. `reconcile()` runs the choice.
     */
    arrival: {
        mode: "suggest" | "draft";
        tailLength: number;
        structural: number;
    } | null;
    reconcile: (mode: "suggest" | "draft") => void;
    /** How many times the connection self-healed a drifted optimistic replica
     * (encrypted mode's quiescent mirror check — the B6a class). Telemetry;
     * the heal itself is automatic and already reflected in doc/docEpoch. */
    selfHeals: number;
    /**
     * Submits the connection DROPPED because it was not ready yet (no replica
     * to apply against, no confirmed seq, no key to seal with). Should stay 0:
     * the editor is gated on `ready`, so nothing user-driven can reach that
     * guard. A non-zero value means some caller is submitting before the
     * welcome — which is a silent edit loss, and was the shape of B13's
     * 231-intent gap. Telemetry, deliberately visible rather than swallowed.
     */
    droppedPreReady: number;
    /**
     * Persistence writes that FAILED (quota exceeded, blocked storage, private
     * mode). Non-zero means this browser's stored bundle is stale — and in a
     * zero-custody design that bundle IS the durable copy, so "saved" is
     * currently a lie. Consumers should SHOW this, not just log it: the user
     * can still rescue the document with the download button, but only if they
     * are told. Sticky (never decremented): the doubt persists until reload.
     */
    persistErrors: number;
    /**
     * The connect-path storage read exceeded its deadline, so this session
     * joined COLD instead of resuming. Says nothing about whether storage is
     * full — a slow or blocked store looks identical from here, and
     * conflating the two with {@link persistErrors} told one user their
     * storage was full when it held a few megabytes.
     */
    storeSlow: boolean;
    /**
     * Submits whose seal or transport send THREW after the edit was already
     * applied optimistically — i.e. lost edits, counted at the far end of the
     * path from droppedPreReady.
     *
     * The two together are a PARTITION, which is the point: droppedPreReady is
     * "refused before it was ever applied", this is "accepted locally then lost
     * on the way out". A swarm run that loses N intents can now subtract both
     * from N and see what remains unexplained, instead of a night of guessing —
     * see B13, where exactly that arithmetic was missing.
     */
    sendFailures: number;
    /** Upload image bytes to the relay and get the address fields an
     * insertImage intent must carry (doc 16 §5.1). Null = the relay refused,
     * and the caller must not reserve anything. Always null when the app
     * supplied no `httpBase`. */
    uploadMedia: (bytes: Uint8Array) => Promise<{
        blobSha: string;
        bytesLen: number;
        iv?: string;
    } | null>;
    /**
     * Largest single upload the relay accepts, in bytes, as published in the
     * welcome — so an oversized file can be refused locally instead of after a
     * full seal-hash-upload round trip.
     *
     * NULL MEANS "NO PUBLISHED LIMIT", and the only correct response to it is to
     * skip the check: a server that publishes nothing still enforces its real
     * limit, so a client that invents a default either blocks uploads that
     * would have succeeded or promises a size that will be refused.
     */
    mediaMaxBlobBytes: number | null;
    /**
     * Reverse this participant's last SEQUENCED action (plan doc 03 Phase 8).
     * The inverse is computed locally and submitted as an ORDINARY intent, so
     * it converges, broadcasts, and — if the target has since been deleted or
     * changed — is rejected cleanly on every replica alike.
     *
     * The outcome is reported rather than thrown so the UI can say the right
     * thing: `cannot-undo` means the last action has no inverse yet (a
     * different message from having nothing to undo), and `unavailable` means
     * this connection has no collaborative undo at all — today the plaintext
     * connection, whose authority lives on the server.
     */
    undoLast: () => UndoOutcome;
}
declare function useCollab(opts: UseCollabOptions): CollabSession;
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
/** Toolbar groups OFF by default in the collab editor. Uploads (image/icon/
 * screenshot/3D/media/object) are excluded per the demo threat model (no
 * object upload → no stored-payload surface); the rest are commands not yet
 * routed through intents — showing them would make local-only edits that
 * silently diverge from other participants. Flip any of these on explicitly
 * via `toolbarFeatures` once wired. */
declare const COLLAB_TOOLBAR_DEFAULTS: Partial<Record<ToolbarFeature, boolean>>;
declare function CollabEditor(opts: UseCollabOptions & {
    editable?: boolean;
    /** Render the Word-style ribbon toolbar above the page (default true). */
    toolbar?: boolean;
    /** Ribbon mode: "simple" (Home only) or "advanced" (default). */
    toolbarMode?: ToolbarMode;
    /** Per-group overrides merged over the collab-safe defaults. */
    toolbarFeatures?: Partial<Record<ToolbarFeature, boolean>>;
    /** Observe the imperative document API (find/replace, inserts, ...). */
    onReady?: (api: DocxViewApi) => void;
    /** Observe the live CollabSession (roster, activity, epochChanged,
     * doc-for-download) — how an app shell renders chips/banners/buttons
     * around the editor without re-implementing its composition. */
    onSession?: (session: CollabSession) => void;
    /** Custom refusal UI (e.g. already-open -> "use here instead",
     * no-session -> "bring it back live"). Default: a refresh notice.
     *
     * `ctx.docVisible` says which surface the content lands on. False (the
     * pre-live refusals: already-open, room-full, version mismatch, a code
     * gate): the content IS the page, exactly as before — render a full
     * screen. True (post-live refusals: idle-timeout, session-expired): the
     * document is still in hand and stays on screen READ-ONLY behind the
     * content, which is rendered in an overlay covering it — render a
     * dialog/scrim, not a full page. Existing one-argument hosts keep their
     * old markup untouched; the second argument only ADDS the information. */
    refusedContent?: (reason: string, ctx?: {
        docVisible: boolean;
    }) => ReactNode;
}): ReactNode;

export { BundlePersister, type BundleStore, COLLAB_TOOLBAR_DEFAULTS, CollabEditor, type CollabSession, type DocBundle, InMemoryBundleStore, IndexedDbBundleStore, type CollabSession as InjectedCollabSession, type ParsedBundleKey, type StoredDocKind, type StoredDocSummary, UndoOutcome, type UseCollabOptions, docHash, draftKey, mediaAddressesOf, parseBundleKey, supersededKey, useCollab, versionKey };
