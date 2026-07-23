import { CollabHub } from "./hub.js";

/**
 * Demo/showcase utilities implementing the two public modes from the threat
 * model (plan doc 11): magic-link (full-feature, unguessable-url-as-capability)
 * and party doc (round-robin strangers, minimal surface). These are the
 * testable server-side pieces; the frontend app consumes them. All abuse
 * limits and the enabled-intent allowlist live here, not in the core hub.
 */

/** Generate an unguessable document id. The url IS the access control for the
 * magic-link mode, so ids must be ≥128-bit random (doc 11). Uses the injected
 * randomness source (Node crypto in production; a stub in tests). */
export function makeDocId(randomBytes: (n: number) => Uint8Array): string {
  const bytes = randomBytes(16); // 128 bits
  let hex = "";
  for (const b of bytes) hex += b.toString(16).padStart(2, "0");
  return `d_${hex}`;
}

/**
 * Party-doc pool: a fixed set of shared documents that anonymous visitors are
 * round-robined onto, so "a bunch of strangers typing together" stays bounded
 * (doc 11 party mode). Assignment is deterministic given the join order — no
 * per-user document creation, capping storage and abuse surface.
 */
export class PartyPool {
  private next = 0;
  readonly docIds: string[];

  constructor(size: number, idFor: (i: number) => string = (i) => `party_${i}`) {
    if (size < 1) throw new Error("PartyPool size must be >= 1");
    this.docIds = Array.from({ length: size }, (_, i) => idFor(i));
  }

  /** Assign the next visitor to a party doc (round-robin over the pool). */
  assign(): string {
    const id = this.docIds[this.next % this.docIds.length];
    this.next++;
    return id;
  }
}

/** The enabled-intent allowlist per demo mode (doc 11 — "disabled by omission
 * is not a control"; the list is explicit). The frontend disables toolbar
 * actions not in the set; a hardened server would also reject them. */
export const DEMO_INTENT_ALLOWLIST = {
  // Full implemented feature set. Authored-URL intents (hyperlink/webVideo)
  // stay OFF until the scheme-allowlist gate; media/import stay off in the demo.
  magicLink: [
    "insertText",
    "deleteText",
    "splitParagraph",
    "mergeParagraph",
    "formatRun",
    "formatParagraph",
    "formatRange",
    "setListType",
    "tableOp",
  ],
  // Minimal surface: text-structural + presence only, so the highest-exposure
  // anonymous mode is structurally XSS-free (no formatting/table/authored-URL/
  // media/import sink enabled — only plain text editing).
  party: ["insertText", "deleteText", "splitParagraph", "mergeParagraph"],
} as const;

export type DemoMode = keyof typeof DEMO_INTENT_ALLOWLIST;

/** True if an intent kind is permitted in the given demo mode. */
export function intentAllowedInDemo(mode: DemoMode, kind: string): boolean {
  return (DEMO_INTENT_ALLOWLIST[mode] as readonly string[]).includes(kind);
}

/**
 * Simple per-key rate limiter (per-IP doc creation, per-connection intents —
 * doc 11 abuse limits). Token bucket; `now` injected for deterministic tests.
 */
export class RateLimiter {
  private buckets = new Map<string, { tokens: number; last: number }>();
  constructor(private capacity: number, private refillPerMs: number) {}

  /** Try to consume one token for `key` at time `nowMs`. Returns true if
   * allowed, false if the bucket is empty (rate exceeded). */
  allow(key: string, nowMs: number): boolean {
    const b = this.buckets.get(key) ?? { tokens: this.capacity, last: nowMs };
    b.tokens = Math.min(this.capacity, b.tokens + (nowMs - b.last) * this.refillPerMs);
    b.last = nowMs;
    if (b.tokens < 1) {
      this.buckets.set(key, b);
      return false;
    }
    b.tokens -= 1;
    this.buckets.set(key, b);
    return true;
  }
}

/** Build a demo hub factory (blank docs, in-memory storage) — used by the
 * frontend's dev server. The provider/storage are passed in by the caller so
 * this stays dependency-light. */
export type DemoHub = CollabHub;
