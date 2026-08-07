import type { Intent } from "./intents.js";
import type { CollabConnection } from "./connection.js";

/**
 * Offline-edit reconciliation (plan doc 15 §2 / the arrival ladder). Zero
 * custody means a copy edited while the session was gone must reconcile on
 * rejoin. This module does NOT do a document merge — the architecture's core
 * primitive is already "replay my edits on a changed base," so an offline
 * tail replays through the same submit path a reconnecting client's pending
 * queue uses. What lives here is the two pieces that make the OWNER-adopted
 * default — rebase-as-SUGGESTIONS — work: the intent→suggestion converter,
 * and a drained replay driver (one-in-flight discipline, the supported
 * mode; a bounded offline tail replayed with drains needs no sibling math).
 */

/**
 * How many intents the offline tail may hold before RECORDING STOPS AND THE
 * EDITOR IS GATED (doc 15 §4.3: the tail must be bounded). The number:
 *
 *  - Anything past the suggest threshold (50) arrives as a draft anyway, so
 *    the cap governs only how much offline work the draft can retain — not
 *    the review burden, which the ladder already bounds.
 *  - Each entry is a small JSON object (~100 bytes), and the tail rides
 *    EVERY throttled bundle write, so the cap bounds write amplification to
 *    ~200KB per write in the worst case.
 *  - 2000 keystroke-grade intents is well past the feature's design target
 *    (doc 15 §5: "the laptop-closed-on-the-train case — small tails, high
 *    frequency"); at that scale the honest answer is a downloaded copy.
 *
 * AT THE CAP THE GATE CLOSES LOUDLY — the editor goes read-only and the UI
 * says so — rather than recording silently stopping, which would be the
 * exact accept-then-lose failure the write gate exists to prevent.
 */
export const OFFLINE_TAIL_CAP = 2000;

/**
 * Convert an offline intent tail to its SUGGESTION form (doc 15: the
 * default reconciliation). Text-producing intents become tracked changes —
 * visible, attributed, non-destructive (Word keeps both sides), reviewable
 * with the already-wired accept/reject flow:
 *
 *  - `insertText`  → `insertText` with `suggest{author,date}` (a w:ins).
 *  - `deleteText`  → `suggestRevision` striking that range (a w:del wrap;
 *                    nothing is removed until the crowd accepts).
 *  - `splitParagraph` → a split with an inserted paragraph-mark revision.
 *
 * Other structural intents (merges, tables, images, format) have no
 * clean non-destructive suggestion form and are DROPPED from the suggestion
 * replay — surfaced to the caller as `dropped` so the UI can say "N
 * structural changes couldn't be suggested; keep them in a draft instead"
 * (no silent loss — doc 16 §honest-failure discipline applied to rebase).
 * The alternative path (destructive rebase) keeps them but is gated on the
 * doc-03 sibling math for tails > 1.
 */
export function toSuggestions(
  tail: Intent[],
  author: string,
  date: string,
): { suggestions: Omit<Intent, "clientId" | "clientSeq" | "base">[]; dropped: Intent[] } {
  const suggestions: Omit<Intent, "clientId" | "clientSeq" | "base">[] = [];
  const dropped: Intent[] = [];
  for (const intent of tail) {
    if (intent.kind === "insertText") {
      suggestions.push({ kind: "insertText", at: intent.at, text: intent.text, suggest: { author, date } } as never);
    } else if (intent.kind === "insertSeparator") {
      suggestions.push({ kind: "insertSeparator", at: intent.at, separator: intent.separator, suggest: { author, date } } as never);
    } else if (intent.kind === "deleteText") {
      suggestions.push({
        kind: "suggestRevision",
        ranges: [{ blockId: intent.blockId, runId: intent.runId, start: intent.start, end: intent.end }],
        suggest: { author, date },
      } as never);
    } else if (intent.kind === "splitParagraph") {
      suggestions.push({
        kind: "splitParagraph",
        at: intent.at,
        newBlockId: intent.newBlockId,
        newRunId: intent.newRunId,
        suggest: { author, date },
      } as never);
    } else {
      dropped.push(intent);
    }
  }
  return { suggestions, dropped };
}

/**
 * Replay a list of local intents through a connection with the one-in-flight
 * discipline: submit one, wait for it to confirm (drain), submit the next.
 * This is the supported concurrency mode (doc 03) — a bounded offline tail
 * replayed with drains never needs multi-pending sibling rebasing, which is
 * why rebase-as-suggestions ships now while destructive multi-pending rebase
 * stays gated. Each replayed intent transforms against everything the crowd
 * did while you were gone (the connection's normal submit path), so it lands
 * correctly on the current document, not the one you left.
 *
 * `waitDrained` is injected (the connection exposes `pendingCount` via its
 * replica; the app polls it, or in tests a synchronous loopback drains
 * instantly). Returns when the whole tail is submitted-and-drained.
 */
export async function replayDrained(
  conn: CollabConnection,
  intents: Omit<Intent, "clientId" | "clientSeq" | "base">[],
  opts: { submit: (i: Omit<Intent, "clientId" | "clientSeq" | "base">) => void; waitDrained: () => Promise<void> },
): Promise<void> {
  for (const intent of intents) {
    opts.submit(intent);
    await opts.waitDrained();
  }
  void conn;
}

/**
 * The arrival-ladder decision (doc 15): given the offline tail size and the
 * divergence verdict, pick the reconciliation mode. Pure — the UI renders
 * the choice; this encodes the policy so it is testable and consistent.
 */
export function arrivalMode(
  tailLength: number,
  diverged: boolean,
  opts: { suggestThreshold?: number } = {},
): "fast-forward" | "suggest" | "draft" {
  if (!diverged) return "fast-forward";
  const threshold = opts.suggestThreshold ?? 50;
  // Small divergence → suggestions (the default); large → draft only (a
  // giant suggestion pile is review spam — the burden stays on the incomer).
  return tailLength <= threshold ? "suggest" : "draft";
}
