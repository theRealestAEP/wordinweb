/**
 * Wire intents for replicated editing.
 *
 * Positions address content by stable id (plan doc 02): `{ blockId, runId,
 * offset }`, where offset is a character offset within the run's text. Ids
 * come from the core StableIds side table; the server validates every
 * client-supplied id before applying (validate-then-install, doc 11) — the
 * types here are the transport shape, not a trust boundary.
 *
 * `base` is the last server sequence number the originating client had
 * applied when it produced the intent; the server transforms the intent
 * against everything sequenced in `(base, seq)` before applying and
 * broadcasting the canonical form (doc 03).
 */

export type StableId = number;

/** A caret/endpoint position within a run. */
export interface Position {
  blockId: StableId;
  runId: StableId;
  offset: number;
}

interface IntentBase {
  /** Originating client (authenticated identity; the server never trusts a
   * claimed id — see doc 11). */
  clientId: string;
  /** Per-client monotonic counter; `(clientId, clientSeq)` is the idempotency
   * key that dedups re-sends after reconnect (doc 03). */
  clientSeq: number;
  /** Last applied server seq at production time. */
  base: number;
}

/** Insert `text` at a position. */
export interface InsertTextIntent extends IntentBase {
  kind: "insertText";
  at: Position;
  text: string;
}

/** Delete `[start, end)` characters within a single run. */
export interface DeleteTextIntent extends IntentBase {
  kind: "deleteText";
  blockId: StableId;
  runId: StableId;
  start: number;
  end: number;
}

/**
 * Split a paragraph at a position. Text after `at.offset` in run `at.runId`
 * (and following runs) moves to a new paragraph. The originating client
 * allocates the ids for the new paragraph and the new run holding the moved
 * tail; every replica installs these recorded values (doc 03), and they also
 * let the transform remap concurrent positions that fall in the moved tail.
 */
export interface SplitParagraphIntent extends IntentBase {
  kind: "splitParagraph";
  at: Position;
  newBlockId: StableId;
  newRunId: StableId;
}

/**
 * Character-format an entire run (bold/italic/underline/…). Whole-run only:
 * it mutates the run's w:rPr in place, so no run splits, no new ids, and the
 * run id is preserved — which makes its transform identity (formatting moves
 * no text). Sub-range formatting splits the run into up to three, needing the
 * run-split id-inheritance + position remapping (plan doc 03 F3); that is a
 * documented extension, not implemented here.
 */
export interface FormatRunIntent extends IntentBase {
  kind: "formatRun";
  blockId: StableId;
  runId: StableId;
  /** RunFormatPatch (bold/italic/underline/strike/color/…). Structural shape
   * mirrors @wordinweb/core's RunFormatPatch; carried verbatim. */
  patch: Record<string, unknown>;
}

/**
 * Paragraph-level formatting: alignment and/or paragraph style. Block-level —
 * it mutates the paragraph's w:pPr, creates no tracked nodes, and moves no
 * text, so the block id is preserved and its transform is identity.
 */
export interface FormatParagraphIntent extends IntentBase {
  kind: "formatParagraph";
  blockId: StableId;
  align?: "left" | "center" | "right" | "justify";
  /** Paragraph style id; null clears to Normal. Omit to leave unchanged. */
  styleId?: string | null;
}

export type Intent =
  | InsertTextIntent
  | DeleteTextIntent
  | SplitParagraphIntent
  | FormatRunIntent
  | FormatParagraphIntent;

/** A sequenced log entry: an applied intent with its assigned seq, or a
 * rejection no-op (doc 03) that still occupies a position in the total order
 * so every replica agrees where a drop took effect. */
export type LogEntry =
  | { seq: number; kind: "applied"; intent: Intent }
  | { seq: number; kind: "rejected"; clientId: string; clientSeq: number; reason: string };

export function idempotencyKey(i: { clientId: string; clientSeq: number }): string {
  return `${i.clientId}:${i.clientSeq}`;
}
