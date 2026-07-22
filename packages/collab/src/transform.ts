import { Intent, Position } from "./intents.js";

/**
 * The canonical transform (plan doc 03). Stable ids make structural
 * concurrency safe — a run is re-found by id regardless of what moved around
 * it — but a character *offset* within a run can be stale when a concurrent
 * edit changed that same run's text ahead of it. This adjusts positions
 * run-locally: OT-lite, not a general OT system, implemented once and used by
 * both the server (canonical) and the client (reconciliation replay).
 *
 * A prior intent's effect on a run is summarized as a RunEdit. Two shapes:
 *  - text change: at `at`, `del` chars removed then `ins` inserted (insert,
 *    delete).
 *  - split: text at/after `at` in `runId` moves to `movedToRunId` starting at
 *    offset 0 (splitParagraph). A position after the split point is remapped
 *    to the new run rather than shifted.
 */
export interface RunEdit {
  runId: number;
  at: number;
  del: number;
  ins: number;
  /** Present for splits: the run the tail (offset >= at) moved into. */
  movedToRunId?: number;
}

/** The run edits a prior intent imposes on the document (only same-run edits
 * matter to offset transformation; cross-run structure is handled by ids). */
export function runEditsOf(intent: Intent): RunEdit[] {
  switch (intent.kind) {
    case "insertText":
      return [{ runId: intent.at.runId, at: intent.at.offset, del: 0, ins: intent.text.length }];
    case "deleteText":
      return [{ runId: intent.runId, at: intent.start, del: intent.end - intent.start, ins: 0 }];
    case "splitParagraph":
      // The tail (offset >= at.offset) leaves this run for the new run.
      return [{ runId: intent.at.runId, at: intent.at.offset, del: 0, ins: 0, movedToRunId: intent.newRunId }];
    case "formatRun":
      // Whole-run formatting moves no text and preserves the run id — no
      // position of any concurrent intent is affected.
      return [];
  }
}

/** Adjust a single position against one prior run edit. */
export function transformPositionAgainstEdit(pos: Position, edit: RunEdit): Position {
  if (pos.runId !== edit.runId) return pos;

  if (edit.movedToRunId !== undefined) {
    // Split: everything strictly after the split point moved to the new run.
    // The split point itself stays with the original run (caret before the
    // break); content beyond it is re-addressed into the moved run.
    if (pos.offset > edit.at) {
      return { blockId: pos.blockId, runId: edit.movedToRunId, offset: pos.offset - edit.at };
    }
    return pos;
  }

  const delEnd = edit.at + edit.del;
  if (delEnd <= pos.offset) {
    // Edit lies fully before the position: shift by the net length change.
    return { ...pos, offset: pos.offset + edit.ins - edit.del };
  }
  if (edit.at >= pos.offset) {
    // Edit lies at or after the position: no shift.
    return pos;
  }
  // The position falls inside the deleted span: collapse to the edit point,
  // now sitting just after any inserted replacement text.
  return { ...pos, offset: edit.at + edit.ins };
}

/** Adjust a position against an ordered list of prior intents (oldest first).
 * `movedRunId` mapping composes: a position remapped into a new run by a split
 * is then transformed against later edits addressed to that new run. */
export function transformPosition(pos: Position, ahead: Intent[]): Position {
  let out = pos;
  for (const intent of ahead) {
    for (const edit of runEditsOf(intent)) {
      out = transformPositionAgainstEdit(out, edit);
    }
  }
  return out;
}

/** Transform an intent so its positions are valid against the state produced
 * by `ahead` (the intents sequenced since this intent's base). Returns a new
 * intent with adjusted positions; `base` is advanced past `ahead`. */
export function transformIntent(intent: Intent, ahead: Intent[]): Intent {
  const newBase = intent.base + ahead.length;
  switch (intent.kind) {
    case "insertText":
      return { ...intent, at: transformPosition(intent.at, ahead), base: newBase };
    case "splitParagraph":
      return { ...intent, at: transformPosition(intent.at, ahead), base: newBase };
    case "formatRun":
      // Addressed by run id only; nothing to transform (whole-run).
      return { ...intent, base: newBase };
    case "deleteText": {
      const s = transformPosition({ blockId: intent.blockId, runId: intent.runId, offset: intent.start }, ahead);
      const e = transformPosition({ blockId: intent.blockId, runId: intent.runId, offset: intent.end }, ahead);
      // A delete whose endpoints were split into different runs no longer
      // describes a single-run range; callers treat a collapsed/cross-run
      // range as a no-op rather than corrupting text.
      return { ...intent, runId: s.runId, blockId: s.blockId, start: s.offset, end: e.runId === s.runId ? e.offset : s.offset, base: newBase };
    }
  }
}
