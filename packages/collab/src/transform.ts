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
  /** Present for a format-range split: the old run is replaced by up to three
   * pieces and a position remaps by which piece it falls into. */
  formatSplit?: { start: number; end: number; beforeId?: number; middleId: number; afterId?: number };
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
    case "formatParagraph":
      // Block-level; moves no text, preserves ids.
      return [];
    case "setListType":
      // Block-level; moves no text, preserves ids.
      return [];
    case "formatRange":
      // The old run is replaced by up to three pieces; positions remap.
      return [{
        runId: intent.runId, at: intent.start, del: 0, ins: 0,
        formatSplit: { start: intent.start, end: intent.end, beforeId: intent.beforeId, middleId: intent.middleId, afterId: intent.afterId },
      }];
    case "tableOp":
      // Non-node-creating table ops don't shift any surviving run's offsets;
      // concurrent edits to deleted cells fail cleanly at apply (retired ids).
      return [];
    case "mergeParagraph":
      // Runs move (identity preserved) into the previous paragraph; run-
      // addressed offsets are unchanged.
      return [];
    case "commentRun":
      // Comment markers are run siblings; no run's text offsets shift.
      return [];
    case "pasteBlocks":
      // Separate new blocks inserted after a paragraph; no existing offsets shift.
      return [];
    case "suggestRevision":
      // Strikes WRAP text (w:del) and marks annotate paragraph glyphs —
      // no text moves, so concurrent intents' positions are unaffected.
      // (The wrap does split runs at strike boundaries; a concurrent
      // same-run edit past the boundary resolves-or-rejects IDENTICALLY on
      // every replica — degraded concurrency fidelity, never divergence.
      // The formatRange-style split remap is the designed follow-on.)
      return [];
    case "insertImage":
      // Image is a sibling drawing run; no run's text offsets shift.
      return [];
    case "insertBreak":
      // Break at end-of-run inserts sibling runs; no text offsets shift.
      return [];
    case "insertMath":
    case "insertShape":
      // Sibling insertion at end of run; no text offsets shift.
      return [];
    case "replyComment":
      // Comment-part change; no document run offsets shift.
      return [];
    case "adjustIndent":
    case "setSpacing":
      // Block-level; no text offsets shift.
      return [];
    case "insertPageField":
      // Sibling field run at end of run; no text offsets shift.
      return [];
    case "setLink":
    case "insertFootnote":
    case "setDropCap":
    case "setDivider":
    case "insertBookmark":
    case "insertBlankPage":
    case "insertSectionBreak":
    case "insertCrossRef":
    case "insertCoverPage":
    case "setPageLayout":
    case "setListLevel":
    case "insertWordArt":
    case "insertChart":
    case "insertSmartArt":
    case "setLineNumbering":
    case "insertDateTimeField":
    case "insertField":
    case "setDrawingRotation":
    case "setDrawingFill":
    case "setChartData":
    case "setSmartArtNodeText":
    case "setDrawingWordArtText":
    case "setDrawingLineStyle":
    case "setImageAltText":
    case "removeLink":
    case "setImageWrap":
    case "setDrawingOrder":
    case "setSmartArtData":
    case "setSmartArtFill":
    case "setSmartArtTextFormat":
    case "setFloatingPagePosition":
    case "resizeDrawing":
    case "resizeTableColumn":
    case "resizeTableRow":
    case "moveTable":
    case "removeDrawing":
    case "setMathLinear":
    case "deleteMath":
    case "ensureHeaderFooter":
    case "deleteComment":
    case "insertBookmarkRange":
    case "toggleCheckbox":
      // Run/block/document-level; no existing run's text offsets shift.
      return [];
    case "moveMath":
      // Same shape as insertMath above: dropping the equation mid-text splits
      // the destination run, and that split is not modeled here — a concurrent
      // edit into the tail half of that run resolves-or-rejects identically on
      // every replica (degraded concurrency fidelity, never divergence). The
      // movedToRunId remap is the designed follow-on for both.
      return [];
    case "insertTable":
    case "acceptRevision":
    case "rejectRevision":
    case "acceptAllRevisions":
    case "rejectAllRevisions":
      // Accept/reject CAN change text (an accepted deletion / rejected
      // insertion removes runs). Under the session's one-in-flight constraint
      // these administrative review ops don't overlap a concurrent text edit,
      // so no cross-intent RunEdit is modeled here (see plan doc 03: multi-
      // pending OT is out of scope).
      return [];
  }
}

/** Adjust a single position against one prior run edit. `leftGravity` is for
 * a position that should NOT move when an insert lands exactly on it (e.g. the
 * END of a delete range — text inserted at the end stays outside the range).
 * Default (right gravity) shifts a position right when an insert is at it. */
export function transformPositionAgainstEdit(pos: Position, edit: RunEdit, leftGravity = false): Position {
  if (pos.runId !== edit.runId) return pos;

  if (edit.formatSplit) {
    // The old run split into before[0,start) / middle[start,end) / after[end,).
    // Remap the position into whichever piece it lands in (pieces absent when
    // start==0 or end==runLen — positions can't fall in those ranges).
    const { start, end, beforeId, middleId, afterId } = edit.formatSplit;
    if (pos.offset < start && beforeId !== undefined) {
      return { blockId: pos.blockId, runId: beforeId, offset: pos.offset };
    }
    if (pos.offset <= end) {
      return { blockId: pos.blockId, runId: middleId, offset: pos.offset - start };
    }
    if (afterId !== undefined) {
      return { blockId: pos.blockId, runId: afterId, offset: pos.offset - end };
    }
    return { blockId: pos.blockId, runId: middleId, offset: pos.offset - start };
  }

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
  // A pure insert (del==0) exactly at the position: right gravity shifts the
  // position after the inserted text; left gravity keeps it before.
  if (edit.del === 0 && edit.at === pos.offset) {
    return leftGravity ? pos : { ...pos, offset: pos.offset + edit.ins };
  }
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
export function transformPosition(pos: Position, ahead: Intent[], leftGravity = false): Position {
  let out = pos;
  for (const intent of ahead) {
    for (const edit of runEditsOf(intent)) {
      out = transformPositionAgainstEdit(out, edit, leftGravity);
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
    case "formatParagraph":
      // Addressed by block id only; nothing to transform.
      return { ...intent, base: newBase };
    case "setListType":
      return { ...intent, base: newBase };
    case "formatRange": {
      // Transform the [start,end) endpoints against prior edits in the run.
      const s = transformPosition({ blockId: intent.blockId, runId: intent.runId, offset: intent.start }, ahead);
      const e = transformPosition({ blockId: intent.blockId, runId: intent.runId, offset: intent.end }, ahead);
      // If the endpoints landed in different runs (a prior split), the range
      // no longer describes one run — collapse to a no-op the apply rejects.
      if (e.runId !== s.runId) return { ...intent, runId: s.runId, blockId: s.blockId, start: s.offset, end: s.offset, base: newBase };
      return { ...intent, runId: s.runId, blockId: s.blockId, start: s.offset, end: e.offset, base: newBase };
    }
    case "tableOp":
      // Addressed by a stable paragraph id; nothing to transform.
      return { ...intent, base: newBase };
    case "mergeParagraph":
      return { ...intent, base: newBase };
    case "commentRun":
      return { ...intent, base: newBase };
    case "pasteBlocks":
      return { ...intent, base: newBase };
    case "suggestRevision": {
      if (!intent.ranges?.length) return { ...intent, base: newBase };
      const ranges = intent.ranges
        .map((r) => {
          const s2 = transformPosition({ blockId: r.blockId, runId: r.runId, offset: r.start }, ahead);
          const e2 = transformPosition({ blockId: r.blockId, runId: r.runId, offset: r.end }, ahead);
          // Endpoints scattered across a prior split: collapse to empty —
          // apply drops empty ranges (clean no-op, identical everywhere).
          if (s2.runId !== e2.runId || e2.offset <= s2.offset) return { ...r, start: 0, end: 0 };
          return { blockId: s2.blockId, runId: s2.runId, start: s2.offset, end: e2.offset };
        })
        .filter((r) => r.end > r.start);
      return { ...intent, ranges, base: newBase };
    }
    case "insertImage":
      return { ...intent, base: newBase };
    case "insertBreak":
      return { ...intent, base: newBase };
    case "insertMath":
    case "insertShape":
      return { ...intent, base: newBase };
    case "moveMath":
      // The equation is block-addressed, but the DROP is an ordinary caret and
      // shifts with prior edits in the destination run.
      return { ...intent, at: transformPosition(intent.at, ahead), base: newBase };
    case "replyComment":
      return { ...intent, base: newBase };
    case "adjustIndent":
    case "setSpacing":
    case "insertPageField":
    case "setLink":
    case "insertFootnote":
    case "setDropCap":
    case "setDivider":
    case "insertBookmark":
    case "insertBlankPage":
    case "insertSectionBreak":
    case "insertCrossRef":
    case "insertCoverPage":
    case "setPageLayout":
    case "setListLevel":
    case "insertWordArt":
    case "insertChart":
    case "insertSmartArt":
    case "setLineNumbering":
    case "insertDateTimeField":
    case "insertField":
    case "setDrawingRotation":
    case "setDrawingFill":
    case "setChartData":
    case "setSmartArtNodeText":
    case "setDrawingWordArtText":
    case "setDrawingLineStyle":
    case "setImageAltText":
    case "removeLink":
    case "setImageWrap":
    case "setDrawingOrder":
    case "setSmartArtData":
    case "setSmartArtFill":
    case "setSmartArtTextFormat":
    case "setFloatingPagePosition":
    case "resizeDrawing":
    case "resizeTableColumn":
    case "resizeTableRow":
    case "moveTable":
    case "removeDrawing":
    case "setMathLinear":
    case "deleteMath":
    case "ensureHeaderFooter":
    case "deleteComment":
    case "insertBookmarkRange":
    case "toggleCheckbox":
    case "insertTable":
    case "acceptRevision":
    case "rejectRevision":
    case "acceptAllRevisions":
    case "rejectAllRevisions":
      return { ...intent, base: newBase };
    case "deleteText": {
      const s = transformPosition({ blockId: intent.blockId, runId: intent.runId, offset: intent.start }, ahead);
      // The delete's END uses left gravity: text inserted exactly at the end
      // stays OUTSIDE the deleted range (so an undo of an insert doesn't
      // swallow content a concurrent insert placed right after it).
      const e = transformPosition({ blockId: intent.blockId, runId: intent.runId, offset: intent.end }, ahead, true);
      // A delete whose endpoints were split into different runs no longer
      // describes a single-run range; callers treat a collapsed/cross-run
      // range as a no-op rather than corrupting text.
      return { ...intent, runId: s.runId, blockId: s.blockId, start: s.offset, end: e.runId === s.runId ? e.offset : s.offset, base: newBase };
    }
  }
}
