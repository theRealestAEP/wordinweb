import { DocxDocument, StableIds, resolveRunOffset } from "@wordinweb/core";
import { Intent, IntentBody } from "./intents.js";

/**
 * Collaborative undo via rebased inverses (plan doc 03 Phase 8). The inverse
 * of an applied intent is computed at apply time (when the pre-state is known)
 * and stored. Undo submits that inverse with `base` set to the ORIGINAL
 * intent's sequence number, so the existing canonical transform rebases it
 * against every intent sequenced since — which is exactly what makes undo
 * correct under concurrency (the flaw round-2 F5 identified in the
 * inverse-at-fixed-state designs is resolved by rebasing through the log,
 * ProseMirror-style).
 *
 * Returns null for intents whose inverse isn't defined yet (formatting/table/
 * comment undo needs prior-state capture — a documented extension); those are
 * simply not undoable for now.
 */
export function invertIntent(doc: DocxDocument, ids: StableIds, intent: Intent): IntentBody | null {
  switch (intent.kind) {
    case "insertText": {
      // Inverse: delete the text that was inserted.
      return {
        kind: "deleteText",
        blockId: intent.at.blockId,
        runId: intent.at.runId,
        start: intent.at.offset,
        end: intent.at.offset + intent.text.length,
      };
    }
    case "deleteText": {
      // Inverse: re-insert the removed text — capture it from the live run
      // BEFORE the delete applies (the caller invokes invert pre-apply).
      // Wire offsets are cumulative within the run (checkpoint B1): resolve
      // to the concrete w:t and slice locally.
      const runEl = ids.elOf(intent.runId);
      if (!runEl) return null;
      const hit = resolveRunOffset(runEl, intent.start);
      if (!hit) return null;
      const localEnd = hit.offset + (intent.end - intent.start);
      if (localEnd > hit.t.text.length) return null; // cross-w:t range: apply rejects it too
      const removed = hit.t.text.slice(hit.offset, localEnd);
      if (removed.length === 0) return null;
      return {
        kind: "insertText",
        at: { blockId: intent.blockId, runId: intent.runId, offset: intent.start },
        text: removed,
      };
    }
    case "splitParagraph": {
      // Inverse: merge the newly created paragraph back into its predecessor.
      return { kind: "mergeParagraph", blockId: intent.newBlockId };
    }
    case "insertImage": {
      // Inverse: remove the drawing's carrier run. The apply gives the newly
      // created carrier the FIRST carried id (the same shape insertShape
      // produces), so the inverse addresses it without a new wire kind.
      //
      // The media PART stays registered and its bytes stay in the package —
      // undo is a document edit, not a relay operation. Re-doing the same
      // insert would re-upload to the same content address and hit the
      // relay's already-present path, so the orphaned registration costs a
      // dead relationship, never a wrong image.
      const carrier = intent.nodeIds?.[0];
      if (carrier === undefined) return null;
      return { kind: "removeDrawing", runId: carrier };
    }
    // Formatting / table / comment / merge inverses need prior-state capture
    // (was-bold?, the merged paragraph's structure, …) — a documented
    // extension; not undoable yet.
    default:
      return null;
  }
}
