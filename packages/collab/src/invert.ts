import { DocxDocument, StableIds } from "@wordinweb/core";
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
      const runEl = ids.elOf(intent.runId);
      if (!runEl) return null;
      const t = firstText(runEl);
      if (!t) return null;
      const removed = t.text.slice(intent.start, intent.end);
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
    // Formatting / table / comment / merge inverses need prior-state capture
    // (was-bold?, the merged paragraph's structure, …) — a documented
    // extension; not undoable yet.
    default:
      return null;
  }
}

function firstText(runEl: import("@wordinweb/core").XmlElement): import("@wordinweb/core").XmlElement | null {
  if (runEl.name === "w:t" || runEl.name.endsWith(":t")) return runEl;
  for (const c of runEl.children) {
    const f = firstText(c);
    if (f) return f;
  }
  return null;
}
