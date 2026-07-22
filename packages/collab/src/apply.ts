import {
  DocxDocument,
  StableIds,
  applyInsertText,
  applySplitParagraph,
  applyDeleteRange,
  type EditCaret,
  type Run,
  type Block,
  type XmlElement,
} from "@wordinweb/core";
import { Intent, Position } from "./intents.js";

/**
 * Bridge from wire intents to the core mutation functions. Resolves stable
 * ids to model positions, applies the mutation headlessly (no DOM), and
 * assigns carried ids to newly created nodes so every replica agrees on ids
 * (plan doc 03). Callers refresh() and re-run the id table afterward; this
 * function performs the tree mutation only.
 *
 * Returns true if the intent applied, false for a clean no-op / unresolvable
 * position (the caller records a rejection, doc 03).
 */
export function applyIntent(doc: DocxDocument, ids: StableIds, intent: Intent): boolean {
  const runMap = buildRunMap(doc);
  // Headless apply does not yet support suggesting mode (needs provenance
  // threaded through the intent); the session forbids it upstream.
  const ctx = { suggesting: false, revMeta: () => { throw new Error("suggesting mode unsupported headlessly"); } };

  switch (intent.kind) {
    case "insertText": {
      const caret = resolveCaret(ids, runMap, intent.at);
      if (!caret) return false;
      applyInsertText(doc, caret, intent.text, ctx);
      return true;
    }
    case "deleteText": {
      if (intent.end <= intent.start) return false;
      const caret = resolveCaret(ids, runMap, { blockId: intent.blockId, runId: intent.runId, offset: intent.start });
      if (!caret) return false;
      if (intent.end > caret.t.text.length) return false;
      // Engine-independent splice via the shared core mutation (offsets were
      // resolved client-side, so no Intl.Segmenter dependency here).
      applyDeleteRange(caret, intent.start, intent.end);
      return true;
    }
    case "splitParagraph": {
      const caret = resolveCaret(ids, runMap, intent.at);
      if (!caret) return false;
      const res = applySplitParagraph(doc, caret, ctx);
      if (!res) return false;
      ids.assign(res.after, intent.newBlockId);
      const newRun = res.after.children.find((c) => localRun(c));
      if (newRun) ids.assign(newRun, intent.newRunId);
      return true;
    }
  }
}

function localRun(el: XmlElement): boolean {
  const n = el.name;
  return n === "w:r" || n.endsWith(":r") || n === "r";
}

interface RunEntry {
  run: Run;
  firstT: XmlElement | undefined;
}

/** Map a run's source XML element to its model Run (needed for the checkbox
 * guard in applyInsertText and to locate the run's first w:t). Rebuilt per
 * apply — cheap relative to refresh(); a batching session can hoist it. */
function buildRunMap(doc: DocxDocument): Map<XmlElement, RunEntry> {
  const map = new Map<XmlElement, RunEntry>();
  const visitRun = (run: Run): void => {
    if (!run.src) return;
    const firstT = run.content.find((c) => c.kind === "text")?.srcT;
    map.set(run.src, { run, firstT });
  };
  const visitBlocks = (blocks: Block[]): void => {
    for (const b of blocks) {
      if (b.type === "paragraph") {
        for (const item of b.children) {
          if (item.type === "run") visitRun(item);
          else for (const r of item.runs) visitRun(r);
        }
      } else if (b.type === "table") {
        for (const row of b.rows) for (const cell of row.cells) visitBlocks(cell.blocks);
      }
    }
  };
  for (const section of doc.sections) visitBlocks(section.blocks);
  return map;
}

function resolveCaret(ids: StableIds, runMap: Map<XmlElement, RunEntry>, pos: Position): EditCaret | null {
  const runEl = ids.elOf(pos.runId);
  if (!runEl) return null;
  const entry = runMap.get(runEl);
  if (!entry || !entry.firstT) return null;
  if (pos.offset < 0 || pos.offset > entry.firstT.text.length) return null;
  return { t: entry.firstT, run: entry.run, offset: pos.offset };
}
