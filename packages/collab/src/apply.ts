import {
  DocxDocument,
  StableIds,
  applyInsertText,
  applySplitParagraph,
  applyDeleteRange,
  applyRunFormat,
  setParagraphAlignment,
  setParagraphStyle,
  setListType,
  applyTableOp,
  mergeParagraphBackward,
  addComment,
  recordedProvenance,
  validatePastedOoxml,
  parseXml,
  localName,
  type EditCaret,
  type Run,
  type Block,
  type XmlElement,
  type SelectionSegment,
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
    case "formatRun": {
      const runEl = ids.elOf(intent.runId);
      if (!runEl) return false;
      const entry = runMap.get(runEl);
      if (!entry) return false;
      // Whole-run format: a segment with t=null covers the entire run, so
      // applyRunFormat takes the in-place (no-split) path — run id preserved.
      const seg: SelectionSegment = { run: entry.run, t: null, start: 0, end: 0, props: entry.run.props };
      applyRunFormat(doc, [seg], intent.patch as never);
      return true;
    }
    case "formatParagraph": {
      const blockEl = ids.elOf(intent.blockId);
      if (!blockEl) return false;
      // setParagraphAlignment/Style resolve the paragraph by walking UP from a
      // target, so pass a descendant w:t (or the block itself as a fallback).
      const target = firstTextDescendant(blockEl) ?? blockEl;
      let changed = false;
      if (intent.align) changed = setParagraphAlignment(doc, [target], intent.align) || changed;
      if (intent.styleId !== undefined) changed = setParagraphStyle(doc, [target], intent.styleId) || changed;
      return changed;
    }
    case "setListType": {
      const blockEl = ids.elOf(intent.blockId);
      if (!blockEl) return false;
      const target = firstTextDescendant(blockEl) ?? blockEl;
      return setListType(doc, [target], intent.listKind);
    }
    case "formatRange": {
      if (intent.end <= intent.start) return false;
      const runEl = ids.elOf(intent.runId);
      if (!runEl) return false;
      const entry = runMap.get(runEl);
      if (!entry || !entry.firstT) return false;
      const len = entry.firstT.text.length;
      if (intent.start < 0 || intent.end > len) return false;
      // Sub-range format: splits the run into before/middle/after (all new).
      const seg: SelectionSegment = { run: entry.run, t: entry.firstT, start: intent.start, end: intent.end, props: entry.run.props };
      const formatted = applyRunFormat(doc, [seg], intent.patch as never);
      if (formatted.length === 0) return false;
      // Locate the pieces via the returned middle w:t (robust to whatever
      // parent applyRunFormat spliced into): middle run = parent of middleT;
      // before/after are its previous/next run siblings.
      const middleT = formatted[0].t;
      const middleRun = doc.findParentOf(middleT);
      if (!middleRun) return true; // formatted but not id-splittable; ids stay as-is
      const container = doc.findParentOf(middleRun);
      if (!container) return true;
      const mIdx = container.children.indexOf(middleRun);
      // reassign (not assign): applyRunFormat's internal refresh() may have
      // auto-assigned parse-order ids to the pieces; override with the carried
      // ones so every replica addresses them identically.
      ids.reassign(middleRun, intent.middleId);
      if (intent.start > 0 && intent.beforeId !== undefined) {
        const before = container.children[mIdx - 1];
        if (before && localRun(before)) ids.reassign(before, intent.beforeId);
      }
      if (intent.end < len && intent.afterId !== undefined) {
        const after = container.children[mIdx + 1];
        if (after && localRun(after)) ids.reassign(after, intent.afterId);
      }
      return true;
    }
    case "tableOp": {
      const paraEl = ids.elOf(intent.cellParagraphId);
      if (!paraEl) return false;
      const target = firstTextDescendant(paraEl) ?? paraEl;
      const ok = applyTableOp(doc, target, intent.op as never);
      // Structural table ops removed nodes; retire their stale ids so the
      // table stops resolving deleted content.
      if (ok) ids.prune(doc.editableRoots());
      return ok;
    }
    case "mergeParagraph": {
      const pEl = ids.elOf(intent.blockId);
      if (!pEl) return false;
      const ok = mergeParagraphBackward(doc, pEl);
      if (ok) ids.prune(doc.editableRoots()); // retire the merged paragraph's id
      return ok;
    }
    case "commentRun": {
      const runEl = ids.elOf(intent.runId);
      if (!runEl) return false;
      const entry = runMap.get(runEl);
      if (!entry) return false;
      // Whole-run comment (t=null covers the run, so no run split). Carried
      // provenance makes the w14:paraId and w:date identical on every replica.
      const seg: SelectionSegment = { run: entry.run, t: null, start: 0, end: 0, props: entry.run.props };
      const prov = recordedProvenance({ dates: [intent.date], paraIds: [intent.paraId] });
      const ok = addComment(doc, [seg], intent.text, intent.author, intent.initials, prov);
      // addComment created a commentReference run + comment paragraph; keep the
      // id table filled for them (parse-order, consistent across live replicas).
      if (ok && doc.stableIds) doc.stableIds.assignFromRoots(doc.editableRoots());
      return ok;
    }
    case "pasteBlocks": {
      const afterEl = ids.elOf(intent.afterBlockId);
      if (!afterEl) return false;
      const parent = doc.findParentOf(afterEl);
      if (!parent) return false;
      // Parse + VALIDATE the client-supplied OOXML before it enters the tree
      // (doc 11 gate 2). Reject on any violation.
      let blocks: XmlElement[];
      try {
        const root = parseXml(`<w:root xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${intent.blocksXml}</w:root>`);
        blocks = root.children.filter((c) => localName(c.name) === "p" || localName(c.name) === "tbl");
      } catch {
        return false;
      }
      if (blocks.length === 0) return false;
      const check = validatePastedOoxml(blocks);
      if (!check.ok) return false;
      // Splice after the target paragraph.
      const at = parent.children.indexOf(afterEl);
      parent.children.splice(at + 1, 0, ...blocks);
      doc.refresh();
      // Assign carried ids to the new tracked nodes in document order,
      // overriding any parse-order ids refresh() gave them.
      const newTracked: XmlElement[] = [];
      const walk = (el: XmlElement): void => {
        const ln = localName(el.name);
        if (ln === "p" || ln === "tbl" || ln === "r") newTracked.push(el);
        el.children.forEach(walk);
      };
      blocks.forEach(walk);
      for (let k = 0; k < newTracked.length && k < intent.nodeIds.length; k++) {
        ids.reassign(newTracked[k], intent.nodeIds[k]);
      }
      return true;
    }
  }
}

function localRun(el: XmlElement): boolean {
  const n = el.name;
  return n === "w:r" || n.endsWith(":r") || n === "r";
}

function firstTextDescendant(el: XmlElement): XmlElement | null {
  if (el.name === "w:t" || el.name.endsWith(":t")) return el;
  for (const c of el.children) {
    const found = firstTextDescendant(c);
    if (found) return found;
  }
  return null;
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
