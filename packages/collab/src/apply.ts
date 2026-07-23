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
  adjustIndent,
  setParagraphSpacing,
  insertPageField,
  setLink,
  insertFootnote,
  setDropCapAt,
  setParagraphDivider,
  insertBookmarkAt,
  insertBlankPageAt,
  insertSectionBreak,
  insertCrossReference,
  insertCoverPage,
  setPageLayout,
  setListLevel,
  insertWordArtAt,
  insertChartAt,
  insertSmartArtAt,
  setLineNumbering,
  insertDateTimeField,
  insertField,
  setDrawingRotation,
  setDrawingFill,
  setChartData,
  setSmartArtNodeText,
  setDrawingWordArtText,
  setDrawingLineStyle,
  setImageAltText,
  removeLink,
  setImageWrap,
  setDrawingOrder,
  setSmartArtData,
  setSmartArtFill,
  setSmartArtTextFormat,
  setFloatingPagePosition,
  setMathLinear,
  deleteMath,
  deleteComment,
  insertBookmarkAroundSelection,
  checkboxStateElement,
  toggleCheckbox,
  collectRevisions,
  acceptRevision,
  rejectRevision,
  acceptAllRevisions,
  isSafeUrl,
  applyTableOp,
  mergeParagraphBackward,
  addComment,
  replyToComment,
  recordedProvenance,
  validatePastedOoxml,
  parseXml,
  localName,
  insertImageAt,
  insertBreakAt,
  insertMathAt,
  insertShapeAt,
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
      if (intent.suggest) {
        // Tracked change: record the insertion as a w:ins with carried
        // author/date; the revision w:id is scan-based (deterministic).
        const s = intent.suggest;
        const suggestCtx = { suggesting: true, revMeta: () => ({ author: s.author, date: s.date, nextId: () => doc.nextRevisionId() }) };
        const before = trackedSet(ids, doc);
        applyInsertText(doc, caret, intent.text, suggestCtx);
        if (doc.stableIds) { doc.refresh(); ids.assignFromRoots(doc.editableRoots()); }
        void before;
        return true;
      }
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
      const isInsert = intent.op === "rowAbove" || intent.op === "rowBelow" || intent.op === "colLeft" || intent.op === "colRight";
      // For insert ops, snapshot the tracked-node set so we can find the new
      // nodes afterward and give them the carried ids.
      const before = isInsert ? trackedSet(ids, doc) : null;
      const ok = applyTableOp(doc, target, intent.op as never);
      if (!ok) return false;
      if (isInsert && before && intent.nodeIds) {
        // Assign carried ids to the newly created tracked nodes in doc order.
        const fresh: XmlElement[] = [];
        walkTracked(doc, (el) => { if (!before.has(el)) fresh.push(el); });
        for (let k = 0; k < fresh.length && k < intent.nodeIds.length; k++) {
          ids.reassign(fresh[k], intent.nodeIds[k]);
        }
      } else {
        // Delete/shading ops: retire stale ids for removed content.
        ids.prune(doc.editableRoots());
      }
      return true;
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
    case "insertImage": {
      const runEl = ids.elOf(intent.runId);
      if (!runEl) return false;
      const entry = runMap.get(runEl);
      if (!entry || !entry.firstT) return false;
      let bytes: Uint8Array;
      try {
        bytes = base64ToBytes(intent.imageBase64);
      } catch {
        return false;
      }
      // Register the media part (rId/part name are scan-max, deterministic
      // across replicas applying the same canonical intent) and splice the
      // drawing run. Client-measured px are carried so layout is deterministic.
      const relId = doc.addImageResource(bytes, intent.ext);
      const before = trackedSet(ids, doc);
      const newRun = insertImageAt(doc, entry.firstT, relId, intent.widthPx, intent.heightPx);
      if (!newRun) return false;
      doc.refresh();
      const fresh: XmlElement[] = [];
      walkTracked(doc, (el) => { if (!before.has(el)) fresh.push(el); });
      for (let k = 0; k < fresh.length && k < intent.nodeIds.length; k++) {
        ids.reassign(fresh[k], intent.nodeIds[k]);
      }
      return true;
    }
    case "insertBlankPage": {
      const runEl = ids.elOf(intent.runId);
      if (!runEl) return false;
      const entry = runMap.get(runEl);
      if (!entry || !entry.firstT) return false;
      const before = trackedSet(ids, doc);
      const ok = insertBlankPageAt(doc, entry.firstT, entry.firstT.text.length);
      if (ok) assignFreshTracked(ids, doc, before, intent.nodeIds);
      return ok;
    }
    case "insertSectionBreak": {
      const runEl = ids.elOf(intent.runId);
      if (!runEl) return false;
      const entry = runMap.get(runEl);
      if (!entry || !entry.firstT) return false;
      const before = trackedSet(ids, doc);
      const ok = insertSectionBreak(doc, entry.firstT, intent.breakType);
      if (ok) assignFreshTracked(ids, doc, before, intent.nodeIds);
      return ok;
    }
    case "insertCrossRef": {
      const runEl = ids.elOf(intent.runId);
      if (!runEl) return false;
      const entry = runMap.get(runEl);
      if (!entry || !entry.firstT) return false;
      const before = trackedSet(ids, doc);
      const ok = insertCrossReference(doc, entry.firstT, entry.firstT.text.length, intent.bookmark, intent.refKind);
      if (ok) assignFreshTracked(ids, doc, before, intent.nodeIds);
      return ok;
    }
    case "insertCoverPage": {
      const before = trackedSet(ids, doc);
      const ok = insertCoverPage(doc, intent.content as never);
      if (ok) assignFreshTracked(ids, doc, before, intent.nodeIds);
      return ok;
    }
    case "setPageLayout":
      // Document-level page setup (all sections). No new nodes to id.
      return setPageLayout(doc, intent.patch as never);
    case "setListLevel": {
      const blockEl = ids.elOf(intent.blockId);
      if (!blockEl) return false;
      const target = firstTextDescendant(blockEl) ?? blockEl;
      return setListLevel(doc, [target], intent.delta);
    }
    case "insertWordArt": {
      const runEl = ids.elOf(intent.runId);
      if (!runEl) return false;
      const entry = runMap.get(runEl);
      if (!entry || !entry.firstT) return false;
      const before = trackedSet(ids, doc);
      const drawing = insertWordArtAt(doc, entry.firstT, intent.text, intent.preset);
      if (!drawing) return false;
      assignFreshTracked(ids, doc, before, intent.nodeIds);
      return true;
    }
    case "insertChart": {
      const runEl = ids.elOf(intent.runId);
      if (!runEl) return false;
      const entry = runMap.get(runEl);
      if (!entry || !entry.firstT) return false;
      const before = trackedSet(ids, doc);
      const drawing = insertChartAt(doc, entry.firstT, intent.chart as never);
      if (!drawing) return false;
      assignFreshTracked(ids, doc, before, intent.nodeIds);
      return true;
    }
    case "insertSmartArt": {
      const runEl = ids.elOf(intent.runId);
      if (!runEl) return false;
      const entry = runMap.get(runEl);
      if (!entry || !entry.firstT) return false;
      const before = trackedSet(ids, doc);
      const drawing = insertSmartArtAt(doc, entry.firstT, intent.smartArt as never);
      if (!drawing) return false;
      assignFreshTracked(ids, doc, before, intent.nodeIds);
      return true;
    }
    case "setLineNumbering":
      // Document-level (all sections). No new nodes to id.
      return setLineNumbering(doc, intent.patch as never);
    case "insertDateTimeField": {
      const runEl = ids.elOf(intent.runId);
      if (!runEl) return false;
      const entry = runMap.get(runEl);
      if (!entry || !entry.firstT) return false;
      const before = trackedSet(ids, doc);
      const ok = insertDateTimeField(doc, entry.firstT, entry.firstT.text.length, intent.dtKind, intent.picture);
      if (ok) assignFreshTracked(ids, doc, before, intent.nodeIds);
      return ok;
    }
    case "insertField": {
      const runEl = ids.elOf(intent.runId);
      if (!runEl) return false;
      const entry = runMap.get(runEl);
      if (!entry || !entry.firstT) return false;
      const before = trackedSet(ids, doc);
      const ok = insertField(doc, entry.firstT, entry.firstT.text.length, intent.instruction, intent.cachedResult ?? "");
      if (ok) assignFreshTracked(ids, doc, before, intent.nodeIds);
      return ok;
    }
    case "setDrawingRotation": {
      const runEl = ids.elOf(intent.runId);
      if (!runEl) return false;
      const drawing = firstDrawingIn(runEl);
      if (!drawing) return false;
      return setDrawingRotation(doc, drawing, intent.degrees);
    }
    case "setDrawingFill": {
      const runEl = ids.elOf(intent.runId);
      if (!runEl) return false;
      const drawing = firstDrawingIn(runEl);
      if (!drawing) return false;
      return setDrawingFill(doc, drawing, intent.color);
    }
    case "setChartData": {
      const runEl = ids.elOf(intent.runId);
      if (!runEl) return false;
      const drawing = firstDrawingIn(runEl);
      if (!drawing) return false;
      return setChartData(doc, drawing, intent.chart as never);
    }
    case "setSmartArtNodeText": {
      const runEl = ids.elOf(intent.runId);
      if (!runEl) return false;
      const drawing = firstDrawingIn(runEl);
      if (!drawing) return false;
      return setSmartArtNodeText(doc, drawing, intent.index, intent.text);
    }
    case "setDrawingWordArtText": {
      const runEl = ids.elOf(intent.runId);
      if (!runEl) return false;
      const drawing = firstDrawingIn(runEl);
      if (!drawing) return false;
      return setDrawingWordArtText(doc, drawing, intent.text);
    }
    case "setDrawingLineStyle": {
      const runEl = ids.elOf(intent.runId);
      if (!runEl) return false;
      const drawing = firstDrawingIn(runEl);
      if (!drawing) return false;
      return setDrawingLineStyle(doc, drawing, intent.color, intent.widthPx, intent.dash);
    }
    case "setImageAltText": {
      const runEl = ids.elOf(intent.runId);
      if (!runEl) return false;
      const drawing = firstDrawingIn(runEl);
      if (!drawing) return false;
      return setImageAltText(doc, drawing, intent.alt);
    }
    case "removeLink": {
      const runEl = ids.elOf(intent.runId);
      if (!runEl) return false;
      const entry = runMap.get(runEl);
      if (!entry || !entry.firstT) return false;
      return removeLink(doc, entry.firstT);
    }
    case "setImageWrap": {
      const runEl = ids.elOf(intent.runId);
      if (!runEl) return false;
      const drawing = firstDrawingIn(runEl);
      if (!drawing) return false;
      return setImageWrap(doc, drawing, intent.mode);
    }
    case "setDrawingOrder": {
      const runEl = ids.elOf(intent.runId);
      if (!runEl) return false;
      const drawing = firstDrawingIn(runEl);
      if (!drawing) return false;
      return setDrawingOrder(doc, drawing, intent.order);
    }
    case "setSmartArtData": {
      const runEl = ids.elOf(intent.runId);
      if (!runEl) return false;
      const drawing = firstDrawingIn(runEl);
      if (!drawing) return false;
      return setSmartArtData(doc, drawing, intent.smartArt as never);
    }
    case "setSmartArtFill": {
      const runEl = ids.elOf(intent.runId);
      if (!runEl) return false;
      const drawing = firstDrawingIn(runEl);
      if (!drawing) return false;
      return setSmartArtFill(doc, drawing, intent.color, intent.nodeIndex);
    }
    case "setSmartArtTextFormat": {
      const runEl = ids.elOf(intent.runId);
      if (!runEl) return false;
      const drawing = firstDrawingIn(runEl);
      if (!drawing) return false;
      return setSmartArtTextFormat(doc, drawing, intent.format as never, intent.nodeIndex);
    }
    case "setFloatingPagePosition": {
      const runEl = ids.elOf(intent.runId);
      if (!runEl) return false;
      const drawing = firstDrawingIn(runEl);
      if (!drawing) return false;
      return setFloatingPagePosition(doc, drawing, intent.xPx, intent.yPx);
    }
    case "setMathLinear": {
      const blockEl = ids.elOf(intent.blockId);
      if (!blockEl) return false;
      const math = firstMathIn(blockEl);
      if (!math) return false;
      return setMathLinear(doc, math, intent.mathText);
    }
    case "deleteMath": {
      const blockEl = ids.elOf(intent.blockId);
      if (!blockEl) return false;
      const math = firstMathIn(blockEl);
      if (!math) return false;
      return deleteMath(doc, math);
    }
    case "deleteComment":
      return deleteComment(doc, intent.commentId);
    case "insertBookmarkRange": {
      const runEl = ids.elOf(intent.runId);
      if (!runEl) return false;
      const entry = runMap.get(runEl);
      if (!entry || !entry.firstT) return false;
      if (intent.end > entry.firstT.text.length) return false;
      const seg: SelectionSegment = { run: entry.run, t: entry.firstT, start: intent.start, end: intent.end, props: entry.run.props };
      return insertBookmarkAroundSelection(doc, [seg], intent.name);
    }
    case "toggleCheckbox": {
      const runEl = ids.elOf(intent.runId);
      if (!runEl) return false;
      const entry = runMap.get(runEl);
      if (!entry) return false;
      const cbEl = checkboxStateElement(entry.run, entry.firstT);
      if (!cbEl) return false;
      toggleCheckbox(doc, cbEl);
      return true;
    }
    case "acceptRevision": {
      const refs = collectRevisions(doc);
      if (intent.index >= refs.length) return false;
      return acceptRevision(doc, refs[intent.index]);
    }
    case "rejectRevision": {
      const refs = collectRevisions(doc);
      if (intent.index >= refs.length) return false;
      return rejectRevision(doc, refs[intent.index]);
    }
    case "acceptAllRevisions":
      return acceptAllRevisions(doc) > 0;
    case "setLink": {
      if (!isSafeUrl(intent.url)) return false; // reject javascript:/data: etc.
      const runEl = ids.elOf(intent.runId);
      if (!runEl) return false;
      const entry = runMap.get(runEl);
      if (!entry || !entry.firstT) return false;
      const seg: SelectionSegment = { run: entry.run, t: entry.firstT, start: 0, end: entry.firstT.text.length, props: entry.run.props };
      const before = trackedSet(ids, doc);
      const ok = setLink(doc, [seg], intent.url);
      if (ok) assignFreshTracked(ids, doc, before, intent.nodeIds);
      return ok;
    }
    case "insertFootnote": {
      const runEl = ids.elOf(intent.runId);
      if (!runEl) return false;
      const entry = runMap.get(runEl);
      if (!entry || !entry.firstT) return false;
      const before = trackedSet(ids, doc);
      const id = insertFootnote(doc, entry.firstT, entry.firstT.text.length, intent.text);
      if (id === null) return false;
      assignFreshTracked(ids, doc, before, intent.nodeIds);
      return true;
    }
    case "setDropCap": {
      const blockEl = ids.elOf(intent.blockId);
      if (!blockEl) return false;
      const target = firstTextDescendant(blockEl) ?? blockEl;
      const before = trackedSet(ids, doc);
      const ok = setDropCapAt(doc, target, intent.mode as never);
      if (ok) assignFreshTracked(ids, doc, before, intent.nodeIds);
      return ok;
    }
    case "setDivider": {
      const blockEl = ids.elOf(intent.blockId);
      if (!blockEl) return false;
      const target = firstTextDescendant(blockEl) ?? blockEl;
      return setParagraphDivider(doc, [target], intent.divider as never);
    }
    case "insertBookmark": {
      const runEl = ids.elOf(intent.runId);
      if (!runEl) return false;
      const entry = runMap.get(runEl);
      if (!entry || !entry.firstT) return false;
      return insertBookmarkAt(doc, entry.firstT, entry.firstT.text.length, intent.name);
    }
    case "adjustIndent": {
      const blockEl = ids.elOf(intent.blockId);
      if (!blockEl) return false;
      const target = firstTextDescendant(blockEl) ?? blockEl;
      return adjustIndent(doc, [target], intent.direction);
    }
    case "setSpacing": {
      const blockEl = ids.elOf(intent.blockId);
      if (!blockEl) return false;
      const target = firstTextDescendant(blockEl) ?? blockEl;
      return setParagraphSpacing(doc, [target], intent.patch as never);
    }
    case "insertPageField": {
      const runEl = ids.elOf(intent.runId);
      if (!runEl) return false;
      const entry = runMap.get(runEl);
      if (!entry || !entry.firstT) return false;
      const before = trackedSet(ids, doc);
      const ok = insertPageField(doc, entry.firstT, entry.firstT.text.length, intent.fieldKind);
      if (!ok) return false;
      assignFreshTracked(ids, doc, before, intent.nodeIds);
      return true;
    }
    case "replyComment": {
      const prov = recordedProvenance({ dates: [intent.date], paraIds: intent.paraIds });
      const ok = replyToComment(doc, intent.parentId, intent.text, intent.author, intent.initials, prov);
      if (ok && doc.stableIds) doc.stableIds.assignFromRoots(doc.editableRoots());
      return ok;
    }
    case "insertBreak": {
      const runEl = ids.elOf(intent.runId);
      if (!runEl) return false;
      const entry = runMap.get(runEl);
      if (!entry || !entry.firstT) return false;
      // At END of the run: inserts sibling break + tail runs, no text split.
      const before = trackedSet(ids, doc);
      const res = insertBreakAt(doc, entry.firstT, entry.firstT.text.length, intent.breakKind);
      if (!res) return false;
      const fresh: XmlElement[] = [];
      walkTracked(doc, (el) => { if (!before.has(el)) fresh.push(el); });
      for (let k = 0; k < fresh.length && k < intent.nodeIds.length; k++) {
        ids.reassign(fresh[k], intent.nodeIds[k]);
      }
      return true;
    }
    case "insertMath": {
      const runEl = ids.elOf(intent.runId);
      if (!runEl) return false;
      const entry = runMap.get(runEl);
      if (!entry || !entry.firstT) return false;
      const before = trackedSet(ids, doc);
      const res = insertMathAt(doc, entry.firstT, entry.firstT.text.length, intent.mathText);
      if (!res) return false;
      assignFreshTracked(ids, doc, before, intent.nodeIds);
      return true;
    }
    case "insertShape": {
      const runEl = ids.elOf(intent.runId);
      if (!runEl) return false;
      const entry = runMap.get(runEl);
      if (!entry || !entry.firstT) return false;
      const before = trackedSet(ids, doc);
      const res = insertShapeAt(doc, entry.firstT, intent.preset, intent.text ?? "");
      if (!res) return false;
      assignFreshTracked(ids, doc, before, intent.nodeIds);
      return true;
    }
  }
}

/** Assign carried ids to tracked nodes created since `before`, in doc order. */
function assignFreshTracked(ids: StableIds, doc: DocxDocument, before: Set<XmlElement>, nodeIds: number[]): void {
  const fresh: XmlElement[] = [];
  walkTracked(doc, (el) => { if (!before.has(el)) fresh.push(el); });
  for (let k = 0; k < fresh.length && k < nodeIds.length; k++) ids.reassign(fresh[k], nodeIds[k]);
}

function base64ToBytes(b64: string): Uint8Array {
  const bin = typeof atob === "function" ? atob(b64) : Buffer.from(b64, "base64").toString("binary");
  const out = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
  return out;
}

function localRun(el: XmlElement): boolean {
  const n = el.name;
  return n === "w:r" || n.endsWith(":r") || n === "r";
}

/** Walk all id-tracked elements (p / tbl / r) reachable from editable roots. */
function walkTracked(doc: DocxDocument, visit: (el: XmlElement) => void): void {
  const walk = (el: XmlElement): void => {
    const ln = localName(el.name);
    if (ln === "p" || ln === "tbl" || ln === "r") visit(el);
    el.children.forEach(walk);
  };
  doc.editableRoots().forEach(walk);
}
function trackedSet(_ids: StableIds, doc: DocxDocument): Set<XmlElement> {
  const set = new Set<XmlElement>();
  walkTracked(doc, (el) => set.add(el));
  return set;
}

function firstTextDescendant(el: XmlElement): XmlElement | null {
  if (el.name === "w:t" || el.name.endsWith(":t")) return el;
  for (const c of el.children) {
    const found = firstTextDescendant(c);
    if (found) return found;
  }
  return null;
}

/** The w:drawing element inside a run's subtree (drawing-edit intents address
 * a drawing via the run that carries it — drawings aren't tracked/id'd). */
function firstDrawingIn(el: XmlElement): XmlElement | null {
  if (localName(el.name) === "drawing") return el;
  for (const c of el.children) {
    const found = firstDrawingIn(c);
    if (found) return found;
  }
  return null;
}

/** The m:oMath element inside a run's subtree (math-edit intents address a math
 * object via the run that carries it). */
function firstMathIn(el: XmlElement): XmlElement | null {
  if (localName(el.name) === "oMath") return el;
  for (const c of el.children) {
    const found = firstMathIn(c);
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
