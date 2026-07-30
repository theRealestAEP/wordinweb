import {
  DocxDocument,
  DocxToolbar,
  DocxView,
  acceptAllRevisions,
  acceptRevision,
  addComment,
  adjustIndent,
  applyDeleteRange,
  applyInsertText,
  applyRunFormat,
  applySplitParagraph,
  applyTableOp,
  checkboxStateElement,
  collectRevisions,
  deleteComment,
  deleteMath,
  deleteSuggestedRange,
  insertBlankPageAt,
  insertBookmarkAroundSelection,
  insertBookmarkAt,
  insertBreakAt,
  insertChartAt,
  insertCoverPage,
  insertCrossReference,
  insertDateTimeField,
  insertField,
  insertFootnote,
  insertImageAt,
  insertMathAt,
  insertPageField,
  insertSectionBreak,
  insertShapeAt,
  insertSmartArtAt,
  insertTableAfter,
  insertWordArtAt,
  isSafeUrl,
  localName,
  markParagraphGlyph,
  mergeParagraphBackward,
  moveMath,
  moveTableTo,
  parseXml,
  recordedProvenance,
  rejectRevision,
  removeDrawingRun,
  removeLink,
  replyToComment,
  resizeDrawing,
  resizeTableColumn,
  resizeTableRow,
  resolveRunOffset,
  runWireLength,
  serializeXml,
  setChartData,
  setDrawingFill,
  setDrawingLineStyle,
  setDrawingOrder,
  setDrawingRotation,
  setDrawingWordArtText,
  setDropCapAt,
  setFloatingPagePosition,
  setImageAltText,
  setImageWrap,
  setLineNumbering,
  setLink,
  setListLevel,
  setListType,
  setMathLinear,
  setPageLayout,
  setParagraphAlignment,
  setParagraphDivider,
  setParagraphSpacing,
  setParagraphStyle,
  setSmartArtData,
  setSmartArtFill,
  setSmartArtNodeText,
  setSmartArtTextFormat,
  toggleCheckbox,
  validatePastedOoxml
} from "./chunk-H7WDR6FH.js";
import "./chunk-7D4SUZUM.js";

// src/collab.tsx
import { useEffect, useMemo, useRef, useState, createElement } from "react";

// ../collab/src/intents.ts
function idempotencyKey(i) {
  return `${i.clientId}:${i.clientSeq}`;
}

// ../collab/src/transform.ts
function runEditsOf(intent) {
  switch (intent.kind) {
    case "insertText":
      return [{ runId: intent.at.runId, at: intent.at.offset, del: 0, ins: intent.text.length }];
    case "deleteText":
      return [{ runId: intent.runId, at: intent.start, del: intent.end - intent.start, ins: 0 }];
    case "splitParagraph":
      return [{ runId: intent.at.runId, at: intent.at.offset, del: 0, ins: 0, movedToRunId: intent.newRunId }];
    case "formatRun":
      return [];
    case "formatParagraph":
      return [];
    case "setListType":
      return [];
    case "formatRange":
      return [{
        runId: intent.runId,
        at: intent.start,
        del: 0,
        ins: 0,
        formatSplit: { start: intent.start, end: intent.end, beforeId: intent.beforeId, middleId: intent.middleId, afterId: intent.afterId }
      }];
    case "tableOp":
      return [];
    case "mergeParagraph":
      return [];
    case "commentRun":
      return [];
    case "pasteBlocks":
      return [];
    case "suggestRevision":
      return [];
    case "insertImage":
      return [];
    case "insertBreak":
      return [];
    case "insertMath":
    case "insertShape":
      return [];
    case "replyComment":
      return [];
    case "adjustIndent":
    case "setSpacing":
      return [];
    case "insertPageField":
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
      return [];
    case "moveMath":
      return [];
    case "insertTable":
    case "acceptRevision":
    case "rejectRevision":
    case "acceptAllRevisions":
      return [];
  }
}
function transformPositionAgainstEdit(pos, edit, leftGravity = false) {
  if (pos.runId !== edit.runId) return pos;
  if (edit.formatSplit) {
    const { start, end, beforeId, middleId, afterId } = edit.formatSplit;
    if (pos.offset < start && beforeId !== void 0) {
      return { blockId: pos.blockId, runId: beforeId, offset: pos.offset };
    }
    if (pos.offset <= end) {
      return { blockId: pos.blockId, runId: middleId, offset: pos.offset - start };
    }
    if (afterId !== void 0) {
      return { blockId: pos.blockId, runId: afterId, offset: pos.offset - end };
    }
    return { blockId: pos.blockId, runId: middleId, offset: pos.offset - start };
  }
  if (edit.movedToRunId !== void 0) {
    if (pos.offset > edit.at) {
      return { blockId: pos.blockId, runId: edit.movedToRunId, offset: pos.offset - edit.at };
    }
    return pos;
  }
  const delEnd = edit.at + edit.del;
  if (edit.del === 0 && edit.at === pos.offset) {
    return leftGravity ? pos : { ...pos, offset: pos.offset + edit.ins };
  }
  if (delEnd <= pos.offset) {
    return { ...pos, offset: pos.offset + edit.ins - edit.del };
  }
  if (edit.at >= pos.offset) {
    return pos;
  }
  return { ...pos, offset: edit.at + edit.ins };
}
function transformPosition(pos, ahead, leftGravity = false) {
  let out = pos;
  for (const intent of ahead) {
    for (const edit of runEditsOf(intent)) {
      out = transformPositionAgainstEdit(out, edit, leftGravity);
    }
  }
  return out;
}
function transformIntent(intent, ahead) {
  const newBase = intent.base + ahead.length;
  switch (intent.kind) {
    case "insertText":
      return { ...intent, at: transformPosition(intent.at, ahead), base: newBase };
    case "splitParagraph":
      return { ...intent, at: transformPosition(intent.at, ahead), base: newBase };
    case "formatRun":
      return { ...intent, base: newBase };
    case "formatParagraph":
      return { ...intent, base: newBase };
    case "setListType":
      return { ...intent, base: newBase };
    case "formatRange": {
      const s = transformPosition({ blockId: intent.blockId, runId: intent.runId, offset: intent.start }, ahead);
      const e = transformPosition({ blockId: intent.blockId, runId: intent.runId, offset: intent.end }, ahead);
      if (e.runId !== s.runId) return { ...intent, runId: s.runId, blockId: s.blockId, start: s.offset, end: s.offset, base: newBase };
      return { ...intent, runId: s.runId, blockId: s.blockId, start: s.offset, end: e.offset, base: newBase };
    }
    case "tableOp":
      return { ...intent, base: newBase };
    case "mergeParagraph":
      return { ...intent, base: newBase };
    case "commentRun":
      return { ...intent, base: newBase };
    case "pasteBlocks":
      return { ...intent, base: newBase };
    case "suggestRevision": {
      if (!intent.ranges?.length) return { ...intent, base: newBase };
      const ranges = intent.ranges.map((r) => {
        const s2 = transformPosition({ blockId: r.blockId, runId: r.runId, offset: r.start }, ahead);
        const e2 = transformPosition({ blockId: r.blockId, runId: r.runId, offset: r.end }, ahead);
        if (s2.runId !== e2.runId || e2.offset <= s2.offset) return { ...r, start: 0, end: 0 };
        return { blockId: s2.blockId, runId: s2.runId, start: s2.offset, end: e2.offset };
      }).filter((r) => r.end > r.start);
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
      return { ...intent, base: newBase };
    case "deleteText": {
      const s = transformPosition({ blockId: intent.blockId, runId: intent.runId, offset: intent.start }, ahead);
      const e = transformPosition({ blockId: intent.blockId, runId: intent.runId, offset: intent.end }, ahead, true);
      return { ...intent, runId: s.runId, blockId: s.blockId, start: s.offset, end: e.runId === s.runId ? e.offset : s.offset, base: newBase };
    }
  }
}

// ../collab/src/apply.ts
var DOC_SCOPE = { kind: "doc" };
function unionScopes(a, b) {
  if (!a) return b;
  if (a.kind === "doc" || b.kind === "doc") return DOC_SCOPE;
  if (a.kind === "split" && b.kind === "split") {
    return a.before === b.before && a.after === b.after ? a : DOC_SCOPE;
  }
  if (a.kind === "split" || b.kind === "split") {
    const split = a.kind === "split" ? a : b;
    const block = a.kind === "block" ? a : b;
    return block.blocks.every((el) => el === split.before || el === split.after) ? split : DOC_SCOPE;
  }
  const blocks = [...a.blocks];
  for (const el of b.blocks) if (!blocks.includes(el)) blocks.push(el);
  return { kind: "block", blocks };
}
function applyIntentScoped(doc, ids, intent) {
  const out = { scope: DOC_SCOPE };
  try {
    const applied = applyIntentInner(doc, ids, intent, out);
    return { ...out.scope, applied };
  } catch {
    return { kind: "doc", applied: false };
  }
}
function resyncScope(doc, ids, scope) {
  if (scope.kind === "block") {
    let ok = true;
    for (const block of scope.blocks) ok = doc.reparseBodyParagraph(block) !== null && ok;
    if (ok) {
      ids.assignFromSubtrees(scope.blocks);
      return;
    }
  } else if (scope.kind === "split") {
    if (doc.reparseDirectBodyParagraphSplit(scope.before, scope.after)) {
      ids.assignFromSubtrees([scope.before, scope.after]);
      return;
    }
  }
  doc.refresh();
  ids.assignFromRoots(doc.editableRoots());
}
function applyIntentInner(doc, ids, intent, out) {
  const runOf = makeRunLookup(doc, ids);
  const ctx = { suggesting: false, revMeta: () => {
    throw new Error("suggesting mode unsupported headlessly");
  } };
  switch (intent.kind) {
    case "insertText": {
      const caret = resolveCaret(ids, runOf, intent.at);
      if (!caret) return false;
      if (intent.suggest) {
        const s = intent.suggest;
        const suggestCtx = { suggesting: true, revMeta: () => ({ author: s.author, date: s.date, nextId: () => doc.nextRevisionId() }) };
        const scope = blockScope(ids, intent.at.blockId, intent.at.runId);
        applyInsertText(doc, caret, intent.text, suggestCtx);
        if (doc.stableIds) resyncScope(doc, ids, scope);
        out.scope = scope;
        return true;
      }
      applyInsertText(doc, caret, intent.text, ctx);
      out.scope = blockScope(ids, intent.at.blockId, intent.at.runId);
      return true;
    }
    case "deleteText": {
      if (intent.end <= intent.start) return false;
      const caret = resolveCaret(ids, runOf, { blockId: intent.blockId, runId: intent.runId, offset: intent.start });
      if (!caret) return false;
      const localEnd = caret.offset + (intent.end - intent.start);
      if (localEnd > caret.t.text.length) return false;
      applyDeleteRange(caret, caret.offset, localEnd);
      out.scope = blockScope(ids, intent.blockId, intent.runId);
      return true;
    }
    case "splitParagraph": {
      const caret = resolveCaret(ids, runOf, intent.at);
      if (!caret) return false;
      const res = applySplitParagraph(doc, caret, ctx);
      if (!res) return false;
      ids.assign(res.after, intent.newBlockId);
      const newRun = res.after.children.find((c) => localRun(c));
      if (newRun) ids.assign(newRun, intent.newRunId);
      out.scope = { kind: "split", before: res.before, after: res.after };
      return true;
    }
    case "formatRun": {
      const runEl = ids.elOf(intent.runId);
      if (!runEl) return false;
      const entry = runOf(runEl);
      if (!entry) return false;
      const seg = { run: entry.run, t: null, start: 0, end: 0, props: entry.run.props };
      applyRunFormat(doc, [seg], intent.patch);
      return true;
    }
    case "formatParagraph": {
      const blockEl = ids.elOf(intent.blockId);
      if (!blockEl) return false;
      const target = firstTextDescendant(blockEl) ?? blockEl;
      let changed = false;
      if (intent.align) changed = setParagraphAlignment(doc, [target], intent.align) || changed;
      if (intent.styleId !== void 0) changed = setParagraphStyle(doc, [target], intent.styleId) || changed;
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
      const entry = runOf(runEl);
      if (!entry || !entry.firstT) return false;
      const len = runWireLength(runEl);
      if (intent.start < 0 || intent.end > len) return false;
      const hit = resolveRunOffset(runEl, intent.start);
      if (!hit) return false;
      const localEnd = hit.offset + (intent.end - intent.start);
      if (localEnd > hit.t.text.length) return false;
      const seg = { run: entry.run, t: hit.t, start: hit.offset, end: localEnd, props: entry.run.props };
      const formatted = applyRunFormat(doc, [seg], intent.patch);
      if (formatted.length === 0) return false;
      const middleT = formatted[0].t;
      const middleRun = doc.findParentOf(middleT);
      if (!middleRun) return true;
      const container = doc.findParentOf(middleRun);
      if (!container) return true;
      const mIdx = container.children.indexOf(middleRun);
      ids.reassign(middleRun, intent.middleId);
      if (intent.start > 0 && intent.beforeId !== void 0) {
        const before = container.children[mIdx - 1];
        if (before && localRun(before)) ids.reassign(before, intent.beforeId);
      }
      if (intent.end < len && intent.afterId !== void 0) {
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
      const before = isInsert ? trackedSet(ids, doc) : null;
      const ok = applyTableOp(doc, target, intent.op);
      if (!ok) return false;
      if (isInsert && before && intent.nodeIds) {
        const fresh = [];
        walkTracked(doc, (el) => {
          if (!before.has(el)) fresh.push(el);
        });
        for (let k = 0; k < fresh.length && k < intent.nodeIds.length; k++) {
          ids.reassign(fresh[k], intent.nodeIds[k]);
        }
      } else {
        ids.prune(doc.editableRoots());
      }
      return true;
    }
    case "mergeParagraph": {
      const pEl = ids.elOf(intent.blockId);
      if (!pEl) return false;
      const ok = mergeParagraphBackward(doc, pEl);
      if (ok) ids.prune(doc.editableRoots());
      return ok;
    }
    case "commentRun": {
      const runEl = ids.elOf(intent.runId);
      if (!runEl) return false;
      const entry = runOf(runEl);
      if (!entry) return false;
      const seg = { run: entry.run, t: null, start: 0, end: 0, props: entry.run.props };
      const prov = recordedProvenance({ dates: [intent.date], paraIds: [intent.paraId] });
      const ok = addComment(doc, [seg], intent.text, intent.author, intent.initials, prov);
      if (ok && doc.stableIds) doc.stableIds.assignFromRoots(doc.editableRoots());
      return ok;
    }
    case "pasteBlocks": {
      const afterEl = ids.elOf(intent.afterBlockId);
      if (!afterEl) return false;
      const parent = doc.findParentOf(afterEl);
      if (!parent) return false;
      let blocks;
      try {
        const root = parseXml(`<w:root xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${intent.blocksXml}</w:root>`);
        blocks = root.children.filter((c) => localName(c.name) === "p" || localName(c.name) === "tbl");
      } catch {
        return false;
      }
      if (blocks.length === 0) return false;
      const check = validatePastedOoxml(blocks);
      if (!check.ok) return false;
      const at = parent.children.indexOf(afterEl);
      parent.children.splice(at + 1, 0, ...blocks);
      doc.refresh();
      const newTracked = [];
      const walk = (el) => {
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
    case "suggestRevision": {
      const meta = {
        author: intent.suggest.author,
        date: intent.suggest.date,
        nextId: () => doc.nextRevisionId()
      };
      const resolved = [];
      for (const r of intent.ranges ?? []) {
        const caret = resolveCaret(ids, runOf, { blockId: r.blockId, runId: r.runId, offset: r.start });
        if (!caret) continue;
        const localEnd = caret.offset + (r.end - r.start);
        if (localEnd > caret.t.text.length) continue;
        const sc = blockScope(ids, r.blockId, r.runId);
        resolved.push({ t: caret.t, start: caret.offset, end: localEnd, block: sc.kind === "block" ? sc.blocks[0] : null });
      }
      const markEls = [];
      for (const m of intent.marks ?? []) {
        const el = ids.elOf(m.blockId);
        if (el) markEls.push({ el, glyph: m.glyph, block: localName(el.name) === "p" ? el : null });
      }
      if (resolved.length === 0 && markEls.length === 0) return false;
      if (resolved.length) deleteSuggestedRange(doc, resolved, meta);
      for (const m of markEls) markParagraphGlyph(m.el, m.glyph, meta);
      const parts = [...resolved.map((x) => x.block), ...markEls.map((x) => x.block)];
      if (parts.every((b) => b !== null)) {
        const blocks = [];
        for (const b of parts) if (!blocks.includes(b)) blocks.push(b);
        const scope = { kind: "block", blocks };
        resyncScope(doc, ids, scope);
        out.scope = scope;
      } else {
        doc.refresh();
        ids.assignFromRoots(doc.editableRoots());
      }
      return true;
    }
    case "insertImage": {
      const runEl = ids.elOf(intent.runId);
      if (!runEl) return false;
      const entry = runOf(runEl);
      if (!entry || !entry.firstT) return false;
      const relId = doc.registerPendingImage(intent.blobSha, intent.ext, { iv: intent.iv });
      const before = trackedSet(ids, doc);
      const newRun = insertImageAt(doc, entry.firstT, relId, intent.widthPx, intent.heightPx);
      if (!newRun) return false;
      doc.refresh();
      const fresh = [];
      walkTracked(doc, (el) => {
        if (!before.has(el)) fresh.push(el);
      });
      for (let k = 0; k < fresh.length && k < intent.nodeIds.length; k++) {
        ids.reassign(fresh[k], intent.nodeIds[k]);
      }
      return true;
    }
    case "insertBlankPage": {
      const runEl = ids.elOf(intent.runId);
      if (!runEl) return false;
      const entry = runOf(runEl);
      if (!entry || !entry.firstT) return false;
      const before = trackedSet(ids, doc);
      const ok = insertBlankPageAt(doc, entry.firstT, entry.firstT.text.length);
      if (ok) assignFreshTracked(ids, doc, before, intent.nodeIds);
      return ok;
    }
    case "insertSectionBreak": {
      const runEl = ids.elOf(intent.runId);
      if (!runEl) return false;
      const entry = runOf(runEl);
      if (!entry || !entry.firstT) return false;
      const before = trackedSet(ids, doc);
      const ok = insertSectionBreak(doc, entry.firstT, intent.breakType);
      if (ok) assignFreshTracked(ids, doc, before, intent.nodeIds);
      return ok;
    }
    case "insertCrossRef": {
      const runEl = ids.elOf(intent.runId);
      if (!runEl) return false;
      const entry = runOf(runEl);
      if (!entry || !entry.firstT) return false;
      const before = trackedSet(ids, doc);
      const ok = insertCrossReference(doc, entry.firstT, entry.firstT.text.length, intent.bookmark, intent.refKind);
      if (ok) assignFreshTracked(ids, doc, before, intent.nodeIds);
      return ok;
    }
    case "insertCoverPage": {
      const before = trackedSet(ids, doc);
      const ok = insertCoverPage(doc, intent.content);
      if (ok) assignFreshTracked(ids, doc, before, intent.nodeIds);
      return ok;
    }
    case "setPageLayout":
      return setPageLayout(doc, intent.patch);
    case "setListLevel": {
      const blockEl = ids.elOf(intent.blockId);
      if (!blockEl) return false;
      const target = firstTextDescendant(blockEl) ?? blockEl;
      return setListLevel(doc, [target], intent.delta);
    }
    case "insertWordArt": {
      const runEl = ids.elOf(intent.runId);
      if (!runEl) return false;
      const entry = runOf(runEl);
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
      const entry = runOf(runEl);
      if (!entry || !entry.firstT) return false;
      const before = trackedSet(ids, doc);
      const drawing = insertChartAt(doc, entry.firstT, intent.chart);
      if (!drawing) return false;
      assignFreshTracked(ids, doc, before, intent.nodeIds);
      return true;
    }
    case "insertSmartArt": {
      const runEl = ids.elOf(intent.runId);
      if (!runEl) return false;
      const entry = runOf(runEl);
      if (!entry || !entry.firstT) return false;
      const before = trackedSet(ids, doc);
      const drawing = insertSmartArtAt(doc, entry.firstT, intent.smartArt);
      if (!drawing) return false;
      assignFreshTracked(ids, doc, before, intent.nodeIds);
      return true;
    }
    case "setLineNumbering":
      return setLineNumbering(doc, intent.patch);
    case "insertDateTimeField": {
      const runEl = ids.elOf(intent.runId);
      if (!runEl) return false;
      const entry = runOf(runEl);
      if (!entry || !entry.firstT) return false;
      const before = trackedSet(ids, doc);
      const ok = insertDateTimeField(doc, entry.firstT, entry.firstT.text.length, intent.dtKind, intent.picture);
      if (ok) assignFreshTracked(ids, doc, before, intent.nodeIds);
      return ok;
    }
    case "insertField": {
      const runEl = ids.elOf(intent.runId);
      if (!runEl) return false;
      const entry = runOf(runEl);
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
      return setChartData(doc, drawing, intent.chart);
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
      const entry = runOf(runEl);
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
      return setSmartArtData(doc, drawing, intent.smartArt);
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
      return setSmartArtTextFormat(doc, drawing, intent.format, intent.nodeIndex);
    }
    case "setFloatingPagePosition": {
      const runEl = ids.elOf(intent.runId);
      if (!runEl) return false;
      const drawing = firstDrawingIn(runEl);
      if (!drawing) return false;
      return setFloatingPagePosition(doc, drawing, intent.xPx, intent.yPx);
    }
    case "resizeDrawing": {
      const runEl = ids.elOf(intent.runId);
      if (!runEl) return false;
      const drawing = firstDrawingIn(runEl);
      if (!drawing) return false;
      return resizeDrawing(doc, drawing, intent.widthPx, intent.heightPx);
    }
    case "resizeTableColumn": {
      const tbl = tableOfParagraph(doc, ids, intent.cellParagraphId);
      if (!tbl) return false;
      return resizeTableColumn(doc, tbl, intent.boundary, intent.deltaPx, intent.renderedWidths);
    }
    case "resizeTableRow": {
      const tbl = tableOfParagraph(doc, ids, intent.cellParagraphId);
      if (!tbl) return false;
      return resizeTableRow(doc, tbl, intent.rowIdx, intent.heightPx);
    }
    case "moveTable": {
      const tbl = tableOfParagraph(doc, ids, intent.cellParagraphId);
      if (!tbl) return false;
      return moveTableTo(doc, tbl, intent.xPx, intent.yPx, intent.preservePageStart, intent.pageDelta);
    }
    case "removeDrawing": {
      const runEl = ids.elOf(intent.runId);
      if (!runEl) return false;
      const drawing = firstDrawingIn(runEl);
      if (!drawing) return false;
      const ok = removeDrawingRun(doc, drawing);
      if (ok) ids.prune(doc.editableRoots());
      return ok;
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
    case "moveMath": {
      const blockEl = ids.elOf(intent.blockId);
      if (!blockEl) return false;
      const math = firstMathIn(blockEl);
      if (!math) return false;
      const dest = resolveCaret(ids, runOf, intent.at);
      if (!dest) return false;
      const before = trackedSet(ids, doc);
      if (!moveMath(doc, math, dest.t, dest.offset)) return false;
      const fresh = [];
      walkTracked(doc, (el) => {
        if (!before.has(el)) fresh.push(el);
      });
      for (let k = 0; k < fresh.length && k < intent.nodeIds.length; k++) ids.reassign(fresh[k], intent.nodeIds[k]);
      return true;
    }
    case "ensureHeaderFooter": {
      if (doc.hasHfPart(intent.hfKind)) return false;
      const before = trackedSet(ids, doc);
      doc.ensureHfPart(intent.hfKind);
      const fresh = [];
      walkTracked(doc, (el) => {
        if (!before.has(el)) fresh.push(el);
      });
      for (let k = 0; k < fresh.length && k < intent.nodeIds.length; k++) ids.reassign(fresh[k], intent.nodeIds[k]);
      return true;
    }
    case "deleteComment":
      return deleteComment(doc, intent.commentId);
    case "insertBookmarkRange": {
      const runEl = ids.elOf(intent.runId);
      if (!runEl) return false;
      const entry = runOf(runEl);
      if (!entry || !entry.firstT) return false;
      if (intent.end > runWireLength(runEl)) return false;
      const hit = resolveRunOffset(runEl, intent.start);
      if (!hit) return false;
      const localEnd = hit.offset + (intent.end - intent.start);
      if (localEnd > hit.t.text.length) return false;
      const seg = { run: entry.run, t: hit.t, start: hit.offset, end: localEnd, props: entry.run.props };
      return insertBookmarkAroundSelection(doc, [seg], intent.name);
    }
    case "toggleCheckbox": {
      const runEl = ids.elOf(intent.runId);
      if (!runEl) return false;
      const entry = runOf(runEl);
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
    case "insertTable": {
      const runEl = ids.elOf(intent.runId);
      if (!runEl) return false;
      const entry = runOf(runEl);
      if (!entry || !entry.firstT) return false;
      const before = trackedSet(ids, doc);
      const ok = insertTableAfter(doc, entry.firstT, intent.rows, intent.cols);
      if (ok) assignFreshTracked(ids, doc, before, intent.nodeIds);
      return ok;
    }
    case "setLink": {
      if (!isSafeUrl(intent.url)) return false;
      const runEl = ids.elOf(intent.runId);
      if (!runEl) return false;
      const entry = runOf(runEl);
      if (!entry || !entry.firstT) return false;
      const seg = { run: entry.run, t: entry.firstT, start: 0, end: entry.firstT.text.length, props: entry.run.props };
      const before = trackedSet(ids, doc);
      const ok = setLink(doc, [seg], intent.url);
      if (ok) assignFreshTracked(ids, doc, before, intent.nodeIds);
      return ok;
    }
    case "insertFootnote": {
      const runEl = ids.elOf(intent.runId);
      if (!runEl) return false;
      const entry = runOf(runEl);
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
      const ok = setDropCapAt(doc, target, intent.mode);
      if (ok) assignFreshTracked(ids, doc, before, intent.nodeIds);
      return ok;
    }
    case "setDivider": {
      const blockEl = ids.elOf(intent.blockId);
      if (!blockEl) return false;
      const target = firstTextDescendant(blockEl) ?? blockEl;
      return setParagraphDivider(doc, [target], intent.divider);
    }
    case "insertBookmark": {
      const runEl = ids.elOf(intent.runId);
      if (!runEl) return false;
      const entry = runOf(runEl);
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
      return setParagraphSpacing(doc, [target], intent.patch);
    }
    case "insertPageField": {
      const runEl = ids.elOf(intent.runId);
      if (!runEl) return false;
      const entry = runOf(runEl);
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
      const entry = runOf(runEl);
      if (!entry || !entry.firstT) return false;
      const before = trackedSet(ids, doc);
      const res = insertBreakAt(doc, entry.firstT, entry.firstT.text.length, intent.breakKind);
      if (!res) return false;
      const fresh = [];
      walkTracked(doc, (el) => {
        if (!before.has(el)) fresh.push(el);
      });
      for (let k = 0; k < fresh.length && k < intent.nodeIds.length; k++) {
        ids.reassign(fresh[k], intent.nodeIds[k]);
      }
      return true;
    }
    case "insertMath": {
      const runEl = ids.elOf(intent.runId);
      if (!runEl) return false;
      const entry = runOf(runEl);
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
      const entry = runOf(runEl);
      if (!entry || !entry.firstT) return false;
      const before = trackedSet(ids, doc);
      const res = insertShapeAt(doc, entry.firstT, intent.preset, intent.text ?? "");
      if (!res) return false;
      assignFreshTracked(ids, doc, before, intent.nodeIds);
      return true;
    }
  }
}
function assignFreshTracked(ids, doc, before, nodeIds) {
  const fresh = [];
  walkTracked(doc, (el) => {
    if (!before.has(el)) fresh.push(el);
  });
  for (let k = 0; k < fresh.length && k < nodeIds.length; k++) ids.reassign(fresh[k], nodeIds[k]);
}
function localRun(el) {
  const n = el.name;
  return n === "w:r" || n.endsWith(":r") || n === "r";
}
function walkTracked(doc, visit) {
  const walk = (el) => {
    const ln = localName(el.name);
    if (ln === "p" || ln === "tbl" || ln === "r") visit(el);
    el.children.forEach(walk);
  };
  doc.editableRoots().forEach(walk);
}
function trackedSet(_ids, doc) {
  const set = /* @__PURE__ */ new Set();
  walkTracked(doc, (el) => set.add(el));
  return set;
}
function firstTextDescendant(el) {
  if (el.name === "w:t" || el.name.endsWith(":t")) return el;
  for (const c of el.children) {
    const found = firstTextDescendant(c);
    if (found) return found;
  }
  return null;
}
function tableOfParagraph(doc, ids, cellParagraphId) {
  const paraEl = ids.elOf(cellParagraphId);
  if (!paraEl) return null;
  for (let cur = paraEl; cur; cur = doc.findParentOf(cur) ?? null) {
    if (localName(cur.name) === "tbl") return cur;
  }
  return null;
}
function firstDrawingIn(el) {
  if (localName(el.name) === "drawing") return el;
  for (const c of el.children) {
    const found = firstDrawingIn(c);
    if (found) return found;
  }
  return null;
}
function firstMathIn(el) {
  if (localName(el.name) === "oMath") return el;
  for (const c of el.children) {
    const found = firstMathIn(c);
    if (found) return found;
  }
  return null;
}
function runEntry(run) {
  return { run, firstT: run.content.find((c) => c.kind === "text")?.srcT };
}
function makeRunLookup(doc, ids) {
  let full = null;
  const scoped = /* @__PURE__ */ new Map();
  return (runEl, blockIdHint) => {
    if (!full) {
      const cached = scoped.get(runEl);
      if (cached) return cached;
      if (blockIdHint !== void 0) {
        const blockEl = ids.elOf(blockIdHint);
        const para = blockEl ? doc.paragraphBySource(blockEl) : null;
        if (para) {
          for (const item of para.children) {
            for (const run of item.type === "run" ? [item] : item.runs) {
              if (run.src) scoped.set(run.src, runEntry(run));
            }
          }
          const hit = scoped.get(runEl);
          if (hit) return hit;
        }
      }
      full = buildRunMap(doc);
    }
    return full.get(runEl);
  };
}
function contains(el, needle) {
  if (el === needle) return true;
  for (const c of el.children) if (contains(c, needle)) return true;
  return false;
}
function blockScope(ids, blockId, runId) {
  const blockEl = ids.elOf(blockId);
  const runEl = ids.elOf(runId);
  if (!blockEl || !runEl || localName(blockEl.name) !== "p") return DOC_SCOPE;
  return contains(blockEl, runEl) ? { kind: "block", blocks: [blockEl] } : DOC_SCOPE;
}
function buildRunMap(doc) {
  const map = /* @__PURE__ */ new Map();
  const visitRun = (run) => {
    if (run.src) map.set(run.src, runEntry(run));
    for (const c of run.content) {
      if (c.kind === "anchor" && c.shape.type === "textbox") visitBlocks(c.shape.blocks);
    }
  };
  const visitBlocks = (blocks) => {
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
  for (const hf of doc.headers.values()) visitBlocks(hf.blocks);
  for (const hf of doc.footers.values()) visitBlocks(hf.blocks);
  return map;
}
function resolveCaret(ids, runOf, pos) {
  const runEl = ids.elOf(pos.runId);
  if (!runEl) return null;
  const entry = runOf(runEl, pos.blockId);
  if (!entry || !entry.firstT) return null;
  if (pos.offset < 0 || pos.offset > runWireLength(runEl)) return null;
  const hit = resolveRunOffset(runEl, pos.offset);
  if (!hit) return null;
  return { t: hit.t, run: entry.run, offset: hit.offset };
}

// ../collab/src/replica.ts
var _ClientReplica = class _ClientReplica {
  constructor(bytes, sidecar) {
    /** Server seq the confirmed baseline reflects. */
    this.confirmedSeq = 0;
    this.pending = [];
    /** Confirmed entries NOT yet folded into confirmedBytes: own-echoes
     * deferred mid-burst (the original role) AND — since the lazy-baseline
     * change — every fast-path confirmed entry, so the O(doc) save() stays
     * off the receive hot path entirely. restoreConfirmed replays the tail
     * onto the baseline; exportBundleState folds it for persistence; the
     * fast path folds eagerly once the tail reaches FOLD_TAIL_AT (bounding
     * conflict-replay cost while amortizing saves to ~1/FOLD_TAIL_AT ops). */
    this.confirmedTail = [];
    /**
     * Bumped every time THIS CLASS mutates the live doc (optimistic apply,
     * remote apply, replay, restore) — and deliberately NOT on bookkeeping
     * that leaves the doc untouched (trackLocal, own-echo confirmation).
     *
     * This is the renderer's repaint gate. The react layer used to repaint on
     * EVERY onChange, but for the editor-driven typing path (trackLocal + the
     * own-echo fast path) the doc was already mutated AND painted by the
     * editor, so each keystroke queued a redundant whole-document relayout —
     * invisible on 5 pages, catastrophic past the background-layout threshold
     * (500 pages: the layout ran async with the container inert until it
     * finished, blocking further typing for seconds per keystroke).
     */
    this.docVersion = 0;
    /**
     * Union of the dirty scopes behind the docVersion bumps not yet consumed by
     * a repaint (see takeRenderScope). Every path that bumps docVersion also
     * records what it disturbed: applyAndResync its intent's ApplyScope,
     * restoreConfirmed doc scope (the doc OBJECT was replaced). The renderer
     * drains this when it repaints, so a coalesced repaint covering several
     * bumps sees the union — the batching rule lives in unionScopes.
     */
    this.renderScope = null;
    /**
     * Receive the server's canonical broadcast entries (in seq order) and
     * reconcile. The invariant is: live doc = confirmed baseline + remaining
     * pending. So every receive restores to the confirmed baseline, applies the
     * canonical batch (advancing the baseline and dropping matched pending),
     * then replays the still-pending intents transformed against the remote
     * intents in the batch (plan doc 03). Idempotent for already-seen seqs.
     */
    /** True after a receive() that reloaded the doc object (a true conflict
     * reconciliation) — the renderer must re-key/rebind. Stays false for the
     * in-place fast paths so the editor updates without a re-mount. */
    this.reloaded = false;
    this.doc = DocxDocument.load(bytes);
    this.ids = this.doc.enableStableIds();
    if (sidecar) this.ids.importSidecar(this.doc.editableRoots(), sidecar);
    this.confirmedBytes = this.doc.save();
    this.confirmedSidecar = this.ids.exportSidecar(this.doc.editableRoots());
  }
  /** Drain the dirty scope accumulated since the last take (null when no
   * doc-mutating change was recorded). The caller repaints exactly this. */
  takeRenderScope() {
    const s = this.renderScope;
    this.renderScope = null;
    return s;
  }
  /**
   * Apply a locally produced intent optimistically and enqueue it as pending.
   * The intent's `base` must be the current confirmed seq.
   *
   * CONSTRAINT (one-in-flight): a client should hold at most one un-confirmed
   * pending intent at a time — submit the next only after the previous is
   * confirmed. This is the model production OT servers use (e.g. ShareDB).
   * Multiple concurrent same-client pending whose offsets assume each other
   * cannot be correctly rebased against an interleaved remote edit without
   * operation inverses, which the OT-lite transform deliberately does not
   * implement (plan doc 03 marks adjusted-sibling replay a scoped next step).
   * The transform below is correct for the single-in-flight case; the replay
   * still stores transformed forms so a burst that stays same-client-only
   * (no remote interleave) also composes.
   */
  submitLocal(intent) {
    this.applyAndResync(intent);
    this.pending.push(intent);
  }
  /** Track an intent the caller ALREADY applied to this replica's live doc
   * (the editor-driven path: DocxEditor mutates `doc` through its own command
   * and emits the intent afterwards). Applying it again here would double it —
   * the demo typed "Hello" and got "Hello" + its reversal. Pending tracking
   * and reconciliation behave exactly as for submitLocal. */
  trackLocal(intent) {
    this.pending.push(intent);
  }
  /** Number of un-confirmed local intents (for the one-in-flight discipline). */
  get pendingCount() {
    return this.pending.length;
  }
  /** Identity of the OLDEST un-confirmed local intent — the delivery-retry
   * probe (a stuck front pending means the op or its echo was lost). */
  firstPending() {
    const p = this.pending[0];
    return p ? { clientId: p.clientId, clientSeq: p.clientSeq } : null;
  }
  /** Copies of every pending intent in submit order (their reconciliation-
   * transformed forms) — the rate-limit re-drive resends the WHOLE queue:
   * chained bursts (each op addressing ids a previous op allocated) strand
   * everything behind one refused op, and dedup makes resends free. */
  pendingCopies() {
    return [...this.pending];
  }
  /** The current pending copy of an intent by idempotency key, or null if it
   * was confirmed or dropped. This is the RECONCILIATION-TRANSFORMED form —
   * replayPending rewrites pending against every remote entry ingested since
   * submission — and therefore the only correct body to re-seal on a
   * stale-base retry: resealing the raw original resubmits untransformed
   * coordinates that the server then applies verbatim, while the optimistic
   * doc holds the transformed form — a silently divergent baseline once the
   * own-echo fast path snapshots it (the checkpoint-boundary divergence). */
  pendingIntent(clientId, clientSeq) {
    return this.pending.find((p) => p.clientId === clientId && p.clientSeq === clientSeq) ?? null;
  }
  /**
   * The replica's durable state — exactly the fields the doc-12 §4 bundle
   * persists (the persistence layer is a thin observer over this). The bytes
   * and sidecar are the CONFIRMED baseline, never the live optimistic doc:
   * a bundle must be replayable (confirmed + pending re-submit), and
   * snapshotting optimistic state would bake unsequenced edits into what a
   * re-seed later presents as canonical (doc 03 re-capture rule).
   * `pending` is a copy — callers can't mutate reconciliation state.
   */
  exportBundleState() {
    let confirmedBytes = this.confirmedBytes;
    let confirmedSidecar = this.confirmedSidecar;
    if (this.confirmedTail.length) {
      if (this.pending.length === 0) {
        confirmedBytes = this.doc.save();
        confirmedSidecar = this.ids.exportSidecar(this.doc.editableRoots());
      } else {
        const base = DocxDocument.load(this.confirmedBytes);
        const baseIds = base.enableStableIds();
        baseIds.importSidecar(base.editableRoots(), this.confirmedSidecar);
        for (const e of this.confirmedTail) {
          if (e.kind !== "applied") continue;
          const res = applyIntentScoped(base, baseIds, e.intent);
          if (res.applied) resyncScope(base, baseIds, res);
        }
        this.carryInstalledMedia(this.doc, base);
        confirmedBytes = base.save();
        confirmedSidecar = baseIds.exportSidecar(base.editableRoots());
      }
    }
    return {
      confirmedSeq: this.confirmedSeq,
      confirmedBytes,
      confirmedSidecar,
      pending: [...this.pending],
      mediaMeta: [...this.doc.mediaMeta]
    };
  }
  async exportBundleStateAsync() {
    let confirmedBytes = this.confirmedBytes;
    let confirmedSidecar = this.confirmedSidecar;
    if (this.confirmedTail.length) {
      if (this.pending.length === 0) {
        confirmedBytes = await this.doc.saveAsync();
        confirmedSidecar = this.ids.exportSidecar(this.doc.editableRoots());
      } else {
        const base = DocxDocument.load(this.confirmedBytes);
        const baseIds = base.enableStableIds();
        baseIds.importSidecar(base.editableRoots(), this.confirmedSidecar);
        for (const e of this.confirmedTail) {
          if (e.kind !== "applied") continue;
          const res = applyIntentScoped(base, baseIds, e.intent);
          if (res.applied) resyncScope(base, baseIds, res);
        }
        this.carryInstalledMedia(this.doc, base);
        confirmedBytes = await base.saveAsync();
        confirmedSidecar = baseIds.exportSidecar(base.editableRoots());
      }
    }
    return {
      confirmedSeq: this.confirmedSeq,
      confirmedBytes,
      confirmedSidecar,
      pending: [...this.pending],
      mediaMeta: [...this.doc.mediaMeta]
    };
  }
  receive(entries) {
    this.reloaded = false;
    const fresh = entries.filter((e) => e.seq > this.confirmedSeq);
    if (fresh.length === 0) return;
    const remoteAhead = fresh.filter((e) => e.kind === "applied" && !isOurs(e, this.pending)).map((e) => e.intent);
    const rejectedOurs = fresh.some((e) => e.kind === "rejected" && isOurs(e, this.pending));
    if (rejectedOurs) {
      this.restoreConfirmed();
      this.reloaded = true;
      for (const e of fresh) this.advanceConfirmed(e, e.kind === "applied");
      this.snapshotConfirmed();
      this.replayPending(remoteAhead);
      this.resync();
      return;
    }
    const freshKeys = new Set(
      fresh.map((e) => e.kind === "applied" ? idempotencyKey(e.intent) : `${e.clientId}:${e.clientSeq}`)
    );
    const drainsAllPending = this.pending.every((p) => freshKeys.has(idempotencyKey(p)));
    if (this.pending.length === 0 || remoteAhead.length === 0 && drainsAllPending) {
      const applyToDoc = remoteAhead.length > 0;
      for (const e of fresh) {
        this.advanceConfirmed(
          e,
          /*applyToDoc*/
          applyToDoc
        );
        if (e.kind === "applied") this.confirmedTail.push(e);
      }
      if (this.confirmedTail.length >= _ClientReplica.FOLD_TAIL_AT) this.snapshotConfirmed();
      return;
    }
    if (remoteAhead.length === 0) {
      for (const e of fresh) {
        if (e.kind === "applied") this.confirmedTail.push(e);
        const key = e.kind === "applied" ? idempotencyKey(e.intent) : `${e.clientId}:${e.clientSeq}`;
        this.pending = this.pending.filter((p) => idempotencyKey(p) !== key);
        this.confirmedSeq = e.seq;
      }
      return;
    }
    this.restoreConfirmed();
    this.reloaded = true;
    for (const e of fresh) this.advanceConfirmed(e, true);
    this.snapshotConfirmed();
    this.replayPending(remoteAhead);
    this.resync();
  }
  /**
   * Apply one intent to the live doc and reconcile ONLY what it touched
   * (perf B9/B10). Every applied intent used to cost a full model reparse plus
   * a full id walk somewhere on this path, so a keystroke arriving into a
   * 600-paragraph document cost 600 paragraphs of work — the seeding curve and
   * the stall under an incoming flood. Text-level intents now reparse their
   * own paragraph; anything structural or unverifiable still takes the full
   * refresh (see resyncScope). Returns false when the intent didn't apply.
   */
  applyAndResync(intent) {
    const res = applyIntentScoped(this.doc, this.ids, intent);
    if (!res.applied) return false;
    resyncScope(this.doc, this.ids, res);
    this.docVersion++;
    this.renderScope = unionScopes(this.renderScope, res);
    return true;
  }
  advanceConfirmed(e, applyToDoc) {
    if (e.kind === "applied" && applyToDoc) this.applyAndResync(e.intent);
    const key = e.kind === "applied" ? idempotencyKey(e.intent) : `${e.clientId}:${e.clientSeq}`;
    this.pending = this.pending.filter((p) => idempotencyKey(p) !== key);
    this.confirmedSeq = e.seq;
  }
  replayPending(remoteAhead) {
    const stillPending = this.pending;
    this.pending = [];
    for (const p of stillPending) {
      const transformed = transformIntent(p, remoteAhead);
      if (this.applyAndResync(transformed)) this.pending.push(transformed);
    }
  }
  /** Copy installed media pixels (out-of-band state — bytes arrive via the
   * relay, not the intent log) from one doc into another whose matching
   * parts are still pending. */
  carryInstalledMedia(from, to) {
    for (const [part] of from.mediaMeta) {
      if (from.pendingMedia.has(part)) continue;
      const bytes = from.pkg.binary(part);
      if (bytes) to.installMedia(part, bytes);
    }
  }
  restoreConfirmed() {
    this.docVersion++;
    this.renderScope = { kind: "doc" };
    const pixelSource = this.doc;
    this.doc = DocxDocument.load(this.confirmedBytes);
    this.ids = this.doc.enableStableIds();
    this.ids.importSidecar(this.doc.editableRoots(), this.confirmedSidecar);
    for (const e of this.confirmedTail) {
      if (e.kind === "applied") this.applyAndResync(e.intent);
    }
    this.carryInstalledMedia(pixelSource, this.doc);
  }
  snapshotConfirmed() {
    this.confirmedBytes = this.doc.save();
    this.confirmedSidecar = this.ids.exportSidecar(this.doc.editableRoots());
    this.confirmedTail = [];
  }
  resync() {
    this.doc.refresh();
    this.ids.assignFromRoots(this.doc.editableRoots());
  }
};
/** Fold cadence for the lazy confirmed baseline (see confirmedTail). */
_ClientReplica.FOLD_TAIL_AT = 100;
var ClientReplica = _ClientReplica;
function isOurs(e, pending) {
  const key = e.kind === "applied" ? idempotencyKey(e.intent) : `${e.clientId}:${e.clientSeq}`;
  return pending.some((p) => idempotencyKey(p) === key);
}

// ../collab/src/bundle.ts
function versionKey(docId, savedAt, label) {
  return `${docId}#version-${savedAt}${label ? `-${label}` : ""}`;
}
function draftKey(docId, genesisId) {
  return `${docId}#draft-${genesisId}`;
}
function supersededKey(docId, genesisId) {
  return `${docId}#superseded-${genesisId}`;
}
function parseBundleKey(key) {
  if (key.startsWith("local:")) return { docId: key, kind: "local" };
  const hash = key.indexOf("#");
  if (hash === -1) return { docId: key, kind: "live" };
  const docId = key.slice(0, hash);
  const suffix = key.slice(hash + 1);
  const version = /^version-(\d+)(?:-(.*))?$/.exec(suffix);
  if (version) return { docId, kind: "version", versionSavedAt: Number(version[1]), label: version[2] };
  if (suffix.startsWith("draft-")) return { docId, kind: "draft" };
  if (suffix.startsWith("superseded-")) return { docId, kind: "superseded" };
  return { docId, kind: "unknown" };
}
var InMemoryBundleStore = class {
  constructor() {
    this.bundles = /* @__PURE__ */ new Map();
    /** Write count — lets tests assert the throttle's coalescing precisely. */
    this.writes = 0;
  }
  async get(docId) {
    return this.bundles.get(docId) ?? null;
  }
  async put(bundle) {
    this.writes++;
    this.bundles.set(bundle.docId, bundle);
  }
  async delete(docId) {
    this.bundles.delete(docId);
  }
  async list() {
    return [...this.bundles.values()].map((b) => ({
      ...parseBundleKey(b.docId),
      key: b.docId,
      savedAt: b.savedAt,
      byteLength: b.confirmedBytes.byteLength
    }));
  }
};
var BundlePersister = class {
  constructor(conn, store, docId, opts = {}) {
    this.conn = conn;
    this.store = store;
    this.docId = docId;
    this.opts = opts;
    this.lastWrite = -Infinity;
    this.trailing = null;
    this.stopped = false;
    /** Writes are async (digest + store I/O) and must not overlap on one
     * store; a serial chain guarantees ordering and lets flush() await
     * every in-flight write (round-4 F8: the flush is only meaningful if it
     * durably lands the latest state before pagehide). */
    this.chain = Promise.resolve();
  }
  get throttleMs() {
    return this.opts.throttleMs ?? 1e3;
  }
  now() {
    return (this.opts.now ?? Date.now)();
  }
  /** State changed — write now if the window allows, else arm ONE trailing
   * write for the window's end (never more than one timer in flight; N
   * notifies inside a window coalesce into a single trailing write). */
  notify() {
    if (this.stopped) return;
    const elapsed = this.now() - this.lastWrite;
    if (elapsed >= this.throttleMs) {
      this.lastWrite = this.now();
      this.enqueueWrite();
      return;
    }
    if (this.trailing === null) {
      const set = this.opts.setTimer ?? ((fn, ms) => setTimeout(fn, ms));
      this.trailing = set(() => {
        this.trailing = null;
        this.lastWrite = this.now();
        this.enqueueWrite();
      }, this.throttleMs - elapsed);
    }
  }
  /** Immediate best-effort write (pagehide/visibilitychange/session-end).
   * Awaits ALL in-flight writes plus this one so the latest state is
   * durably landed before the tab goes away. */
  async flush() {
    if (this.stopped) return;
    if (this.trailing !== null) {
      (this.opts.clearTimer ?? clearTimeout)(this.trailing);
      this.trailing = null;
    }
    this.lastWrite = this.now();
    this.enqueueWrite();
    await this.chain;
  }
  /** Append a write to the serial chain (never overlapping). */
  enqueueWrite() {
    this.chain = this.chain.then(() => this.write()).catch((err) => {
      this.opts.onError?.(err);
    });
  }
  /** Detach (unmount): cancel the trailing timer; no further writes. */
  stop() {
    this.stopped = true;
    if (this.trailing !== null) {
      (this.opts.clearTimer ?? clearTimeout)(this.trailing);
      this.trailing = null;
    }
  }
  async write() {
    const bundle = await this.conn.exportBundleAsync(this.docId);
    if (!bundle) return;
    bundle.savedAt = this.lastWrite;
    const prior = await this.store.get(this.docId);
    const chain = [...prior?.lineage ?? []];
    const digest = await crypto.subtle.digest("SHA-256", bundle.confirmedBytes);
    let hash = "";
    for (const b of new Uint8Array(digest)) hash += b.toString(16).padStart(2, "0");
    const head = { genesisId: bundle.genesisId, seq: bundle.confirmedSeq, docHash: hash };
    if (chain.length && chain[chain.length - 1].genesisId === bundle.genesisId) chain[chain.length - 1] = head;
    else chain.push(head);
    bundle.lineage = chain.slice(-50);
    const off = this.opts.offlineTail?.();
    if (off && off.tail.length) {
      bundle.offlineTail = off.tail;
      bundle.offlineTailEpoch = off.epoch;
    }
    await this.store.put(bundle);
  }
};

// ../collab/src/rebase.ts
var OFFLINE_TAIL_CAP = 2e3;
function toSuggestions(tail, author, date) {
  const suggestions = [];
  const dropped = [];
  for (const intent of tail) {
    if (intent.kind === "insertText") {
      suggestions.push({ kind: "insertText", at: intent.at, text: intent.text, suggest: { author, date } });
    } else if (intent.kind === "deleteText") {
      suggestions.push({
        kind: "suggestRevision",
        ranges: [{ blockId: intent.blockId, runId: intent.runId, start: intent.start, end: intent.end }],
        suggest: { author, date }
      });
    } else {
      dropped.push(intent);
    }
  }
  return { suggestions, dropped };
}
function arrivalMode(tailLength, diverged, opts = {}) {
  if (!diverged) return "fast-forward";
  const threshold = opts.suggestThreshold ?? 50;
  return tailLength <= threshold ? "suggest" : "draft";
}

// ../collab/src/protocol.ts
var ENGINE_VERSION = "e6";
var PRESENCE_MAX_RANGES = 64;
function sanitizePresencePosition(pos) {
  if (!pos || !pos.ranges) return pos;
  const int = (n) => typeof n === "number" && Number.isFinite(n) && n >= 0;
  const clean = [];
  for (const r of Array.isArray(pos.ranges) ? pos.ranges : []) {
    if (clean.length >= PRESENCE_MAX_RANGES) break;
    if (!r || typeof r !== "object") continue;
    if (!int(r.blockId) || !int(r.runId) || !int(r.start) || !int(r.end)) continue;
    if (r.end <= r.start) continue;
    clean.push({ blockId: r.blockId, runId: r.runId, start: r.start, end: r.end });
  }
  if (clean.length === pos.ranges.length) return pos;
  return clean.length > 0 ? { ...pos, ranges: clean } : { ...pos, ranges: void 0 };
}
var PROTOCOL_VERSION = 3;

// ../collab/src/media.ts
function mediaAddressesOf(doc) {
  const out = [];
  for (const [part, meta] of doc.mediaMeta) {
    if (!meta.sha) continue;
    out.push({ part, sha: meta.sha, ...meta.iv ? { iv: meta.iv } : {} });
  }
  return out;
}
function applyMediaAddresses(doc, addresses) {
  if (!addresses?.length) return;
  for (const { part, sha, iv } of addresses) {
    if (!sha || doc.mediaStatus(part) === "ready") continue;
    const meta = { sha, ...iv ? { iv } : {} };
    doc.pendingMedia.set(part, meta);
    doc.mediaMeta.set(part, meta);
  }
}
async function sha256Hex(bytes) {
  const digest = await crypto.subtle.digest("SHA-256", bytes);
  let hex = "";
  for (const b of new Uint8Array(digest)) hex += b.toString(16).padStart(2, "0");
  return hex;
}
function mediaUrl(httpBase, docId, sha) {
  return `${httpBase.replace(/\/$/, "")}/docs/${encodeURIComponent(docId)}/media/${sha}`;
}
async function putBlob(opts, docId, sha, blob) {
  const f = opts.fetchImpl ?? globalThis.fetch;
  try {
    const res = await f(mediaUrl(opts.httpBase, docId, sha), { method: "PUT", body: blob });
    return { ok: res.status === 200 || res.status === 201, status: res.status };
  } catch {
    return { ok: false, status: 0 };
  }
}
async function getBlob(opts, docId, sha) {
  const f = opts.fetchImpl ?? globalThis.fetch;
  try {
    const res = await f(mediaUrl(opts.httpBase, docId, sha));
    if (!res.ok) return null;
    return new Uint8Array(await res.arrayBuffer());
  } catch {
    return null;
  }
}
var MediaClient = class {
  constructor(opts, getDoc, cb, crypto2) {
    this.opts = opts;
    this.getDoc = getDoc;
    this.cb = cb;
    this.crypto = crypto2;
    this.inFlight = /* @__PURE__ */ new Set();
    this.states = /* @__PURE__ */ new Map();
    /** Shas we asked the room for, awaiting a media-ready. */
    this.waitingFor = /* @__PURE__ */ new Set();
    /** Plaintext of blobs THIS client uploaded, so its own images never take a
     * network round trip. Bounded: an entry is dropped once installed. */
    this.ownBlobs = /* @__PURE__ */ new Map();
  }
  /** Current placeholder state of a pending part (undefined once ready). */
  stateOf(part) {
    return this.states.get(part);
  }
  /**
   * PLACER (doc 16 §5.1). Seal (E2EE) → address → upload. Returns the intent
   * fields on success, or null when the relay refused — in which case the
   * caller MUST NOT emit an insertImage, or the room gets a reservation
   * pointing at a blob that does not exist.
   */
  async upload(docId, plaintext) {
    const sealed = this.crypto ? await this.crypto.seal(plaintext) : null;
    const blob = sealed ? sealed.blob : plaintext;
    const blobSha = await sha256Hex(blob);
    const res = await putBlob(this.opts, docId, blobSha, blob);
    if (!res.ok) return null;
    this.ownBlobs.set(blobSha, plaintext);
    return { blobSha, bytesLen: blob.length, iv: sealed?.iv };
  }
  /**
   * RECEIVER (doc 16 §5.2). Eager-fetch every pending part the doc knows
   * about. Idempotent and safe to call after every applied broadcast: parts
   * already ready or already in flight are skipped.
   */
  async fetchPending(docId) {
    const doc = this.getDoc();
    if (!doc) return;
    for (const [part, meta] of [...doc.pendingMedia]) {
      if (!meta.sha || this.inFlight.has(part)) continue;
      this.inFlight.add(part);
      void this.fetchOne(docId, part, meta).finally(() => this.inFlight.delete(part));
    }
  }
  setState(part, state) {
    this.states.set(part, state);
    this.getDoc()?.mediaTransferState.set(part, state);
    this.cb.onState?.(part, state);
  }
  async fetchOne(docId, part, meta) {
    const doc = this.getDoc();
    if (!doc || doc.mediaStatus(part) === "ready") return;
    const own = this.ownBlobs.get(meta.sha);
    if (own) {
      doc.installMedia(part, own);
      this.states.delete(part);
      doc.mediaTransferState.delete(part);
      this.ownBlobs.delete(meta.sha);
      this.cb.onChange?.();
      return;
    }
    this.setState(part, "fetching");
    const blob = await getBlob(this.opts, docId, meta.sha);
    if (!blob) {
      this.setState(part, "waiting");
      this.waitingFor.add(meta.sha);
      this.cb.need(meta.sha);
      return;
    }
    await this.installVerified(part, meta, blob);
  }
  /**
   * Verify then install. The sha check is against the value COMMITTED IN THE
   * INTENT, never against anything the deliverer said — defense in depth
   * behind the relay's own check (doc 16 §5.2 step 3), and the only thing
   * standing between a compromised relay and arbitrary bytes in the document.
   */
  async installVerified(part, meta, blob) {
    if (await sha256Hex(blob) !== meta.sha) {
      this.setState(part, "waiting");
      return false;
    }
    let bytes = blob;
    if (this.crypto) {
      if (!meta.iv) return false;
      try {
        bytes = await this.crypto.open(blob, meta.iv);
      } catch {
        this.setState(part, "waiting");
        return false;
      }
    }
    const doc = this.getDoc();
    if (!doc) return false;
    doc.installMedia(part, bytes);
    this.states.delete(part);
    doc.mediaTransferState.delete(part);
    this.waitingFor.delete(meta.sha);
    this.cb.onChange?.();
    return true;
  }
  /** Server says the blob is fetchable now — retry the parts waiting on it. */
  async onReady(docId, sha) {
    this.waitingFor.delete(sha);
    const doc = this.getDoc();
    if (!doc) return;
    for (const [part, meta] of [...doc.pendingMedia]) {
      if (meta.sha !== sha || this.inFlight.has(part)) continue;
      this.inFlight.add(part);
      void this.fetchOne(docId, part, meta).finally(() => this.inFlight.delete(part));
    }
  }
  /** No holder is online. The registration stays; a later media-ready (or a
   * holder rejoining, §5.4) recovers it with no action from the user. */
  onUnavailable(sha) {
    const doc = this.getDoc();
    if (!doc) return;
    for (const [part, meta] of doc.pendingMedia) {
      if (meta.sha === sha) this.setState(part, "unavailable");
    }
  }
  /** Shas this replica can serve: parts whose bytes are PRESENT. */
  heldShas() {
    const doc = this.getDoc();
    if (!doc) return [];
    const out = [];
    for (const [part, meta] of doc.mediaMeta) {
      if (doc.mediaStatus(part) === "ready") out.push(meta.sha);
    }
    return out;
  }
  /** §5.4: intersect the room's outstanding needs with local holdings and
   * volunteer — this is what makes evicted media reappear when a holder
   * comes back, with no polling anywhere. */
  volunteer(mediaNeeded) {
    if (!mediaNeeded?.length) return;
    const held = new Set(this.heldShas());
    const intersection = mediaNeeded.filter((sha) => held.has(sha));
    if (intersection.length) this.cb.have(intersection);
  }
  /** Someone in the room needs a sha: answer if we hold it (§5.3). */
  answerRequest(sha) {
    if (this.heldShas().includes(sha)) this.cb.have([sha]);
  }
  /**
   * HOLDER DUTY (doc 16 §5.3): chosen to re-supply — upload our copy.
   *
   * In an encrypted room this RE-SEALS with the IV recorded on the part, not
   * a fresh one: same key + same IV + same plaintext reproduces the exact
   * ciphertext, which is the only thing that still hashes to the address the
   * intent committed to. A fresh IV would produce a perfectly valid blob at
   * the WRONG address, and the upload would be rejected — do not "fix" it
   * that way (doc 16 §5.3 / doc 13 §4 both carry this warning).
   *
   * The local sha assertion before PUT is deliberate: failing it means THIS
   * replica's pixels are corrupt, and uploading them would just burn the
   * re-supply rotation on bytes the relay will reject anyway.
   */
  async resupply(docId, sha) {
    const doc = this.getDoc();
    if (!doc) return false;
    for (const [part, meta] of doc.mediaMeta) {
      if (meta.sha !== sha || doc.mediaStatus(part) !== "ready") continue;
      const pixels = doc.media(part);
      if (!pixels) continue;
      let blob = pixels;
      if (this.crypto) {
        if (!meta.iv) return false;
        blob = (await this.crypto.seal(pixels, meta.iv)).blob;
      }
      if (await sha256Hex(blob) !== sha) return false;
      const res = await putBlob(this.opts, docId, sha, blob);
      return res.ok;
    }
    return false;
  }
};

// ../collab/src/e2ee.ts
var te = new TextEncoder();
var td = new TextDecoder();
var INTENT_PAD_RUNGS = [384, 1024, 4096];
var INTENT_PAD_STEP = 4096;
var INTENT_PLAIN_CAP = 256 * 1024 * 3 / 4 - 16;
var PRESENCE_PAD_RUNGS = [128, 1024];
var PRESENCE_PAD_STEP = 1024;
var PRESENCE_PLAIN_CAP = INTENT_PLAIN_CAP;
var CHECKPOINT_PAD_STEP = 64 * 1024;
var CHECKPOINT_PLAIN_CAP = 16 * 1024 * 1024 * 3 / 4 - 16;
function padBody(body, rungs, step, cap) {
  const need = 4 + body.length;
  let target = rungs.find((r) => need <= r) ?? Math.ceil(need / step) * step;
  if (target > cap) target = cap;
  if (target < need) target = need;
  const out = new Uint8Array(target);
  new DataView(out.buffer).setUint32(0, body.length);
  out.set(body, 4);
  return out;
}
function unpadBody(pt) {
  const bytes = new Uint8Array(pt);
  if (bytes.length < 4) throw new Error("padded body: too short");
  const len = new DataView(pt).getUint32(0);
  if (len > bytes.length - 4) throw new Error("padded body: bad length prefix");
  return bytes.subarray(4, 4 + len);
}
function bytesToB64(bytes) {
  let bin = "";
  for (let i = 0; i < bytes.length; i++) bin += String.fromCharCode(bytes[i]);
  return typeof btoa === "function" ? btoa(bin) : Buffer.from(bytes).toString("base64");
}
function b64ToBytes(b64) {
  const bin = typeof atob === "function" ? atob(b64) : Buffer.from(b64, "base64").toString("binary");
  const out = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
  return out;
}
function mintDocKey() {
  const bytes = new Uint8Array(32);
  crypto.getRandomValues(bytes);
  return bytesToB64(bytes).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
}
function docKeyFromFragment(fragment) {
  const m = /(?:^|[#&])k=([A-Za-z0-9_-]+)/.exec(fragment);
  return m ? m[1] : null;
}
async function deriveEpochKeys(docKeyB64url, genesisId, stretchedCode) {
  const raw = b64ToBytes(docKeyB64url.replace(/-/g, "+").replace(/_/g, "/"));
  const master = await crypto.subtle.importKey("raw", raw, "HKDF", false, ["deriveKey"]);
  const salt = te.encode(`wordinweb-epoch:${genesisId}`);
  const mix = stretchedCode ?? new Uint8Array(0);
  const derive = (info) => crypto.subtle.deriveKey(
    { name: "HKDF", hash: "SHA-256", salt, info: concat(te.encode(info), mix) },
    master,
    { name: "AES-GCM", length: 256 },
    false,
    ["encrypt", "decrypt"]
  );
  return {
    kContent: await derive("content"),
    kMedia: await derive("media"),
    kPresence: await derive("presence")
  };
}
async function stretchShareCode(code, docId) {
  const key = await crypto.subtle.importKey("raw", te.encode(code), "PBKDF2", false, ["deriveBits"]);
  const bits = await crypto.subtle.deriveBits(
    { name: "PBKDF2", hash: "SHA-256", salt: te.encode(`wordinweb-code:${docId}`), iterations: 6e5 },
    key,
    256
  );
  return new Uint8Array(bits);
}
async function sealIntent(kContent, docId, genesisId, intent) {
  const iv = new Uint8Array(12);
  crypto.getRandomValues(iv);
  const aad = intentAad(docId, genesisId, intent.clientId, intent.clientSeq, intent.base);
  const body = padBody(te.encode(JSON.stringify(intent)), INTENT_PAD_RUNGS, INTENT_PAD_STEP, INTENT_PLAIN_CAP);
  const ct = await crypto.subtle.encrypt(
    { name: "AES-GCM", iv, additionalData: aad },
    kContent,
    body
  );
  return {
    clientId: intent.clientId,
    clientSeq: intent.clientSeq,
    base: intent.base,
    iv: bytesToB64(iv),
    ciphertext: bytesToB64(new Uint8Array(ct))
  };
}
async function openIntent(kContent, docId, genesisId, env) {
  const aad = intentAad(docId, genesisId, env.clientId, env.clientSeq, env.base);
  const pt = await crypto.subtle.decrypt(
    { name: "AES-GCM", iv: b64ToBytes(env.iv), additionalData: aad },
    kContent,
    b64ToBytes(env.ciphertext)
  );
  const intent = JSON.parse(td.decode(unpadBody(pt)));
  if (intent.clientId !== env.clientId || intent.clientSeq !== env.clientSeq || intent.base !== env.base) {
    throw new Error("envelope bookkeeping mismatch");
  }
  return intent;
}
async function sealCheckpoint(kContent, docId, genesisId, seq, body) {
  const iv = new Uint8Array(12);
  crypto.getRandomValues(iv);
  const aad = te.encode(`cp:${docId}:${genesisId}:${seq}`);
  const padded = padBody(te.encode(JSON.stringify(body)), [], CHECKPOINT_PAD_STEP, CHECKPOINT_PLAIN_CAP);
  const ct = await crypto.subtle.encrypt(
    { name: "AES-GCM", iv, additionalData: aad },
    kContent,
    padded
  );
  return { iv: bytesToB64(iv), ciphertext: bytesToB64(new Uint8Array(ct)) };
}
async function openCheckpoint(kContent, docId, genesisId, seq, sealed) {
  const aad = te.encode(`cp:${docId}:${genesisId}:${seq}`);
  const pt = await crypto.subtle.decrypt(
    { name: "AES-GCM", iv: b64ToBytes(sealed.iv), additionalData: aad },
    kContent,
    b64ToBytes(sealed.ciphertext)
  );
  return JSON.parse(td.decode(unpadBody(pt)));
}
function intentAad(docId, genesisId, clientId, clientSeq, base) {
  return te.encode(`in:${docId}:${genesisId}:${clientId}:${clientSeq}:${base}`);
}
function presenceAad(docId, genesisId, clientId) {
  return te.encode(`pr:${docId}:${genesisId}:${clientId}`);
}
function isSealedPresence(pos) {
  return !!pos && typeof pos === "object" && typeof pos.iv === "string" && typeof pos.ciphertext === "string";
}
async function sealPresence(kPresence, docId, genesisId, clientId, position) {
  const iv = new Uint8Array(12);
  crypto.getRandomValues(iv);
  const ct = await crypto.subtle.encrypt(
    {
      name: "AES-GCM",
      iv,
      additionalData: presenceAad(docId, genesisId, clientId)
    },
    kPresence,
    padBody(te.encode(JSON.stringify(position)), PRESENCE_PAD_RUNGS, PRESENCE_PAD_STEP, PRESENCE_PLAIN_CAP)
  );
  return { iv: bytesToB64(iv), ciphertext: bytesToB64(new Uint8Array(ct)) };
}
async function openPresence(kPresence, docId, genesisId, clientId, sealed) {
  const pt = await crypto.subtle.decrypt(
    {
      name: "AES-GCM",
      iv: b64ToBytes(sealed.iv),
      additionalData: presenceAad(docId, genesisId, clientId)
    },
    kPresence,
    b64ToBytes(sealed.ciphertext)
  );
  return JSON.parse(td.decode(unpadBody(pt)));
}
function concat(a, b) {
  const out = new Uint8Array(a.length + b.length);
  out.set(a, 0);
  out.set(b, a.length);
  return out;
}
async function sealMediaBlob(kMedia, plaintext, ivB64) {
  let iv;
  if (ivB64) iv = b64ToBytes(ivB64);
  else {
    iv = new Uint8Array(12);
    crypto.getRandomValues(iv);
  }
  const ct = new Uint8Array(
    await crypto.subtle.encrypt({ name: "AES-GCM", iv }, kMedia, plaintext)
  );
  return { blob: ct, iv: bytesToB64(iv) };
}
async function openMediaBlob(kMedia, blob, ivB64) {
  return new Uint8Array(
    await crypto.subtle.decrypt({ name: "AES-GCM", iv: b64ToBytes(ivB64) }, kMedia, blob)
  );
}

// ../collab/src/connection.ts
function base64ToBytes(b64) {
  const bin = typeof atob === "function" ? atob(b64) : Buffer.from(b64, "base64").toString("binary");
  const out = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
  return out;
}
var _CollabConnection = class _CollabConnection {
  constructor(transport, clientId, cb = {}, mediaOpts) {
    this.transport = transport;
    this.clientId = clientId;
    this.cb = cb;
    this.replica = null;
    this.clientSeq = 0;
    /** Bumps only when reconciliation RELOADED the document (a true conflict) —
     * the renderer re-mounts on this, but updates in place otherwise (no flash
     * for the common non-conflicting edits). */
    this.docEpoch = 0;
    /** Per-client id allocator for carried node ids (split/format-range/insert):
     * a client-specific base block keeps concurrently-allocated ids disjoint
     * across clients so they never collide. */
    this.idCounter = 0;
    this.idBase = -1;
    /** The epoch id of the session this connection joined (from the welcome).
     * The resume layer compares it against the bundle's stored genesisId:
     * same ⇒ seamless rejoin; different ⇒ someone re-seeded while away
     * (doc 12 §5 case 2 — take server state, keep the local copy; the doc-15
     * lineage check decides fast-forward vs draft). Null until welcomed. */
    this.genesisId = null;
    /** Session encryption mode from the welcome (doc 13 §6). The E2EE layer
     * hard-refuses a value contradicting the link's `#k` fragment. */
    this.mode = null;
    /**
     * The relay's per-blob upload limit in bytes, from the welcome — so a file
     * can be checked BEFORE it is uploaded rather than after a 413.
     *
     * `null` means the server did not publish one (an older host). Callers must
     * treat that as "skip the pre-check", NOT as "no limit" and not as a
     * default they invent: the server still enforces its real limit either way,
     * and a client guessing a larger number would promise the user an upload
     * that cannot succeed.
     */
    this.mediaMaxBlobBytes = null;
    /** Out-of-band media transfer (doc 16 §5). Null until `httpBase` is
     * supplied — without a relay origin a client simply has no media duties
     * (every existing caller keeps working unchanged). */
    this.media = null;
    /** The room this connection joined; media URLs are per-doc. */
    this.docId = "";
    /** The latest roster snapshot (doc 14 §2); empty until the first fan-out. */
    this.roster = [];
    /**
     * Attribution layer 1 (doc 14 §3): the canonical log IS the attribution
     * record — every applied entry carries its author's bound clientId. This
     * is the bounded client-side derivation an activity panel renders
     * (clientId → roster name/color). Newest last, capped, derived state:
     * never persisted, dies with the connection (zero custody).
     */
    this.activity = [];
    /** Bundle being resumed from (set by resume(), consumed at welcome). */
    this.resuming = null;
    /** Accumulates doc changes across replica rebuilds (welcome), so the getter
     * below stays monotonic per connection. */
    this.docVersionBase = 0;
    /** docVersionBase at the last takeRenderScope — a base bump between takes
     * (welcome doc replacement, media install) has no narrow scope. */
    this.takenDocVersionBase = 0;
    /** Submits dropped because the connection was not ready (see
     * onSubmitDropped). Exposed like `selfHeals` so a harness or a UI can read
     * it: when a run loses intents, this number says how many died HERE rather
     * than anywhere else on the path. */
    this.droppedPreReady = 0;
    this.rlRedriveTimer = null;
    this.rlRedriveBackoffMs = 0;
    this.transport.onMessage((msg) => this.onServer(msg));
    if (mediaOpts) {
      this.media = new MediaClient(mediaOpts, () => this.replica?.doc ?? null, {
        // Media installs mutate the doc outside the replica's counter — bump
        // the render signal so the skeleton repaints (see docVersion).
        onChange: () => {
          this.docVersionBase++;
          this.cb.onChange?.();
        },
        onState: (part, state) => this.cb.onMediaState?.(part, state),
        need: (sha) => this.mediaNeed(sha),
        have: (shas) => this.mediaHave(shas)
      });
    }
  }
  /** Upload bytes and get the intent's media fields, or null if the relay
   * refused — the caller must not emit an insertImage in that case. */
  async uploadMedia(plaintext) {
    return this.media ? this.media.upload(this.docId, plaintext) : null;
  }
  /** Replace the callbacks (used by bindEditor to attach after construction). */
  setCallbacks(cb) {
    this.cb = cb;
  }
  /** Allocate `n` fresh carried node ids in this client's disjoint block. */
  allocIds(n) {
    if (this.idBase < 0) {
      let h = 0;
      for (let i = 0; i < this.clientId.length; i++) h = h * 31 + this.clientId.charCodeAt(i) >>> 0;
      this.idBase = 1e9 + h % 1e5 * 1e7;
    }
    const out = [];
    for (let i = 0; i < n; i++) out.push(this.idBase + this.idCounter++);
    return out;
  }
  /**
   * Join a document. The server replies with a welcome (snapshot + sidecar +
   * tail). The hello carries this connection's clientId — the hub binds it to
   * the socket and refuses submits under any other id (doc 11 decision 8).
   * `takeover: true` claims the identity from an existing live connection
   * (the doc-12 §7 "use here instead" path for a second same-profile tab);
   * without it, a duplicate join is refused `already-open`.
   */
  join(docId, token, opts) {
    this.docId = docId;
    this.transport.send({
      t: "hello",
      protocolVersion: PROTOCOL_VERSION,
      docId,
      clientId: this.clientId,
      takeover: opts?.takeover,
      token,
      sinceSeq: 0,
      profile: opts?.profile,
      codeProof: opts?.codeProof,
      ownerToken: opts?.ownerToken
    });
  }
  /** Owner admin op (doc 14 §2.5): honored only if this connection proved
   * the owner token at hello — otherwise the server refuses `not-owner`. */
  admin(action) {
    this.transport.send({ t: "admin", action });
  }
  recordActivity(entries) {
    for (const e of entries) {
      if (e.kind === "applied" && e.intent) {
        this.activity.push({ seq: e.seq, clientId: e.intent.clientId, kind: e.intent.kind });
      }
    }
    if (this.activity.length > _CollabConnection.ACTIVITY_CAP) {
      this.activity.splice(0, this.activity.length - _CollabConnection.ACTIVITY_CAP);
    }
  }
  /** Ask the room for a blob this replica needs (doc 16 §5.2). */
  mediaNeed(sha) {
    this.transport.send({ t: "media-need", sha });
  }
  /** Volunteer holdings (reply to a request, or after a welcome whose
   * mediaNeeded intersects local media — §5.4). */
  mediaHave(shas) {
    if (shas.length) this.transport.send({ t: "media-have", shas });
  }
  /** Shas this replica can re-supply: metadata of parts whose bytes are
   * PRESENT (pending parts are exactly what we can't serve). */
  heldMediaShas() {
    const doc = this.replica?.doc;
    if (!doc) return [];
    const out = [];
    for (const [part, meta] of doc.mediaMeta) {
      if (doc.mediaStatus(part) === "ready") out.push(meta.sha);
    }
    return out;
  }
  /** Rename/recolor this participant mid-session. The server sanitizes and
   * fans out the updated roster; the local copy updates on that echo (no
   * optimistic roster — a 1-RTT lag on your own rename is imperceptible and
   * keeps one code path). */
  setProfile(profile) {
    this.transport.send({ t: "profile", profile });
  }
  /**
   * Rejoin a document from a persisted bundle (doc 12 §5). Restores the
   * clientSeq watermark FIRST — a fresh counter would reuse already-
   * sequenced (clientId, clientSeq) keys and the server would dedup this
   * client's NEW edits as re-sends (silent edit loss). The welcome decides
   * the case: same epoch ⇒ replay pending (below); different ⇒
   * onEpochChange, pending withheld.
   */
  resume(bundle, token, opts) {
    this.clientSeq = Math.max(this.clientSeq, bundle.clientSeq);
    this.resuming = bundle;
    this.docId = bundle.docId;
    this.transport.send({
      profile: opts?.profile,
      codeProof: opts?.codeProof,
      t: "hello",
      protocolVersion: PROTOCOL_VERSION,
      docId: bundle.docId,
      clientId: this.clientId,
      // Resume IS the zombie-takeover case (doc 12 §7): after a crash or
      // refresh, the previous socket for this identity may still be inside
      // the 60s grace window server-side. This connection is that identity's
      // continuation — claiming it is correct, and the incumbent (if any)
      // is a dead tab by definition of "we are resuming from the bundle".
      takeover: true,
      token,
      sinceSeq: bundle.confirmedSeq,
      genesisId: bundle.genesisId,
      ownerToken: opts?.ownerToken
    });
  }
  /**
   * The current durable state as a doc-12 bundle, or null before welcome.
   * `savedAt` is stamped by the persister at write time (clock injection —
   * this module never reads Date.now itself).
   */
  exportBundle(docId) {
    if (!this.replica || this.genesisId === null) return null;
    return {
      docId,
      genesisId: this.genesisId,
      ...this.replica.exportBundleState(),
      clientSeq: this.clientSeq,
      savedAt: 0,
      lineage: []
      // the persister maintains the chain across writes
    };
  }
  async exportBundleAsync(docId) {
    if (!this.replica || this.genesisId === null) return null;
    return {
      docId,
      genesisId: this.genesisId,
      ...await this.replica.exportBundleStateAsync(),
      clientSeq: this.clientSeq,
      savedAt: 0,
      lineage: []
    };
  }
  /** The live document (null until welcome). The editor renders this. */
  get doc() {
    return this.replica?.doc ?? null;
  }
  /**
   * Counts changes TO THE RENDERED DOCUMENT — the repaint signal, as opposed
   * to onChange (which also fires for pure bookkeeping such as a tracked own
   * echo). The editor-driven typing path mutates and paints the doc itself,
   * so its submit + echo leave this untouched; remote applies, optimistic
   * canonical applies (toolbar ops), reloads, and media installs advance it.
   * The react layer repaints on THIS, not on onChange — repainting on
   * onChange queued a whole-document relayout per keystroke, catastrophic
   * past the background-layout page threshold.
   */
  get docVersion() {
    return this.docVersionBase + (this.replica?.docVersion ?? 0);
  }
  /**
   * Drain the dirty scope behind the docVersion movement since the last take
   * — what the repaint answering `docVersion` must relayout. Replica applies
   * carry their per-intent scope; connection-level bumps (a replaced doc
   * object, media pixels landing) report document scope. Null means NOTHING
   * was recorded since the last take — the caller already painted everything
   * and may skip the repaint.
   */
  takeRenderScope() {
    const replicaScope = this.replica?.takeRenderScope() ?? null;
    if (this.docVersionBase !== this.takenDocVersionBase) {
      this.takenDocVersionBase = this.docVersionBase;
      return { kind: "doc" };
    }
    return replicaScope;
  }
  get ready() {
    return this.replica !== null;
  }
  /** Un-confirmed local intents in flight — the drained-replay discipline
   * (doc 15 §2 / rebase.ts) polls this to submit one intent at a time. */
  get pendingCount() {
    return this.replica?.pendingCount ?? 0;
  }
  /**
   * Submit a local edit. The caller supplies the intent minus its wire
   * bookkeeping (clientId/clientSeq/base) — the connection fills those from
   * the current confirmed seq and applies it optimistically before sending.
   */
  submit(intent) {
    this.submitFull(
      intent,
      /*preApplied*/
      false
    );
  }
  /** Submit an intent whose mutation the caller ALREADY performed on this
   * connection's live doc (the editor-driven path: DocxEditor applies the
   * command to `conn.doc` and then emits the intent). Skips the optimistic
   * re-apply — applying twice doubled every keystroke — but tracks pending and
   * sends identically, so echoes and reconciliation work unchanged. */
  submitPreApplied(intent) {
    this.submitFull(
      intent,
      /*preApplied*/
      true
    );
  }
  submitFull(intent, preApplied) {
    if (!this.replica) {
      this.droppedPreReady++;
      this.cb.onSubmitDropped?.("not-ready");
      return;
    }
    const full = {
      ...intent,
      clientId: this.clientId,
      clientSeq: ++this.clientSeq,
      base: this.replica.confirmedSeq
    };
    if (preApplied) this.replica.trackLocal(full);
    else this.replica.submitLocal(full);
    this.transport.send({ t: "submit", intent: full });
    this.cb.onChange?.();
  }
  /** Broadcast this client's cursor/selection (ephemeral). Selection ranges
   * are clamped on the way OUT too, so an over-long selection costs the room
   * a bounded payload rather than every receiver a discard. */
  setPresence(position) {
    this.transport.send({ t: "presence", position: sanitizePresencePosition(position) });
  }
  /**
   * Surface any of OUR OWN intents that canonical validation rejected.
   *
   * A rejection is sequenced and agreed by every replica, so it produces no
   * divergence to notice — which is exactly how a stale image-size bound
   * silently ate a user's photos. See ConnectionCallbacks.onIntentRejected.
   */
  reportRejections(entries) {
    if (!this.cb.onIntentRejected) return;
    for (const e of entries) {
      if (e.kind === "rejected" && e.clientId === this.clientId) {
        this.cb.onIntentRejected({ reason: e.reason, clientSeq: e.clientSeq });
      }
    }
  }
  onServer(msg) {
    switch (msg.t) {
      case "welcome": {
        this.docVersionBase += (this.replica?.docVersion ?? 0) + 1;
        this.replica = new ClientReplica(base64ToBytes(msg.snapshot), msg.sidecar);
        this.replica.confirmedSeq = msg.seq;
        this.genesisId = msg.genesisId;
        this.mode = msg.mode;
        this.mediaMaxBlobBytes = msg.mediaMaxBlobBytes ?? null;
        if (msg.tail.length) this.replica.receive(msg.tail);
        if (this.resuming?.mediaMeta) {
          const held = DocxDocument.load(this.resuming.confirmedBytes);
          for (const [part, meta] of this.resuming.mediaMeta) {
            this.replica.doc.mediaMeta.set(part, meta);
            const bytes = held.media(part);
            if (bytes && this.replica.doc.mediaStatus(part) === "pending") {
              this.replica.doc.installMedia(part, bytes);
            }
          }
        }
        if (this.replica) applyMediaAddresses(this.replica.doc, msg.media);
        if (this.media) this.media.volunteer(msg.mediaNeeded);
        else if (msg.mediaNeeded?.length) {
          const held = new Set(this.heldMediaShas());
          this.mediaHave(msg.mediaNeeded.filter((sha) => held.has(sha)));
        }
        this.recordActivity(msg.tail);
        const resumed = this.resuming;
        this.resuming = null;
        if (resumed) {
          if (resumed.genesisId === msg.genesisId) {
            for (const intent of resumed.pending) this.transport.send({ t: "submit", intent });
          } else if (isAncestorOf(resumed, msg.lineage)) {
            this.cb.onFastForward?.(resumed.genesisId, msg.genesisId);
          } else {
            this.cb.onEpochChange?.(resumed.genesisId, msg.genesisId);
          }
        }
        this.cb.onChange?.();
        void this.media?.fetchPending(this.docId);
        return;
      }
      case "broadcast": {
        this.replica?.receive(msg.entries);
        this.reportRejections(msg.entries);
        this.recordActivity(msg.entries);
        if (this.replica?.reloaded) this.docEpoch++;
        this.cb.onChange?.();
        void this.media?.fetchPending(this.docId);
        return;
      }
      case "presence": {
        if (isSealedPresence(msg.position)) return;
        this.cb.onPresence?.(msg.participant, sanitizePresencePosition(msg.position));
        return;
      }
      case "roster": {
        this.roster = msg.roster;
        this.cb.onRoster?.(msg.roster);
        return;
      }
      case "session-warning": {
        this.cb.onSessionWarning?.({ reason: msg.reason, inMs: msg.inMs });
        return;
      }
      case "session-warning-cleared": {
        this.cb.onSessionWarningCleared?.({ reason: msg.reason });
        return;
      }
      case "media-request": {
        this.media?.answerRequest(msg.sha);
        this.cb.onMediaRequest?.(msg.sha);
        return;
      }
      case "media-ready": {
        void this.media?.onReady(this.docId, msg.sha);
        this.cb.onMediaReady?.(msg.sha);
        return;
      }
      case "media-unavailable": {
        this.media?.onUnavailable(msg.sha);
        this.cb.onMediaUnavailable?.(msg.sha);
        return;
      }
      case "media-upload": {
        void this.media?.resupply(this.docId, msg.sha);
        this.cb.onMediaRequest?.(msg.sha);
        return;
      }
      case "refused": {
        if (msg.reason === "rate-limit") {
          if (this.rlRedriveTimer !== null) return;
          this.rlRedriveBackoffMs = Math.min(this.rlRedriveBackoffMs === 0 ? 300 : this.rlRedriveBackoffMs * 2, 5e3);
          this.rlRedriveTimer = setTimeout(() => {
            this.rlRedriveTimer = null;
            if (!this.replica) return;
            const queue = this.replica.pendingCopies();
            if (queue.length === 0) {
              this.rlRedriveBackoffMs = 0;
              return;
            }
            for (const body of queue) {
              this.transport.send({ t: "submit", intent: { ...body, base: this.replica.confirmedSeq } });
            }
          }, this.rlRedriveBackoffMs);
          return;
        }
        this.cb.onRefused?.(msg.reason);
        return;
      }
    }
  }
};
_CollabConnection.ACTIVITY_CAP = 100;
var CollabConnection = _CollabConnection;
function isAncestorOf(bundle, seedLineage) {
  if (!seedLineage) return false;
  const ownHead = bundle.lineage?.[bundle.lineage.length - 1];
  if (!ownHead) return false;
  return seedLineage.some(
    (h) => h.genesisId === ownHead.genesisId && h.seq >= ownHead.seq && h.docHash === ownHead.docHash
  );
}

// ../collab/src/invert.ts
function invertIntent(doc, ids, intent) {
  switch (intent.kind) {
    case "insertText": {
      return {
        kind: "deleteText",
        blockId: intent.at.blockId,
        runId: intent.at.runId,
        start: intent.at.offset,
        end: intent.at.offset + intent.text.length
      };
    }
    case "deleteText": {
      const runEl = ids.elOf(intent.runId);
      if (!runEl) return null;
      const hit = resolveRunOffset(runEl, intent.start);
      if (!hit) return null;
      const localEnd = hit.offset + (intent.end - intent.start);
      if (localEnd > hit.t.text.length) return null;
      const removed = hit.t.text.slice(hit.offset, localEnd);
      if (removed.length === 0) return null;
      return {
        kind: "insertText",
        at: { blockId: intent.blockId, runId: intent.runId, offset: intent.start },
        text: removed
      };
    }
    case "splitParagraph": {
      return { kind: "mergeParagraph", blockId: intent.newBlockId };
    }
    case "insertImage": {
      const carrier = intent.nodeIds?.[0];
      if (carrier === void 0) return null;
      return { kind: "removeDrawing", runId: carrier };
    }
    // Formatting / table / comment / merge inverses need prior-state capture
    // (was-bold?, the merged paragraph's structure, …) — a documented
    // extension; not undoable yet.
    default:
      return null;
  }
}

// ../collab/src/validate.ts
var MAX_IMAGE_BYTES = 256 * 1024 * 1024;
var DEFAULT_INTENT_LIMITS = {
  maxInsertLength: 1e5,
  maxDeleteLength: 1e6,
  maxCommentLength: 2e4,
  maxPasteBytes: 2e6
};
function chartError(c, who) {
  if (typeof c !== "object" || c === null) return `${who}: bad chart`;
  if (!["column", "bar", "line", "pie"].includes(c.type)) return `${who}: bad type`;
  if (!Array.isArray(c.categories) || c.categories.length === 0 || c.categories.length > 100) return `${who}: bad categories`;
  if (c.categories.some((x) => typeof x !== "string" || x.length > 200)) return `${who}: bad category`;
  if (c.title !== void 0 && (typeof c.title !== "string" || c.title.length > 200)) return `${who}: bad title`;
  if (!Array.isArray(c.series) || c.series.length === 0 || c.series.length > 24) return `${who}: bad series`;
  for (const s of c.series) {
    if (typeof s !== "object" || s === null || typeof s.name !== "string" || s.name.length > 200) return `${who}: bad series name`;
    if (!Array.isArray(s.values) || s.values.length > 100 || s.values.some((v) => typeof v !== "number" || !Number.isFinite(v))) return `${who}: bad series values`;
  }
  return null;
}
function smartArtError(a, who) {
  if (typeof a !== "object" || a === null) return `${who}: bad smartArt`;
  if (!["process", "cycle", "hierarchy", "list"].includes(a.layout)) return `${who}: bad layout`;
  if (!Array.isArray(a.items) || a.items.length === 0 || a.items.length > 50) return `${who}: bad items`;
  if (a.items.some((x) => typeof x !== "string" || x.length > 500)) return `${who}: bad item`;
  return null;
}
function validateIntent(intent, limits = DEFAULT_INTENT_LIMITS) {
  const nonNegInt = (n) => Number.isInteger(n) && n >= 0;
  switch (intent.kind) {
    case "insertText":
      if (typeof intent.text !== "string") return "insertText: text not a string";
      if (intent.text.length === 0) return "insertText: empty";
      if (intent.text.length > limits.maxInsertLength) return "insertText: too long";
      if (!nonNegInt(intent.at.offset)) return "insertText: bad offset";
      return null;
    case "suggestRevision": {
      const nRanges = intent.ranges?.length ?? 0;
      const nMarks = intent.marks?.length ?? 0;
      if (nRanges + nMarks === 0) return "suggestRevision: empty";
      if (nRanges > 100 || nMarks > 20) return "suggestRevision: too many";
      for (const r of intent.ranges ?? []) {
        if (!nonNegInt(r.start) || !nonNegInt(r.end) || r.end <= r.start) return "suggestRevision: bad range";
        if (r.end - r.start > limits.maxDeleteLength) return "suggestRevision: too large";
      }
      for (const m of intent.marks ?? []) {
        if (m.glyph !== "ins" && m.glyph !== "del") return "suggestRevision: bad glyph";
      }
      if (typeof intent.suggest?.author !== "string" || intent.suggest.author.length > 100) return "suggestRevision: bad author";
      if (typeof intent.suggest?.date !== "string" || intent.suggest.date.length > 40) return "suggestRevision: bad date";
      return null;
    }
    case "deleteText":
      if (!nonNegInt(intent.start) || !nonNegInt(intent.end)) return "deleteText: bad range";
      if (intent.end <= intent.start) return "deleteText: empty range";
      if (intent.end - intent.start > limits.maxDeleteLength) return "deleteText: too large";
      return null;
    case "splitParagraph":
      if (!nonNegInt(intent.at.offset)) return "splitParagraph: bad offset";
      if (!nonNegInt(intent.newBlockId) || !nonNegInt(intent.newRunId)) return "splitParagraph: bad ids";
      return null;
    case "formatRange":
      if (!nonNegInt(intent.start) || !nonNegInt(intent.end) || intent.end <= intent.start) return "formatRange: bad range";
      if (!nonNegInt(intent.middleId)) return "formatRange: bad middleId";
      return null;
    case "commentRun":
      if (typeof intent.text !== "string" || intent.text.length === 0) return "commentRun: empty";
      if (intent.text.length > limits.maxCommentLength) return "commentRun: too long";
      if (typeof intent.paraId !== "string" || typeof intent.date !== "string") return "commentRun: bad provenance";
      return null;
    case "pasteBlocks":
      if (typeof intent.blocksXml !== "string" || intent.blocksXml.length === 0) return "pasteBlocks: empty";
      if (intent.blocksXml.length > limits.maxPasteBytes) return "pasteBlocks: too large";
      if (!Array.isArray(intent.nodeIds)) return "pasteBlocks: bad nodeIds";
      return null;
    case "insertImage":
      if (!/^[0-9a-f]{64}$/.test(intent.blobSha)) return "insertImage: bad sha";
      if (!nonNegInt(intent.bytesLen) || intent.bytesLen === 0 || intent.bytesLen > MAX_IMAGE_BYTES)
        return "insertImage: bad size";
      if (!["png", "jpg", "jpeg", "gif", "bmp", "webp"].includes(intent.ext)) return "insertImage: bad ext";
      if (intent.iv !== void 0 && !/^[A-Za-z0-9+/]{16}$/.test(intent.iv)) return "insertImage: bad iv";
      if (!Number.isFinite(intent.widthPx) || !Number.isFinite(intent.heightPx)) return "insertImage: bad size";
      return null;
    case "insertBreak":
      if (intent.breakKind !== "page" && intent.breakKind !== "column") return "insertBreak: bad kind";
      return null;
    case "insertMath":
      if (typeof intent.mathText !== "string" || intent.mathText.length === 0) return "insertMath: empty";
      if (intent.mathText.length > 1e4) return "insertMath: too long";
      return null;
    case "insertShape": {
      const presets = ["line", "verticalLine", "rectangle", "roundedRectangle", "ellipse", "diamond", "textBox"];
      if (!presets.includes(intent.preset)) return "insertShape: bad preset";
      return null;
    }
    case "replyComment":
      if (typeof intent.text !== "string" || intent.text.length === 0) return "replyComment: empty";
      if (intent.text.length > 2e4) return "replyComment: too long";
      if (typeof intent.parentId !== "string" || !Array.isArray(intent.paraIds) || typeof intent.date !== "string") return "replyComment: bad fields";
      return null;
    case "adjustIndent":
      if (intent.direction !== 1 && intent.direction !== -1) return "adjustIndent: bad direction";
      return null;
    case "setSpacing":
      if (typeof intent.patch !== "object" || intent.patch === null) return "setSpacing: bad patch";
      return null;
    case "insertPageField":
      if (intent.fieldKind !== "page" && intent.fieldKind !== "pageOfTotal") return "insertPageField: bad kind";
      return null;
    case "setLink":
      if (typeof intent.url !== "string" || intent.url.length === 0 || intent.url.length > 4e3) return "setLink: bad url";
      return null;
    case "insertFootnote":
      if (typeof intent.text !== "string" || intent.text.trim().length === 0) return "insertFootnote: empty";
      if (intent.text.length > 2e4) return "insertFootnote: too long";
      return null;
    case "setDropCap":
      if (intent.mode !== null && intent.mode !== "drop" && intent.mode !== "margin") return "setDropCap: bad mode";
      return null;
    case "setDivider":
      return null;
    case "insertBookmark":
      if (typeof intent.name !== "string" || intent.name.length === 0 || intent.name.length > 400) return "insertBookmark: bad name";
      return null;
    case "insertBlankPage":
      return null;
    case "insertSectionBreak":
      if (intent.breakType !== "nextPage" && intent.breakType !== "continuous") return "insertSectionBreak: bad type";
      return null;
    case "insertCrossRef":
      if (typeof intent.bookmark !== "string" || intent.bookmark.length === 0) return "insertCrossRef: bad bookmark";
      if (intent.refKind !== "text" && intent.refKind !== "page") return "insertCrossRef: bad kind";
      return null;
    case "insertCoverPage":
      if (typeof intent.content !== "object" || intent.content === null) return "insertCoverPage: bad content";
      return null;
    case "setPageLayout": {
      const p = intent.patch;
      if (typeof p !== "object" || p === null) return "setPageLayout: bad patch";
      const rec = p;
      const allowed = ["margins", "mirrorMargins", "size", "orientation", "columns", "columnSeparator", "pageBorders"];
      for (const k of Object.keys(rec)) if (!allowed.includes(k)) return `setPageLayout: unknown key ${k}`;
      const inRange = (v, lo, hi) => typeof v === "number" && Number.isFinite(v) && v >= lo && v <= hi;
      if (rec.orientation !== void 0 && rec.orientation !== "portrait" && rec.orientation !== "landscape") return "setPageLayout: bad orientation";
      if (rec.columns !== void 0 && !inRange(rec.columns, 1, 12)) return "setPageLayout: bad columns";
      if (rec.margins !== void 0) {
        if (typeof rec.margins !== "object" || rec.margins === null) return "setPageLayout: bad margins";
        for (const [mk, mv] of Object.entries(rec.margins)) {
          if (!["top", "right", "bottom", "left"].includes(mk)) return `setPageLayout: bad margin ${mk}`;
          if (!inRange(mv, 0, 22)) return "setPageLayout: margin out of range";
        }
      }
      if (rec.size !== void 0) {
        const s = rec.size;
        if (typeof s !== "object" || s === null || !inRange(s.width, 1, 22) || !inRange(s.height, 1, 22)) return "setPageLayout: bad size";
      }
      return null;
    }
    case "setListLevel":
      if (intent.delta !== 1 && intent.delta !== -1) return "setListLevel: bad delta";
      return null;
    case "insertWordArt": {
      if (typeof intent.text !== "string" || intent.text.length === 0) return "insertWordArt: empty text";
      if (intent.text.length > 500) return "insertWordArt: text too long";
      if (!["plain", "archUp", "archDown", "wave", "chevron"].includes(intent.preset)) return "insertWordArt: bad preset";
      return null;
    }
    case "insertChart":
      return chartError(intent.chart, "insertChart");
    case "setChartData":
      return chartError(intent.chart, "setChartData");
    case "setSmartArtNodeText":
      if (typeof intent.index !== "number" || !Number.isInteger(intent.index) || intent.index < 0 || intent.index > 1e3) return "setSmartArtNodeText: bad index";
      if (typeof intent.text !== "string" || intent.text.length > 500) return "setSmartArtNodeText: bad text";
      return null;
    case "setDrawingWordArtText":
      if (typeof intent.text !== "string" || intent.text.length === 0 || intent.text.length > 500) return "setDrawingWordArtText: bad text";
      return null;
    case "setDrawingLineStyle":
      if (!/^[0-9a-fA-F]{6}$/.test(intent.color)) return "setDrawingLineStyle: bad color";
      if (typeof intent.widthPx !== "number" || !Number.isFinite(intent.widthPx) || intent.widthPx <= 0 || intent.widthPx > 100) return "setDrawingLineStyle: bad width";
      if (!["solid", "dashed", "dotted"].includes(intent.dash)) return "setDrawingLineStyle: bad dash";
      return null;
    case "setImageAltText":
      if (typeof intent.alt !== "string" || intent.alt.length > 1e3) return "setImageAltText: bad alt";
      return null;
    case "removeLink":
      return null;
    case "insertSmartArt":
      return smartArtError(intent.smartArt, "insertSmartArt");
    case "setSmartArtData":
      return smartArtError(intent.smartArt, "setSmartArtData");
    case "setImageWrap":
      if (!["inline", "square", "topAndBottom", "none", "behind"].includes(intent.mode)) return "setImageWrap: bad mode";
      return null;
    case "setDrawingOrder":
      if (intent.order !== "front" && intent.order !== "back") return "setDrawingOrder: bad order";
      return null;
    case "setSmartArtFill":
      if (intent.color !== null && !/^[0-9a-fA-F]{6}$/.test(intent.color)) return "setSmartArtFill: bad color";
      if (intent.nodeIndex !== void 0 && (typeof intent.nodeIndex !== "number" || !Number.isInteger(intent.nodeIndex) || intent.nodeIndex < 0 || intent.nodeIndex > 1e3)) return "setSmartArtFill: bad nodeIndex";
      return null;
    case "setSmartArtTextFormat": {
      const f = intent.format;
      if (typeof f !== "object" || f === null) return "setSmartArtTextFormat: bad format";
      if (typeof f.fontFamily !== "string" || f.fontFamily.length === 0 || f.fontFamily.length > 100) return "setSmartArtTextFormat: bad fontFamily";
      if (typeof f.fontSizePt !== "number" || !Number.isFinite(f.fontSizePt) || f.fontSizePt < 1 || f.fontSizePt > 1638) return "setSmartArtTextFormat: bad fontSizePt";
      if (!/^[0-9a-fA-F]{6}$/.test(f.color)) return "setSmartArtTextFormat: bad color";
      if (typeof f.bold !== "boolean" || typeof f.italic !== "boolean") return "setSmartArtTextFormat: bad bold/italic";
      if (!["left", "center", "right"].includes(f.alignment)) return "setSmartArtTextFormat: bad alignment";
      if (intent.nodeIndex !== void 0 && (typeof intent.nodeIndex !== "number" || !Number.isInteger(intent.nodeIndex) || intent.nodeIndex < 0 || intent.nodeIndex > 1e3)) return "setSmartArtTextFormat: bad nodeIndex";
      return null;
    }
    case "setFloatingPagePosition": {
      const ok = (v) => typeof v === "number" && Number.isFinite(v) && v >= -5e3 && v <= 5e3;
      if (!ok(intent.xPx) || !ok(intent.yPx)) return "setFloatingPagePosition: bad position";
      return null;
    }
    case "resizeDrawing": {
      const ok = (v) => typeof v === "number" && Number.isFinite(v) && v >= 1 && v <= 5e3;
      if (!ok(intent.widthPx) || !ok(intent.heightPx)) return "resizeDrawing: bad extent";
      return null;
    }
    case "resizeTableColumn": {
      const num = (v, lo, hi) => typeof v === "number" && Number.isFinite(v) && v >= lo && v <= hi;
      if (!Number.isInteger(intent.boundary) || intent.boundary < 1 || intent.boundary > 200) return "resizeTableColumn: bad boundary";
      if (!num(intent.deltaPx, -5e3, 5e3)) return "resizeTableColumn: bad delta";
      if (intent.renderedWidths !== void 0) {
        if (!Array.isArray(intent.renderedWidths) || intent.renderedWidths.length > 200 || !intent.renderedWidths.every((w) => num(w, 0, 2e4))) return "resizeTableColumn: bad widths";
      }
      return null;
    }
    case "resizeTableRow": {
      if (!Number.isInteger(intent.rowIdx) || intent.rowIdx < 0 || intent.rowIdx > 5e3) return "resizeTableRow: bad row";
      if (typeof intent.heightPx !== "number" || !Number.isFinite(intent.heightPx) || intent.heightPx < 1 || intent.heightPx > 2e4) return "resizeTableRow: bad height";
      return null;
    }
    case "moveTable": {
      const num = (v) => typeof v === "number" && Number.isFinite(v) && v >= -5e3 && v <= 2e4;
      if (!num(intent.xPx) || !num(intent.yPx)) return "moveTable: bad position";
      if (typeof intent.preservePageStart !== "boolean") return "moveTable: bad flag";
      if (!Number.isInteger(intent.pageDelta) || Math.abs(intent.pageDelta) > 500) return "moveTable: bad pageDelta";
      return null;
    }
    case "removeDrawing":
      return null;
    // run-addressed only; resolution is the whole check
    case "setMathLinear":
      if (typeof intent.mathText !== "string" || intent.mathText.length === 0 || intent.mathText.length > 1e3) return "setMathLinear: bad mathText";
      return null;
    case "deleteMath":
      return null;
    case "moveMath":
      if (!nonNegInt(intent.at?.offset)) return "moveMath: bad offset";
      if (!Array.isArray(intent.nodeIds) || intent.nodeIds.length > 64) return "moveMath: bad nodeIds";
      return null;
    case "ensureHeaderFooter":
      if (intent.hfKind !== "header" && intent.hfKind !== "footer") return "ensureHeaderFooter: bad kind";
      if (!Array.isArray(intent.nodeIds) || intent.nodeIds.length > 64) return "ensureHeaderFooter: bad nodeIds";
      return null;
    case "deleteComment":
      if (typeof intent.commentId !== "string" || intent.commentId.length === 0 || intent.commentId.length > 64) return "deleteComment: bad id";
      return null;
    case "insertBookmarkRange":
      if (typeof intent.name !== "string" || !/^[A-Za-z_][A-Za-z0-9_]{0,39}$/.test(intent.name)) return "insertBookmarkRange: bad name";
      if (!nonNegInt(intent.start) || !nonNegInt(intent.end) || intent.end <= intent.start) return "insertBookmarkRange: bad range";
      return null;
    case "toggleCheckbox":
      return null;
    case "acceptRevision":
    case "rejectRevision":
      if (!nonNegInt(intent.index) || intent.index > 1e5) return `${intent.kind}: bad index`;
      return null;
    case "acceptAllRevisions":
      return null;
    case "insertTable": {
      const okDim = (v) => typeof v === "number" && Number.isInteger(v) && v >= 1 && v <= 50;
      if (!okDim(intent.rows) || !okDim(intent.cols)) return "insertTable: bad dimensions";
      return null;
    }
    case "setLineNumbering": {
      const p = intent.patch;
      if (typeof p !== "object" || p === null || typeof p.enabled !== "boolean") return "setLineNumbering: bad patch";
      if (p.countBy !== void 0 && (typeof p.countBy !== "number" || !Number.isInteger(p.countBy) || p.countBy < 1 || p.countBy > 100)) return "setLineNumbering: bad countBy";
      if (p.restart !== void 0 && !["continuous", "newPage", "newSection"].includes(p.restart)) return "setLineNumbering: bad restart";
      if (p.start !== void 0 && (typeof p.start !== "number" || !Number.isInteger(p.start) || p.start < 0 || p.start > 1e5)) return "setLineNumbering: bad start";
      return null;
    }
    case "insertDateTimeField": {
      if (intent.dtKind !== "date" && intent.dtKind !== "time") return "insertDateTimeField: bad kind";
      if (typeof intent.picture !== "string" || intent.picture.length === 0 || intent.picture.length > 100) return "insertDateTimeField: bad picture";
      if (/[\\{}"]/.test(intent.picture)) return "insertDateTimeField: illegal picture char";
      return null;
    }
    case "insertField": {
      if (typeof intent.instruction !== "string" || intent.instruction.length === 0 || intent.instruction.length > 200) return "insertField: bad instruction";
      if (intent.cachedResult !== void 0 && (typeof intent.cachedResult !== "string" || intent.cachedResult.length > 1e3)) return "insertField: bad cachedResult";
      const SAFE_FIELDS = /* @__PURE__ */ new Set([
        "PAGE",
        "NUMPAGES",
        "SECTIONPAGES",
        "SECTION",
        "DATE",
        "TIME",
        "CREATEDATE",
        "SAVEDATE",
        "PRINTDATE",
        "AUTHOR",
        "TITLE",
        "SUBJECT",
        "KEYWORDS",
        "COMMENTS",
        "FILENAME",
        "NUMWORDS",
        "NUMCHARS",
        "PAGEREF",
        "REF",
        "SEQ",
        "STYLEREF",
        "TOC",
        "INDEX",
        "LISTNUM",
        "QUOTE"
      ]);
      const type = intent.instruction.trim().split(/\s+/)[0]?.toUpperCase();
      if (!type || !SAFE_FIELDS.has(type)) return "insertField: field type not allowed";
      if (!/^[\x20-\x7e]+$/.test(intent.instruction)) return "insertField: illegal instruction char";
      return null;
    }
    case "setDrawingRotation":
      if (typeof intent.degrees !== "number" || !Number.isFinite(intent.degrees)) return "setDrawingRotation: bad degrees";
      return null;
    case "setDrawingFill":
      if (intent.color !== null && !/^[0-9a-fA-F]{6}$/.test(intent.color)) return "setDrawingFill: bad color";
      return null;
    case "formatRun":
    case "formatParagraph":
    case "setListType":
    case "tableOp":
    case "mergeParagraph":
      return null;
  }
}

// ../collab/src/session.ts
var UNDO_CLIENT_SEQ_BASE = 1e9;
var DocumentSession = class {
  constructor(doc) {
    this.log = [];
    this.seen = /* @__PURE__ */ new Set();
    /** Base sequence number when the session was seeded from a mid-history
     * checkpoint (E2EE mirrors, doc 13 §3): entries 1..floor are baked into
     * the seed bytes; numbering continues from here. */
    this.seqFloor = 0;
    /**
     * Undo stack per client: EVERY applied intent that client issued, newest
     * last, with its pre-computed inverse — or `null` when the kind has no
     * inverse yet.
     *
     * The null MARKERS are load-bearing, not bookkeeping. Stacking only the
     * invertible intents makes undo mean "your last action THAT HAPPENS TO BE
     * INVERTIBLE", which silently reaches past the others: type "hello", bold
     * it, press undo, and the bold stays while the typing disappears (measured,
     * not theorised — the bold applies and the depth never moves). Recording
     * every action lets undo mean "your last action" and answer honestly when
     * that action cannot be reversed yet.
     */
    this.undoStacks = /* @__PURE__ */ new Map();
    this.undoSeq = /* @__PURE__ */ new Map();
    this.doc = doc;
    this.ids = doc.enableStableIds();
  }
  /** Newest assigned sequence number (0 before any entry). Derived from the
   * last log entry so it stays correct after a snapshot prunes the log
   * prefix (rehydration). */
  get seq() {
    return this.log.length === 0 ? this.seqFloor : this.log[this.log.length - 1].seq;
  }
  /** Declare that this session's document already reflects seqs 1..n (it
   * was seeded from a checkpoint at n) — numbering continues from there. */
  setSeqFloor(n) {
    this.seqFloor = n;
  }
  entriesSince(base) {
    return this.log.filter((e) => e.seq > base);
  }
  /**
   * A checkpoint bundle (plan doc 03): the snapshot docx bytes at the current
   * seq plus the ID sidecar, so a client bootstrapping from it reproduces the
   * exact id table (parse order alone cannot, because split-created nodes
   * carry non-sequential ids).
   */
  checkpoint() {
    return {
      seq: this.seq,
      docx: this.doc.save(),
      sidecar: this.ids.exportSidecar(this.doc.editableRoots())
    };
  }
  /** Install an ID sidecar (from a checkpoint bundle) onto the current
   * document — used when rehydrating from a snapshot so subsequent tail
   * entries resolve their carried ids correctly. */
  installSidecar(sidecar) {
    this.ids.importSidecar(this.doc.editableRoots(), sidecar);
  }
  /**
   * Rehydrate from a persisted log tail whose entries are already sequenced
   * and canonical (their positions were transformed when first submitted).
   * Applies them in order WITHOUT re-transforming, and restores the log,
   * dedup set, and seq. The document passed to the constructor must be the
   * snapshot the tail continues from.
   *
   * NOTE (plan round-2 F1): this reconstructs stable ids by the snapshot's
   * parse order plus the carried ids in split entries. That is exact for
   * histories without pre-snapshot splits; a snapshot taken after splits
   * needs the ID sidecar (documented next step) to reproduce the id table.
   */
  loadCanonical(tail) {
    for (const e of tail) {
      if (e.kind === "applied") {
        const res = applyIntentScoped(this.doc, this.ids, e.intent);
        if (res.applied && res.kind === "split") resyncScope(this.doc, this.ids, res);
        this.seen.add(idempotencyKey(e.intent));
      } else {
        this.seen.add(`${e.clientId}:${e.clientSeq}`);
      }
      this.log.push(e);
    }
  }
  /**
   * Submit an intent from a client. Deduplicates re-sends by idempotency key,
   * transforms against everything sequenced since the intent's base, applies
   * (or rejects), logs, and returns the sequenced entry(ies) to broadcast.
   */
  submit(intent) {
    const key = idempotencyKey(intent);
    if (this.seen.has(key)) {
      const prior = this.log.find(
        (e) => e.kind === "applied" && idempotencyKey(e.intent) === key || e.kind === "rejected" && `${e.clientId}:${e.clientSeq}` === key
      );
      if (prior) return prior;
    }
    this.seen.add(key);
    const invalid = validateIntent(intent);
    if (invalid) return this.reject(intent, invalid);
    if (intent.base < 0 || intent.base > this.seq) {
      return this.reject(intent, "invalid base");
    }
    const ahead = this.log.filter(
      (e) => e.kind === "applied" && e.seq > intent.base && e.intent.clientId !== intent.clientId
    ).map((e) => e.intent);
    const canonical = transformIntent(intent, ahead);
    const inverse = invertIntent(this.doc, this.ids, canonical);
    let result;
    try {
      result = applyIntentScoped(this.doc, this.ids, canonical);
    } catch {
      result = { kind: "doc", applied: false };
    }
    if (!result.applied) return this.reject(intent, "apply failed");
    if (result.kind === "split") resyncScope(this.doc, this.ids, result);
    const entry = { seq: this.seq + 1, kind: "applied", intent: canonical };
    this.log.push(entry);
    if (!this.isUndoIntent(intent)) {
      const stack = this.undoStacks.get(intent.clientId) ?? [];
      stack.push({ seq: entry.seq, inverse });
      this.undoStacks.set(intent.clientId, stack);
    }
    return entry;
  }
  isUndoIntent(intent) {
    return intent.clientSeq >= UNDO_CLIENT_SEQ_BASE;
  }
  /**
   * Take this client's most recent action off its undo stack WITHOUT applying
   * anything — the caller decides how to submit the inverse.
   *
   * This exists because `undo()` below both pops AND submits, and an
   * ENCRYPTED client must never do the second half here: its mirror is a
   * local re-derivation of the canonical log and may only advance by
   * ingesting sequenced envelopes. Applying an inverse directly to the mirror
   * would desynchronise it from every other client's mirror — the one thing
   * the mirror architecture exists to prevent. So the encrypted path takes
   * the inverse, seals it, and submits it as an ordinary intent; it arrives
   * back through ingest like anything else.
   *
   * `cannot-undo` does NOT pop: undo stops at an action it cannot reverse
   * rather than skipping to an older one. That is a deliberate dead end — the
   * cure is implementing more inverses, not quietly undoing something the
   * user did not ask about.
   */
  takeUndo(clientId) {
    const stack = this.undoStacks.get(clientId);
    if (!stack || stack.length === 0) return { kind: "nothing-to-undo" };
    const top = stack[stack.length - 1];
    if (!top.inverse) return { kind: "cannot-undo" };
    stack.pop();
    return { kind: "undoable", seq: top.seq, inverse: top.inverse };
  }
  /** The clientSeq an undo intent must carry so it is not itself stacked.
   * Exposed because the encrypted client builds its own undo intents. */
  nextUndoClientSeq(clientId) {
    const n = (this.undoSeq.get(clientId) ?? 0) + 1;
    this.undoSeq.set(clientId, n);
    return UNDO_CLIENT_SEQ_BASE + n;
  }
  /**
   * Selective per-user undo (plan doc 03 Phase 8): revert the given client's
   * most recent action by submitting its inverse with `base` set to the
   * ORIGINAL intent's seq — the canonical transform then rebases the inverse
   * against everything sequenced since, which is what makes undo correct under
   * concurrency. Returns the resulting log entry, or null when there is
   * nothing to undo OR the last action has no inverse yet (use `takeUndo` when
   * the difference matters to the caller). The inverse is applied through the
   * normal pipeline, so it converges and broadcasts like any intent.
   *
   * PLAINTEXT path only — see takeUndo for why an encrypted client cannot use
   * this method.
   */
  undo(clientId) {
    const candidate = this.takeUndo(clientId);
    if (candidate.kind !== "undoable") return null;
    const undoIntent = {
      ...candidate.inverse,
      clientId,
      clientSeq: this.nextUndoClientSeq(clientId),
      base: candidate.seq
      // rebase the inverse through everything applied since.
    };
    return this.submit(undoIntent);
  }
  /** How many of this client's actions are on its undo stack — including ones
   * whose inverse doesn't exist yet, because they are still actions the user
   * took. Ask `takeUndo` (or peek via `undoState`) for reversibility. */
  undoDepth(clientId) {
    return this.undoStacks.get(clientId)?.length ?? 0;
  }
  /**
   * Do the nodes an inverse addresses still exist in CANONICAL state?
   *
   * A cheap NECESSARY condition, deliberately not a full validation: every
   * inverse this session produces is id-addressed, so a retired id means the
   * target is definitively gone (someone else's delete already sequenced) and
   * the undo cannot possibly land. An encrypted client checks this against
   * its mirror before painting optimistically — declining up front is both
   * honest and flash-free, where submitting would paint the undo and then
   * visibly revert it on the rejection.
   *
   * It does NOT prove the inverse WILL apply: a target that still exists may
   * have changed in ways only the real apply can reject (a text range that no
   * longer spans what it did). Those fall through to the ordinary optimistic
   * rollback every other op already has. Sound about "no", silent about "yes"
   * — which is the right asymmetry for a pre-flight check, since a false
   * "can't" would block a legitimate undo while a false "can" only costs the
   * flash we already tolerate.
   */
  targetsResolve(body) {
    const ids = [];
    const b = body;
    for (const key of ["blockId", "runId", "cellParagraphId"]) {
      if (b[key] !== void 0) ids.push(b[key]);
    }
    const at = b.at;
    if (at) {
      if (at.blockId !== void 0) ids.push(at.blockId);
      if (at.runId !== void 0) ids.push(at.runId);
    }
    return ids.every((id) => typeof id === "number" && this.ids.elOf(id) !== void 0);
  }
  /** What the next undo WOULD do, without consuming anything — for enabling
   * a toolbar button and choosing its tooltip. */
  undoState(clientId) {
    const stack = this.undoStacks.get(clientId);
    if (!stack || stack.length === 0) return "nothing-to-undo";
    return stack[stack.length - 1].inverse ? "undoable" : "cannot-undo";
  }
  /**
   * E2EE mode (doc 13 §2): an envelope that fails to open (garbage from a
   * malicious participant, or any tamper) still consumed a sequence number
   * — every honest client must agree that seq is a no-op, deterministically,
   * because the applied/rejected verdict feeds the transform's `ahead` set.
   * GCM authentication IS deterministic (same bytes fail for everyone), so
   * ingesting the failure as a sequenced rejection keeps all mirrors
   * byte-identical. Client-side use only; the blind server never calls this.
   */
  ingestOpaqueFailure(clientId, clientSeq) {
    this.seen.add(`${clientId}:${clientSeq}`);
    return this.reject({ clientId, clientSeq }, "undecryptable");
  }
  reject(intent, reason) {
    const entry = {
      seq: this.seq + 1,
      kind: "rejected",
      clientId: intent.clientId,
      clientSeq: intent.clientSeq,
      reason
    };
    this.log.push(entry);
    return entry;
  }
};

// ../collab/src/hash.ts
async function docHash(doc) {
  const parts = doc.editableRoots().map((r) => serializeXml(r));
  const digest = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(parts.join("\0")));
  let hex = "";
  for (const b of new Uint8Array(digest)) hex += b.toString(16).padStart(2, "0");
  return hex;
}

// ../collab/src/enc-connection.ts
var _EncryptedCollabConnection = class _EncryptedCollabConnection {
  constructor(transport, clientId, docKey, cb = {}, stretchedCode, selfCheckDelayMs = 400, mediaOpts) {
    this.transport = transport;
    this.clientId = clientId;
    this.docKey = docKey;
    this.cb = cb;
    this.stretchedCode = stretchedCode;
    this.selfCheckDelayMs = selfCheckDelayMs;
    this.mirror = null;
    this.replica = null;
    this.keys = null;
    this.clientSeq = 0;
    this.docId = "";
    this.genesisId = null;
    /** Relay per-blob upload limit from the welcome, or `null` when the server
     * published none (an older host) — in which case callers SKIP the
     * pre-check rather than inventing a limit. See the plaintext connection's
     * field of the same name. */
    this.mediaMaxBlobBytes = null;
    this.docEpoch = 0;
    /** Envelope processing is async (WebCrypto) but MUST be strictly ordered
     * (seq order is the convergence contract) — a serial promise chain. */
    this.queue = Promise.resolve();
    this.idCounter = 0;
    this.idBase = -1;
    /** Own doc hashes at gossip points (seq → hash), a small ring — peers'
     * gossip for a seq we also hashed is comparable; anything else is skipped
     * (clients hash at the SAME seq cadence, so overlap is the common case). */
    this.ownHashes = /* @__PURE__ */ new Map();
    /** Server-designated checkpointer flag (doc 13 §3) + cadence. */
    this.isCheckpointer = false;
    /** The last intent this connection sealed+sent (one-in-flight makes this
     * the only candidate a `stale-base` refusal can refer to). */
    this.lastSent = null;
    /** SELF-HEAL (the B6a typist-drift catch): at quiescence the optimistic
     * replica must equal the mirror's canonical doc — the mirror is LOCAL
     * ground truth (the same authority code the plaintext server runs, fed
     * the identical ordered inputs), so drift in the optimistic
     * rollback-replay machinery is detectable and repairable without a
     * network round-trip. Debounced past the drain; staleness-guarded. */
    this.selfCheckTimer = null;
    this.lastSelfCheckSeq = -1;
    /** Number of times the replica was rebuilt from the mirror (telemetry). */
    this.selfHeals = 0;
    /** STUCK-PENDING RECOVERY (swarm finding): under load a submit can be
     * lost between pending-tracking and delivery (or its echo lost), leaving
     * the client permanently un-quiescent — which also blocks the drift
     * self-check. The front pending is watched; if it makes no progress, the
     * RECONCILIATION-TRANSFORMED copy is re-sealed at the current base and
     * resent (server dedup by (clientId, clientSeq) makes a resend of an
     * already-sequenced op a no-op). After MAX retries the pending is dropped
     * and the replica heals from the mirror so the room converges. */
    this.stuckTimer = null;
    this.stuckKey = null;
    this.stuckRetries = 0;
    /** Rate-limit recovery (see the "refused" handler): one scheduled
     * full-queue re-drive at a time, exponential backoff 300ms→5s, reset
     * when the queue drains. */
    this.redriveTimer = null;
    this.redriveBackoffMs = 0;
    /** Out-of-band media transfer (doc 16 §5), sealed under K_media. */
    this.media = null;
    /** Accumulates doc changes across replica rebuilds (self-heal/welcome), so
     * the getter below stays monotonic per connection. */
    this.docVersionBase = 0;
    /** docVersionBase at the last takeRenderScope — a base bump between takes
     * (welcome/heal doc replacement, media install) has no narrow scope. */
    this.takenDocVersionBase = 0;
    /** Roster + activity parity with the plaintext connection (doc 14). */
    this.roster = [];
    this.activity = [];
    /** Bundle being resumed from (consumed after the welcome replay). */
    this.resuming = null;
    /**
     * Throws on the ingest/submit chain — in practice a seal or transport send
     * that failed AFTER the edit was accepted locally and counted as pending.
     *
     * Separate from droppedPreReady on purpose, and the arithmetic is why: that
     * one counts edits refused BEFORE they were ever applied, this one counts
     * edits lost AFTER. Summing them against a swarm run's "lost" figure is how
     * B13's 231 would have been attributed in minutes instead of a night, and
     * merging the two would have made that sum ambiguous again.
     */
    this.sendFailures = 0;
    /** Submits dropped because the connection was not ready — see
     * ConnectionCallbacks.onSubmitDropped. The window is WIDER here than in the
     * plaintext connection: this one must also derive epoch keys and open the
     * seeded checkpoint before it can honour anything, which on a large
     * document is real time during which the editor's submit path is reachable. */
    this.droppedPreReady = 0;
    /**
     * OUTBOUND presence: a COALESCING single-flight sealer, not a queue.
     *
     * Sealing is async, so the send can no longer be a straight-line call, and
     * three shapes were considered. Fire-and-forget lets two rapid caret moves
     * finish encrypting out of order and paint a stale cursor. Reusing
     * `this.queue` would put a high-frequency, low-value stream behind — and
     * ahead of — broadcast ingest, which is the ordering the document's
     * correctness depends on; perf B10 and B13 were both about work piling up
     * on that chain, and this is that lesson applied before the fact rather
     * than after. Chaining every move onto a private promise fixes ordering but
     * still queues N seals for N twitches of the mouse.
     *
     * So: at most ONE seal in flight, and at most ONE position waiting. A move
     * that arrives while a seal is running REPLACES the waiting one — which is
     * not a compromise but the correct semantics, because presence is
     * last-write-wins and a superseded caret has no reader. Bounded by
     * construction: two slots, whatever the input rate.
     */
    /** Serial chain for INBOUND presence opens (see the receive branch). */
    this.presenceInQueue = Promise.resolve();
    this.presenceLatest = null;
    this.presenceFlushing = false;
    this.transport.onMessage((msg) => this.onServer(msg));
    if (mediaOpts) {
      this.media = new MediaClient(
        mediaOpts,
        () => this.replica?.doc ?? null,
        {
          // Media installs mutate the doc (pixels land) outside the replica's
          // counter — bump the render signal so the skeleton repaints.
          onChange: () => {
            this.docVersionBase++;
            this.cb.onChange?.();
          },
          onState: (part, state) => this.cb.onMediaState?.(part, state),
          need: (sha) => this.transport.send({ t: "media-need", sha }),
          have: (shas) => {
            if (shas.length) this.transport.send({ t: "media-have", shas });
          }
        },
        {
          seal: async (plaintext, iv) => {
            if (!this.keys) throw new Error("media seal before keys");
            return sealMediaBlob(this.keys.kMedia, plaintext, iv);
          },
          open: async (blob, iv) => {
            if (!this.keys) throw new Error("media open before keys");
            return openMediaBlob(this.keys.kMedia, blob, iv);
          }
        }
      );
    }
  }
  /** Upload bytes and get the intent's media fields (blobSha over the
   * CIPHERTEXT, plus the IV the re-supply path must reuse), or null when the
   * relay refused — the caller must not emit an insertImage in that case. */
  async uploadMedia(plaintext) {
    return this.media ? this.media.upload(this.docId, plaintext) : null;
  }
  join(docId, token, opts) {
    this.docId = docId;
    this.transport.send({
      t: "hello",
      protocolVersion: PROTOCOL_VERSION,
      docId,
      clientId: this.clientId,
      takeover: opts?.takeover,
      token,
      sinceSeq: 0,
      profile: opts?.profile,
      codeProof: opts?.codeProof,
      ownerToken: opts?.ownerToken,
      // owner-capability proof (doc 14 §2.5)
      engineVersion: ENGINE_VERSION
      // the fence for client-derived canon (doc 13 §2)
    });
  }
  get doc() {
    return this.replica?.doc ?? null;
  }
  get ready() {
    return this.replica !== null;
  }
  /**
   * Counts changes TO THE RENDERED DOCUMENT — the repaint signal, as opposed
   * to onChange (which also fires for pure bookkeeping: a tracked own echo, a
   * roster-adjacent state change). The editor-driven typing path mutates and
   * paints the doc itself, so its submit + echo leave this untouched; remote
   * applies, optimistic canonical applies (toolbar ops), reloads, heals, and
   * media installs advance it. The react layer repaints on THIS, not on
   * onChange — repainting on onChange queued a whole-document relayout per
   * keystroke, catastrophic past the background-layout page threshold.
   */
  get docVersion() {
    return this.docVersionBase + (this.replica?.docVersion ?? 0);
  }
  /**
   * Drain the dirty scope behind the docVersion movement since the last take
   * — parity with the plaintext connection (see connection.ts). Replica
   * applies carry their per-intent scope; base bumps report document scope;
   * null means nothing was recorded (the caller may skip the repaint).
   */
  takeRenderScope() {
    const replicaScope = this.replica?.takeRenderScope() ?? null;
    if (this.docVersionBase !== this.takenDocVersionBase) {
      this.takenDocVersionBase = this.docVersionBase;
      return { kind: "doc" };
    }
    return replicaScope;
  }
  /** Un-confirmed local intents in flight — drained-replay parity with the
   * plaintext connection (doc 15 §2 / rebase.ts). */
  get pendingCount() {
    return this.replica?.pendingCount ?? 0;
  }
  setProfile(profile) {
    this.transport.send({ t: "profile", profile });
  }
  /**
   * Rejoin from a persisted bundle (doc 12 §5, encrypted flavor). Identical
   * contract to the plaintext connection: clientSeq watermark restored
   * FIRST; pending replayed fire-and-observe on same-epoch welcomes (the
   * sequencer's plaintext-bookkeeping dedup gives exactly-once, and the
   * whole-epoch welcome-enc replay already reconstructed everything that
   * was sequenced pre-crash); different epoch ⇒ onEpochChange, pending
   * withheld (fork rule). The bundle's confirmed bytes are NOT loaded here
   * — encrypted joiners always rebuild from the sealed checkpoint + tail,
   * which is both simpler and verifiable (hash gossip).
   */
  resume(bundle, token, opts) {
    this.clientSeq = Math.max(this.clientSeq, bundle.clientSeq);
    this.resuming = bundle;
    this.join(bundle.docId, token, { ...opts, takeover: true });
  }
  /** Owner admin op (doc 14 §2.5): honored only if this connection proved
   * the owner token at hello — otherwise the server refuses `not-owner`.
   * The admin channel is PLAINTEXT bookkeeping by design (roles are
   * integrity, not confidentiality — doc 14 §2.5 honest limits), so the
   * encrypted connection sends the same frame the plaintext one does.
   * HISTORY: this method was MISSING here while present on the plaintext
   * connection — the react layer's `as unknown as CollabConnection` cast
   * hid it from the typechecker, so the demo's owner controls were silent
   * TypeErrors on encrypted docs (the default), long misattributed to a
   * Vite prebundle race ("stale session w/o admin"). */
  admin(action) {
    this.transport.send({ t: "admin", action });
  }
  /** Durable state as a doc-12 bundle (kDoc rides in it for revival). */
  exportBundle(docId) {
    if (!this.replica || this.genesisId === null) return null;
    return {
      docId,
      genesisId: this.genesisId,
      ...this.replica.exportBundleState(),
      clientSeq: this.clientSeq,
      savedAt: 0,
      lineage: []
      // the persister maintains the chain across writes
    };
  }
  async exportBundleAsync(docId) {
    if (!this.replica || this.genesisId === null) return null;
    return {
      docId,
      genesisId: this.genesisId,
      ...await this.replica.exportBundleStateAsync(),
      clientSeq: this.clientSeq,
      savedAt: 0,
      lineage: []
    };
  }
  /** Same disjoint carried-id allocation as the plaintext connection. */
  allocIds(n) {
    if (this.idBase < 0) {
      let h = 0;
      for (let i = 0; i < this.clientId.length; i++) h = h * 31 + this.clientId.charCodeAt(i) >>> 0;
      this.idBase = 1e9 + h % 1e5 * 1e7;
    }
    const out = [];
    for (let i = 0; i < n; i++) out.push(this.idBase + this.idCounter++);
    return out;
  }
  /** Submit a local edit (optimistic canonical apply, like the plaintext
   * `submit`); sealing happens on the ordered queue. */
  submit(intent) {
    this.submitFull(intent, false);
  }
  /** Editor-driven path: the mutation is already in the live doc. */
  submitPreApplied(intent) {
    this.submitFull(intent, true);
  }
  /**
   * COLLABORATIVE UNDO (plan doc 03 Phase 8), encrypted flavour.
   *
   * The mirror is a real DocumentSession fed the canonical log, so it already
   * knows the inverse of this client's last action — no new wire message and
   * no server involvement are needed. The inverse is submitted as an ORDINARY
   * sealed intent with `base` set to the ORIGINAL action's seq, which is what
   * makes undo correct under concurrency: the canonical transform rebases it
   * through everything sequenced since, and if the target has been deleted or
   * changed the apply rejects it identically on every replica.
   *
   * Two things this deliberately does NOT do. It does not call the mirror's
   * `undo()`, because that would apply to the mirror out of band and the
   * mirror may only advance by ingesting sequenced envelopes. And it does not
   * paint an undo it can already prove is dead: `targetsResolve` checks the
   * canonical id table first, so the common conflict (someone else removed
   * the thing) is declined up front rather than painted and yanked back.
   */
  undoLast() {
    if (!this.mirror || !this.replica || !this.keys || !this.genesisId) return "unavailable";
    const candidate = this.mirror.takeUndo(this.clientId);
    if (candidate.kind !== "undoable") return candidate.kind;
    if (!this.mirror.targetsResolve(candidate.inverse)) return "changed-since";
    const ahead = this.mirror.entriesSince(candidate.seq).filter((e) => e.kind === "applied" && e.intent.clientId !== this.clientId).map((e) => e.intent);
    const rebased = transformIntent(
      {
        ...candidate.inverse,
        clientId: this.clientId,
        // Reserved range: keeps the undo itself off every replica's undo stack.
        clientSeq: this.mirror.nextUndoClientSeq(this.clientId),
        base: candidate.seq
      },
      ahead
    );
    const full = { ...rebased, base: this.replica.confirmedSeq };
    this.replica.submitLocal(full);
    this.enqueue(async () => {
      try {
        this.lastSent = full;
        const env = await sealIntent(this.keys.kContent, this.docId, this.genesisId, full);
        this.transport.send({ t: "submit-enc", envelope: env });
      } catch (err) {
        this.sendFailures++;
        this.cb.onError?.({ where: "enc.submit", error: err });
      }
    });
    this.scheduleSelfCheck();
    this.cb.onChange?.();
    return "undone";
  }
  /** What the next undo WOULD do, for enabling a button and its tooltip. */
  undoState() {
    if (!this.mirror) return "unavailable";
    const state = this.mirror.undoState(this.clientId);
    return state === "undoable" ? "undone" : state;
  }
  submitFull(intent, preApplied) {
    if (!this.replica || !this.keys || !this.genesisId) {
      this.droppedPreReady++;
      this.cb.onSubmitDropped?.("not-ready");
      return;
    }
    const full = {
      ...intent,
      clientId: this.clientId,
      clientSeq: ++this.clientSeq,
      base: this.replica.confirmedSeq
    };
    if (preApplied) this.replica.trackLocal(full);
    else this.replica.submitLocal(full);
    this.enqueue(async () => {
      try {
        this.lastSent = full;
        const env = await sealIntent(this.keys.kContent, this.docId, this.genesisId, full);
        this.transport.send({ t: "submit-enc", envelope: env });
      } catch (err) {
        this.sendFailures++;
        this.cb.onError?.({ where: "enc.submit", error: err });
      }
    });
    this.scheduleSelfCheck();
    this.cb.onChange?.();
  }
  setPresence(position) {
    const clean = sanitizePresencePosition(position);
    const keys = this.keys;
    const genesisId = this.genesisId;
    if (!keys || !genesisId) {
      return;
    }
    this.presenceLatest = { position: clean };
    if (this.presenceFlushing) return;
    this.presenceFlushing = true;
    void (async () => {
      try {
        while (this.presenceLatest) {
          const next = this.presenceLatest;
          this.presenceLatest = null;
          const sealed = await sealPresence(keys.kPresence, this.docId, genesisId, this.clientId, next.position);
          this.transport.send({ t: "presence", position: sealed });
        }
      } catch {
      } finally {
        this.presenceFlushing = false;
      }
    })();
  }
  enqueue(task) {
    this.queue = this.queue.then(task).catch((err) => {
      this.sendFailures++;
      this.cb.onError?.({ where: "enc.queue", error: err });
    });
  }
  onServer(msg) {
    switch (msg.t) {
      case "welcome-enc": {
        this.mediaMaxBlobBytes = msg.mediaMaxBlobBytes ?? null;
        this.enqueue(async () => {
          const perfOn = !!globalThis.__dxwPerf;
          const t = {};
          let t0 = performance.now();
          this.genesisId = msg.genesisId;
          this.keys = await deriveEpochKeys(this.docKey, msg.genesisId, this.stretchedCode);
          t.deriveMs = performance.now() - t0;
          t0 = performance.now();
          const cp = await openCheckpoint(this.keys.kContent, msg.docId, msg.genesisId, msg.checkpoint.seq, msg.checkpoint);
          const bytes = b64ToBytes(cp.docx);
          t.openMs = performance.now() - t0;
          t0 = performance.now();
          this.mirror = new DocumentSession(DocxDocument.load(bytes));
          this.mirror.setSeqFloor(msg.checkpoint.seq);
          if (cp.sidecar) this.mirror.installSidecar(cp.sidecar);
          t.mirrorParseMs = performance.now() - t0;
          t0 = performance.now();
          this.docVersionBase += (this.replica?.docVersion ?? 0) + 1;
          this.replica = new ClientReplica(bytes, cp.sidecar ?? void 0);
          this.replica.confirmedSeq = msg.checkpoint.seq;
          t.replicaParseMs = performance.now() - t0;
          t0 = performance.now();
          applyMediaAddresses(this.replica.doc, cp.mediaMeta);
          applyMediaAddresses(this.mirror.doc, cp.mediaMeta);
          for (const env of msg.tail) await this.ingest(env);
          t.tailMs = performance.now() - t0;
          if (perfOn) {
            console.log(
              `STRESS-METRIC join-welcome docxBytes=${bytes.byteLength} tail=${msg.tail.length} ` + Object.entries(t).map(([k, v]) => `${k}=${v.toFixed(2)}`).join(" ")
            );
          }
          const resumed = this.resuming;
          this.resuming = null;
          if (resumed) {
            if (resumed.genesisId === msg.genesisId) {
              for (const intent of resumed.pending) {
                const env2 = await sealIntent(this.keys.kContent, this.docId, msg.genesisId, intent);
                this.transport.send({ t: "submit-enc", envelope: env2 });
              }
            } else {
              this.cb.onEpochChange?.(resumed.genesisId, msg.genesisId);
            }
          }
          this.cb.onChange?.();
          if (resumed?.mediaMeta) {
            const held = DocxDocument.load(resumed.confirmedBytes);
            for (const [part, meta] of resumed.mediaMeta) {
              this.replica.doc.mediaMeta.set(part, meta);
              const pixels = held.media(part);
              if (pixels && this.replica.doc.mediaStatus(part) === "pending") {
                this.replica.doc.installMedia(part, pixels);
              }
            }
          }
          this.media?.volunteer(msg.mediaNeeded);
          void this.media?.fetchPending(this.docId);
        });
        return;
      }
      case "broadcast-enc": {
        this.enqueue(async () => {
          if (!this.replica) return;
          for (const env of msg.entries) await this.ingest(env);
          if (this.replica.reloaded) this.docEpoch++;
          this.cb.onChange?.();
          this.scheduleSelfCheck();
          void this.media?.fetchPending(this.docId);
        });
        return;
      }
      case "media-request": {
        this.media?.answerRequest(msg.sha);
        this.cb.onMediaRequest?.(msg.sha);
        return;
      }
      case "media-upload": {
        void this.media?.resupply(this.docId, msg.sha);
        this.cb.onMediaRequest?.(msg.sha);
        return;
      }
      case "media-ready": {
        void this.media?.onReady(this.docId, msg.sha);
        this.cb.onMediaReady?.(msg.sha);
        return;
      }
      case "media-unavailable": {
        this.media?.onUnavailable(msg.sha);
        this.cb.onMediaUnavailable?.(msg.sha);
        return;
      }
      case "welcome": {
        this.cb.onRefused?.("mode-downgrade");
        return;
      }
      case "presence": {
        const raw = msg.position;
        if (raw === null) {
          this.cb.onPresence?.(msg.participant, null);
          return;
        }
        if (!isSealedPresence(raw)) {
          return;
        }
        const keys = this.keys;
        const genesisId = this.genesisId;
        if (!keys || !genesisId) return;
        this.presenceInQueue = this.presenceInQueue.then(async () => {
          const opened = await openPresence(keys.kPresence, this.docId, genesisId, msg.participant, raw);
          this.cb.onPresence?.(msg.participant, sanitizePresencePosition(opened));
        }).catch(() => {
        });
        return;
      }
      case "roster": {
        this.roster = msg.roster;
        this.cb.onRoster?.(msg.roster);
        return;
      }
      // Session lifecycle warnings are PLAINTEXT bookkeeping and identical in
      // both modes: the deadline is the server's own policy, not document
      // content, so a blind sequencer announces it in the clear. Dispatched
      // directly rather than through the serial queue — nothing here touches
      // keys, the mirror, or the replica.
      case "session-warning": {
        this.cb.onSessionWarning?.({ reason: msg.reason, inMs: msg.inMs });
        return;
      }
      case "session-warning-cleared": {
        this.cb.onSessionWarningCleared?.({ reason: msg.reason });
        return;
      }
      case "gossip": {
        this.enqueue(() => this.onGossip(msg.iv, msg.ciphertext));
        return;
      }
      case "checkpointer": {
        this.isCheckpointer = msg.active;
        return;
      }
      case "refused": {
        if (msg.reason === "stale-base") {
          const retry = this.lastSent;
          if (retry && this.replica && this.keys && this.genesisId) {
            this.enqueue(async () => {
              const body = this.replica.pendingIntent(retry.clientId, retry.clientSeq);
              if (!body) return;
              const rebased = { ...body, base: this.replica.confirmedSeq };
              const env = await sealIntent(this.keys.kContent, this.docId, this.genesisId, rebased);
              this.transport.send({ t: "submit-enc", envelope: env });
            });
          }
          return;
        }
        if (msg.reason === "rate-limit") {
          this.stuckRetries = 0;
          this.scheduleRateLimitRedrive();
          return;
        }
        this.cb.onRefused?.(msg.reason);
        return;
      }
    }
  }
  /** Re-drive the whole pending queue after rate-limit backoff. One timer
   * at a time; backoff doubles per refusal and resets when the queue
   * drains (see scheduleSelfCheck). */
  scheduleRateLimitRedrive() {
    if (this.redriveTimer !== null) return;
    this.redriveBackoffMs = Math.min(this.redriveBackoffMs === 0 ? 300 : this.redriveBackoffMs * 2, 5e3);
    this.redriveTimer = setTimeout(() => {
      this.redriveTimer = null;
      this.enqueue(async () => {
        if (!this.replica || !this.keys || !this.genesisId) return;
        for (const body of this.replica.pendingCopies()) {
          const rebased = { ...body, base: this.replica.confirmedSeq };
          const env = await sealIntent(this.keys.kContent, this.docId, this.genesisId, rebased);
          this.transport.send({ t: "submit-enc", envelope: env });
        }
      });
    }, this.redriveBackoffMs);
  }
  /** Debounce a quiescent self-check: only when nothing is pending (a burst
   * still in flight legitimately diverges live from canonical), re-armed by
   * every broadcast so the check runs once the wire goes quiet. While
   * pending exists, arm the stuck-pending watchdog instead. */
  scheduleSelfCheck() {
    if (!this.replica) return;
    if (this.replica.pendingCount !== 0) {
      this.scheduleStuckCheck();
      return;
    }
    this.redriveBackoffMs = 0;
    if (this.redriveTimer !== null) {
      clearTimeout(this.redriveTimer);
      this.redriveTimer = null;
    }
    if (this.stuckTimer !== null) {
      clearTimeout(this.stuckTimer);
      this.stuckTimer = null;
    }
    this.stuckKey = null;
    this.stuckRetries = 0;
    if (this.selfCheckTimer !== null) clearTimeout(this.selfCheckTimer);
    this.selfCheckTimer = setTimeout(() => {
      this.selfCheckTimer = null;
      this.enqueue(() => this.selfCheck());
    }, this.selfCheckDelayMs);
  }
  /** Watch the FRONT pending op; progress (a different front, or empty)
   * resets the retry budget. The window is 5× the self-check debounce. */
  scheduleStuckCheck() {
    const front = this.replica?.firstPending();
    if (!front) return;
    const key = `${front.clientId}:${front.clientSeq}`;
    if (this.stuckKey !== key) {
      this.stuckKey = key;
      this.stuckRetries = 0;
    }
    if (this.stuckTimer !== null) clearTimeout(this.stuckTimer);
    this.stuckTimer = setTimeout(() => {
      this.stuckTimer = null;
      this.enqueue(() => this.stuckCheck(front.clientId, front.clientSeq));
    }, this.selfCheckDelayMs * 5);
  }
  /** Runs inside the serial queue: if the same front pending is still
   * stuck, resend its transformed copy at the current base; past the retry
   * budget, drop it and heal from the mirror (convergence over delivery). */
  async stuckCheck(clientId, clientSeq) {
    if (!this.replica || !this.mirror || !this.keys || !this.genesisId) return;
    const front = this.replica.firstPending();
    if (!front || front.clientId !== clientId || front.clientSeq !== clientSeq) {
      this.scheduleSelfCheck();
      return;
    }
    if (this.stuckRetries >= _EncryptedCollabConnection.STUCK_MAX_RETRIES) {
      const liveHash = await docHash(this.replica.doc);
      const canonicalHash = await docHash(this.mirror.doc);
      const cp = this.mirror.checkpoint();
      this.docVersionBase += this.replica.docVersion + 1;
      this.replica = new ClientReplica(cp.docx, cp.sidecar);
      this.replica.confirmedSeq = cp.seq;
      this.docEpoch++;
      this.selfHeals++;
      this.stuckKey = null;
      this.stuckRetries = 0;
      this.cb.onSelfHeal?.({ seq: cp.seq, liveHash, canonicalHash });
      this.cb.onChange?.();
      return;
    }
    this.stuckRetries++;
    const body = this.replica.pendingIntent(clientId, clientSeq);
    if (body) {
      const rebased = { ...body, base: this.replica.confirmedSeq };
      const env = await sealIntent(this.keys.kContent, this.docId, this.genesisId, rebased);
      this.transport.send({ t: "submit-enc", envelope: env });
    }
    this.scheduleStuckCheck();
  }
  /** Runs INSIDE the serial queue (no ingest can interleave): compare the
   * optimistic replica against the mirror at quiescence; on drift, rebuild
   * the replica from the mirror — the same shape as a reload heal, but
   * automatic, local, and scoped to the one client that drifted. */
  async selfCheck() {
    if (!this.replica || !this.mirror) return;
    if (this.replica.pendingCount !== 0) return;
    const seq = this.mirror.seq;
    if (seq === this.lastSelfCheckSeq) return;
    const liveHash = await docHash(this.replica.doc);
    const canonicalHash = await docHash(this.mirror.doc);
    if (this.replica.pendingCount !== 0 || this.mirror.seq !== seq) return;
    this.lastSelfCheckSeq = seq;
    if (liveHash === canonicalHash) return;
    const cp = this.mirror.checkpoint();
    this.docVersionBase += this.replica.docVersion + 1;
    this.replica = new ClientReplica(cp.docx, cp.sidecar);
    this.replica.confirmedSeq = cp.seq;
    this.docEpoch++;
    this.selfHeals++;
    this.cb.onSelfHeal?.({ seq, liveHash, canonicalHash });
    this.cb.onChange?.();
  }
  /** Feed one sequenced envelope through mirror + replica (in seq order). */
  async ingest(env) {
    let entry;
    try {
      const intent = await openIntent(this.keys.kContent, this.docId, this.genesisId, env);
      entry = this.mirror.submit(intent);
    } catch {
      entry = this.mirror.ingestOpaqueFailure(env.clientId, env.clientSeq);
    }
    if (entry.seq !== env.seq) {
      this.cb.onRefused?.("sequence-desync");
      return;
    }
    this.replica.receive([entry]);
    if (entry.kind === "rejected" && entry.clientId === this.clientId) {
      this.cb.onIntentRejected?.({ reason: entry.reason, clientSeq: entry.clientSeq });
    }
    if (entry.kind === "applied") {
      this.activity.push({ seq: entry.seq, clientId: entry.intent.clientId, kind: entry.intent.kind });
      if (this.activity.length > 100) this.activity.splice(0, this.activity.length - 100);
    }
    if (this.isCheckpointer && env.seq % _EncryptedCollabConnection.CHECKPOINT_EVERY === 0) {
      const cp = this.mirror.checkpoint();
      const sealed = await sealCheckpoint(this.keys.kContent, this.docId, this.genesisId, cp.seq, {
        docx: bytesToB64(cp.docx),
        sidecar: cp.sidecar,
        docHash: await docHash(this.mirror.doc),
        // Doc 16 §6 late-join, sealed half: the MIRROR never installs pixels,
        // but it holds every REGISTRATION — which is all a joiner needs to go
        // and fetch them. Inside the sealed body, so a blind server still
        // learns nothing about the document's part structure.
        mediaMeta: mediaAddressesOf(this.mirror.doc)
      });
      this.transport.send({ t: "checkpoint", checkpoint: { seq: cp.seq, ...sealed } });
    }
    if (env.seq % _EncryptedCollabConnection.GOSSIP_EVERY === 0) {
      const hash = await docHash(this.mirror.doc);
      this.ownHashes.set(env.seq, hash);
      if (this.ownHashes.size > 5) {
        this.ownHashes.delete(Math.min(...this.ownHashes.keys()));
      }
      const iv = new Uint8Array(12);
      crypto.getRandomValues(iv);
      const body = new TextEncoder().encode(JSON.stringify({ seq: env.seq, hash }));
      const aad = new TextEncoder().encode(`gs:${this.docId}:${this.genesisId}`);
      const ct = await crypto.subtle.encrypt(
        { name: "AES-GCM", iv, additionalData: aad },
        this.keys.kContent,
        body
      );
      this.transport.send({ t: "gossip", iv: bytesToB64(iv), ciphertext: bytesToB64(new Uint8Array(ct)) });
    }
  }
  /** Verify a peer's gossiped hash against our own at the same seq. */
  async onGossip(iv, ciphertext) {
    if (!this.keys || !this.genesisId) return;
    let claim;
    try {
      const aad = new TextEncoder().encode(`gs:${this.docId}:${this.genesisId}`);
      const pt = await crypto.subtle.decrypt(
        { name: "AES-GCM", iv: b64ToBytes(iv), additionalData: aad },
        this.keys.kContent,
        b64ToBytes(ciphertext)
      );
      claim = JSON.parse(new TextDecoder().decode(pt));
    } catch {
      return;
    }
    const own = this.ownHashes.get(claim.seq);
    if (own !== void 0 && own !== claim.hash) {
      this.cb.onRefused?.("divergence");
    }
  }
};
/** Every K ingested seqs, hash + gossip (doc 13 §2). */
_EncryptedCollabConnection.GOSSIP_EVERY = 20;
_EncryptedCollabConnection.CHECKPOINT_EVERY = 50;
_EncryptedCollabConnection.STUCK_MAX_RETRIES = 3;
var EncryptedCollabConnection = _EncryptedCollabConnection;

// ../collab/src/ws-transport.ts
function createWebSocketTransport(socket) {
  let open = false;
  const backlog = [];
  socket.addEventListener("open", () => {
    open = true;
    for (const frame of backlog.splice(0)) socket.send(frame);
  });
  let handler = null;
  socket.addEventListener("message", (ev) => {
    let msg;
    try {
      msg = JSON.parse(typeof ev.data === "string" ? ev.data : String(ev.data));
    } catch {
      return;
    }
    handler?.(msg);
  });
  return {
    send(msg) {
      const frame = JSON.stringify(msg);
      if (open) socket.send(frame);
      else backlog.push(frame);
    },
    onMessage(cb) {
      handler = cb;
    }
  };
}

// ../collab/src/liveness.ts
var DEFAULT_INTERVAL_MS = 15e3;
var DEFAULT_TIMEOUT_MS = 1e4;
var LivenessMonitor = class {
  constructor(send, onDead, opts = {}) {
    this.send = send;
    this.onDead = onDead;
    this.opts = opts;
    this.probeTimer = null;
    this.deadlineTimer = null;
    this.nonce = 0;
    /** The nonce of the probe currently awaiting an answer; null when idle. */
    this.awaiting = null;
    this.stopped = false;
  }
  get intervalMs() {
    return this.opts.intervalMs ?? DEFAULT_INTERVAL_MS;
  }
  get timeoutMs() {
    return this.opts.timeoutMs ?? DEFAULT_TIMEOUT_MS;
  }
  setTimer(fn, ms) {
    return (this.opts.setTimer ?? ((f, m) => setTimeout(f, m)))(fn, ms);
  }
  clearTimer(t) {
    if (t !== null) (this.opts.clearTimer ?? clearTimeout)(t);
  }
  /** Begin probing. Call when the socket opens. */
  start() {
    if (this.stopped) return;
    this.arm();
  }
  /** Stop every timer and answer nothing further. Idempotent — the reconnect
   * path calls it from both the close handler and the effect teardown. */
  stop() {
    this.stopped = true;
    this.clearTimer(this.probeTimer);
    this.clearTimer(this.deadlineTimer);
    this.probeTimer = null;
    this.deadlineTimer = null;
    this.awaiting = null;
  }
  /** Schedule the next probe one interval out. */
  arm() {
    if (this.stopped) return;
    this.clearTimer(this.probeTimer);
    this.probeTimer = this.setTimer(() => {
      this.probeTimer = null;
      this.probe();
    }, this.intervalMs);
  }
  /**
   * Probe NOW rather than waiting for the interval — wired to
   * `visibilitychange`, because a tab waking from sleep is the exact moment
   * the connection is most likely already dead and the worst moment to make
   * the user wait a full interval to find out. Without this the motivating
   * scenario (unlock phone, start typing) spends up to 25s in the silent-loss
   * state the arc exists to close.
   *
   * A probe already in flight is left alone: re-arming its deadline on every
   * focus event would let a user who keeps switching tabs postpone the verdict
   * indefinitely.
   */
  probe() {
    if (this.stopped || this.awaiting !== null) return;
    const nonce = ++this.nonce;
    this.awaiting = nonce;
    this.deadlineTimer = this.setTimer(() => {
      this.deadlineTimer = null;
      if (this.stopped || this.awaiting !== nonce) return;
      this.stop();
      this.onDead();
    }, this.timeoutMs);
    try {
      this.send({ t: "ping", nonce });
    } catch {
    }
  }
  /**
   * Feed an inbound frame to the monitor.
   *
   * ONLY A MATCHING PONG COUNTS AS PROOF OF LIFE — deliberately, and this is
   * the subtle decision in the file. The obvious optimisation is to treat any
   * inbound traffic as evidence the connection works, which is what most
   * heartbeats do. It is wrong here: it assumes both directions fail together.
   * A connection whose OUTBOUND half is broken while broadcasts still arrive
   * would look permanently healthy under that rule, and that is precisely the
   * user-visible bug being fixed — edits going nowhere while the screen keeps
   * updating with everyone else's. A pong proves the full round trip an edit
   * actually needs, so nothing weaker is accepted.
   */
  noteInbound(msg) {
    if (this.stopped) return;
    if (msg.t !== "pong") return;
    if (this.awaiting === null || msg.nonce !== this.awaiting) return;
    this.awaiting = null;
    this.clearTimer(this.deadlineTimer);
    this.deadlineTimer = null;
    this.arm();
  }
};
function monitorTransport(inner, onDead, opts = {}) {
  const monitor = new LivenessMonitor((msg) => inner.send(msg), onDead, opts);
  const transport = {
    send: (msg) => inner.send(msg),
    onMessage: (cb) => {
      inner.onMessage((msg) => {
        monitor.noteInbound(msg);
        if (msg.t === "pong") return;
        cb(msg);
      });
    }
  };
  return { transport, monitor };
}

// src/bundle-store.ts
var LEGACY_BACKFILL_PER_LIST = 4;
var IndexedDbBundleStore = class {
  constructor(dbName = "wordinweb-bundles") {
    this.dbName = dbName;
    this.db = null;
    this.persistRequested = false;
  }
  open() {
    if (!this.db) {
      this.db = new Promise((resolve, reject) => {
        const req = indexedDB.open(this.dbName, 2);
        req.onupgradeneeded = () => {
          const db = req.result;
          if (!db.objectStoreNames.contains("bundles")) db.createObjectStore("bundles", { keyPath: "docId" });
          if (!db.objectStoreNames.contains("meta")) db.createObjectStore("meta", { keyPath: "docId" });
        };
        req.onsuccess = () => resolve(req.result);
        req.onerror = () => reject(req.error);
      });
    }
    return this.db;
  }
  /** One request in one transaction, as a promise. */
  tx(stores, mode, run) {
    return this.open().then(
      (db) => new Promise((resolve, reject) => {
        const t = db.transaction(stores, mode);
        const req = run(t);
        req.onsuccess = () => resolve(req.result);
        req.onerror = () => reject(req.error);
        t.onabort = () => reject(t.error ?? new Error("IndexedDB transaction aborted"));
      })
    );
  }
  async get(docId) {
    return await this.tx("bundles", "readonly", (t) => t.objectStore("bundles").get(docId)) ?? null;
  }
  async put(bundle) {
    if (!this.persistRequested) {
      this.persistRequested = true;
      void globalThis.navigator?.storage?.persist?.();
    }
    const db = await this.open();
    await new Promise((resolve, reject) => {
      const t = db.transaction(["bundles", "meta"], "readwrite");
      t.objectStore("bundles").put(bundle);
      const meta = { docId: bundle.docId, savedAt: bundle.savedAt, byteLength: bundle.confirmedBytes.byteLength };
      t.objectStore("meta").put(meta);
      t.oncomplete = () => resolve();
      t.onerror = () => reject(t.error);
      t.onabort = () => reject(t.error ?? new Error("IndexedDB transaction aborted"));
    });
  }
  async delete(docId) {
    const db = await this.open();
    await new Promise((resolve, reject) => {
      const t = db.transaction(["bundles", "meta"], "readwrite");
      t.objectStore("bundles").delete(docId);
      t.objectStore("meta").delete(docId);
      t.oncomplete = () => resolve();
      t.onerror = () => reject(t.error);
      t.onabort = () => reject(t.error ?? new Error("IndexedDB transaction aborted"));
    });
  }
  async list() {
    const keys = await this.tx("bundles", "readonly", (t) => t.objectStore("bundles").getAllKeys());
    const metas = await this.tx("meta", "readonly", (t) => t.objectStore("meta").getAll());
    const byKey = new Map(metas.map((m) => [m.docId, m]));
    const out = [];
    let budget = LEGACY_BACKFILL_PER_LIST;
    for (const key of keys) {
      let meta = byKey.get(key);
      if (!meta && budget > 0) {
        budget--;
        const bundle = await this.get(key);
        if (!bundle) continue;
        meta = { docId: key, savedAt: bundle.savedAt, byteLength: bundle.confirmedBytes.byteLength };
        await this.tx("meta", "readwrite", (t) => t.objectStore("meta").put(meta));
      }
      out.push({
        ...parseBundleKey(key),
        key,
        // Unmeasured legacy record: 0 means UNKNOWN. A byte-budget caller must
        // not treat it as free, and must not delete it for being large either.
        savedAt: meta?.savedAt ?? 0,
        byteLength: meta?.byteLength ?? 0
      });
    }
    return out;
  }
};

// src/collab.tsx
var TIMED_OUT = /* @__PURE__ */ Symbol("store-timeout");
var STORE_DEADLINE_MS = 2e3;
function withDeadline(p, ms, fallback) {
  return new Promise((resolve) => {
    let done = false;
    const t = setTimeout(() => {
      if (!done) {
        done = true;
        resolve(fallback);
      }
    }, ms);
    void p.then(
      (v) => {
        if (!done) {
          done = true;
          clearTimeout(t);
          resolve(v);
        }
      },
      () => {
        if (!done) {
          done = true;
          clearTimeout(t);
          resolve(fallback);
        }
      }
    );
  });
}
function useCollab(opts) {
  const { url, docId, clientId, token, createSocket, store, profile, takeover, docKey, shareCode, ownerToken, httpBase } = opts;
  const connRef = useRef(null);
  const [doc, setDoc] = useState(null);
  const [version, setVersion] = useState(0);
  const [renderVersion, setRenderVersion] = useState(0);
  const lastDocVersionRef = useRef(0);
  const renderScopeRef = useRef(null);
  const [docEpoch, setDocEpoch] = useState(0);
  const [ready, setReady] = useState(false);
  const [presence, setPresence] = useState({});
  const [refused, setRefused] = useState(null);
  const [deadlines, setDeadlines] = useState({ idle: null, lifetime: null });
  const clearDeadlines = () => setDeadlines((d) => d.idle === null && d.lifetime === null ? d : { idle: null, lifetime: null });
  const sessionWarning = useMemo(() => {
    const { idle, lifetime } = deadlines;
    const reason = idle !== null && (lifetime === null || idle <= lifetime) ? "idle" : lifetime !== null ? "lifetime" : null;
    if (!reason) return null;
    return { reason, inMs: Math.max(0, (reason === "idle" ? idle : lifetime) - Date.now()) };
  }, [deadlines]);
  const [readOnlyBlocked, setReadOnlyBlocked] = useState(false);
  const [notOwner, setNotOwner] = useState(false);
  const [epochChanged, setEpochChanged] = useState(null);
  const [roster, setRoster] = useState([]);
  const [arrival, setArrival] = useState(null);
  const [selfHeals, setSelfHeals] = useState(0);
  const [droppedPreReady, setDroppedPreReady] = useState(0);
  const [persistErrors, setPersistErrors] = useState(0);
  const [storeSlow, setStoreSlow] = useState(false);
  const [sendFailures, setSendFailures] = useState(0);
  const offlineTailRef = useRef(void 0);
  const offlineEpochRef = useRef(null);
  const [offlineHeld, setOfflineHeld] = useState(0);
  const lastGenesisRef = useRef(null);
  const persisterRef = useRef(null);
  const tailConnRef = useRef(null);
  const offlinePersistTimerRef = useRef(null);
  const [connection, setConnection] = useState("live");
  const [reconnectNonce, setReconnectNonce] = useState(0);
  const backoffRef = useRef(0);
  const attemptRef = useRef(0);
  const retryTimerRef = useRef(null);
  const flushDoneRef = useRef(Promise.resolve());
  const maxRetries = opts.liveness?.maxRetries ?? 6;
  const livenessRef = useRef(opts.liveness);
  livenessRef.current = opts.liveness;
  const reconnect = () => {
    if (retryTimerRef.current !== null) clearTimeout(retryTimerRef.current);
    retryTimerRef.current = null;
    backoffRef.current = 0;
    attemptRef.current = 0;
    setConnection("reconnecting");
    setReconnectNonce((n) => n + 1);
  };
  const persistOfflineTail = () => {
    if (!store) return;
    flushDoneRef.current = flushDoneRef.current.then(async () => {
      const prev = await store.get(docId);
      if (!prev) return;
      const tail = offlineTailRef.current ?? [];
      const next = { ...prev };
      if (tail.length) {
        next.offlineTail = [...tail];
        next.offlineTailEpoch = offlineEpochRef.current ?? void 0;
      } else {
        delete next.offlineTail;
        delete next.offlineTailEpoch;
      }
      await store.put(next);
    }).catch((err) => {
      setPersistErrors((n) => n + 1);
      console.error("[wordinweb] offline-tail-persist", err);
    });
  };
  const scheduleOfflinePersist = () => {
    if (!store || offlinePersistTimerRef.current !== null) return;
    offlinePersistTimerRef.current = setTimeout(() => {
      offlinePersistTimerRef.current = null;
      persistOfflineTail();
    }, 1e3);
  };
  const flushOfflinePersist = () => {
    if (offlinePersistTimerRef.current === null) return;
    clearTimeout(offlinePersistTimerRef.current);
    offlinePersistTimerRef.current = null;
    persistOfflineTail();
  };
  const captureOffline = (intent, preApplied) => {
    if (!preApplied) {
      if (!doc) return;
      const ids = doc.enableStableIds();
      const res = applyIntentScoped(doc, ids, intent);
      if (!res.applied) return;
      resyncScope(doc, ids, res);
      renderScopeRef.current = unionScopes(renderScopeRef.current, res);
      setRenderVersion((v) => v + 1);
    }
    const tail = offlineTailRef.current ?? [];
    tail.push(intent);
    offlineTailRef.current = tail;
    offlineEpochRef.current ?? (offlineEpochRef.current = lastGenesisRef.current);
    setOfflineHeld(tail.length);
    setVersion((v) => v + 1);
    if (connRef.current?.ready) persisterRef.current?.notify();
    else scheduleOfflinePersist();
  };
  const runTailReplay = (conn, kind) => {
    const author = profile?.name || clientId;
    const date = (/* @__PURE__ */ new Date()).toISOString();
    const step = () => {
      if (connRef.current !== conn || !conn.ready) return;
      if (conn.pendingCount > 0) {
        setTimeout(step, 25);
        return;
      }
      const tail = offlineTailRef.current ?? [];
      if (!tail.length) {
        offlineTailRef.current = void 0;
        offlineEpochRef.current = null;
        setOfflineHeld(0);
        setArrival(null);
        return;
      }
      const next = tail[0];
      offlineTailRef.current = tail.slice(1);
      setOfflineHeld(tail.length - 1);
      if (kind === "suggest") {
        const { suggestions } = toSuggestions([next], author, date);
        if (suggestions.length) conn.submit(suggestions[0]);
      } else {
        conn.submit(next);
      }
      void persisterRef.current?.flush();
      setTimeout(step, 0);
    };
    step();
  };
  useEffect(() => {
    clearDeadlines();
    const socket = (createSocket ?? ((u) => new WebSocket(u)))(url);
    let disposed = false;
    let dropHandled = false;
    let monitor = null;
    const dropped = () => {
      if (dropHandled || disposed) return;
      dropHandled = true;
      monitor?.stop();
      try {
        socket.close?.();
      } catch {
      }
      if (attemptRef.current >= maxRetries) {
        setConnection("lost");
        return;
      }
      attemptRef.current += 1;
      backoffRef.current = Math.min(backoffRef.current === 0 ? 300 : backoffRef.current * 2, 5e3);
      setConnection("reconnecting");
      retryTimerRef.current = setTimeout(() => {
        retryTimerRef.current = null;
        setReconnectNonce((n) => n + 1);
      }, backoffRef.current);
    };
    socket.onclose = () => dropped();
    socket.onerror = () => dropped();
    const monitored = monitorTransport(
      createWebSocketTransport(socket),
      () => dropped(),
      livenessRef.current ?? {}
    );
    const transport = monitored.transport;
    monitor = monitored.monitor;
    const onVisible = () => {
      if (typeof document !== "undefined" && document.visibilityState !== "visible") return;
      monitor?.probe();
    };
    if (typeof document !== "undefined") document.addEventListener("visibilitychange", onVisible);
    monitor.start();
    let persister = null;
    const callbacks = {
      onChange: () => {
        const c = connRef.current;
        if (!c) return;
        setDoc(c.doc);
        setReady(c.ready);
        setVersion((v) => v + 1);
        const dv = c.docVersion;
        if (dv !== lastDocVersionRef.current) {
          lastDocVersionRef.current = dv;
          setRenderVersion((v) => v + 1);
        }
        setDocEpoch(c.docEpoch);
        persister?.notify();
        if (c.ready) {
          backoffRef.current = 0;
          attemptRef.current = 0;
          setConnection((s) => s === "live" ? s : "live");
          lastGenesisRef.current = c.genesisId;
          if (tailConnRef.current !== c) {
            tailConnRef.current = c;
            const tail = offlineTailRef.current ?? [];
            if (tail.length) {
              if (offlineEpochRef.current === c.genesisId) {
                runTailReplay(c, "plain");
              } else {
                const structural = toSuggestions(tail, "", "").dropped.length;
                const mode = arrivalMode(tail.length, true) === "draft" ? "draft" : "suggest";
                setArrival({ mode, tailLength: tail.length, structural });
              }
            }
          }
        }
      },
      onPresence: (participant, pos) => setPresence((prev) => ({ ...prev, [participant]: pos })),
      onRefused: (reason) => {
        if (reason === "read-only") {
          setReadOnlyBlocked(true);
          return;
        }
        if (reason === "not-owner") {
          setNotOwner(true);
          return;
        }
        clearDeadlines();
        setRefused(reason);
      },
      onSessionWarning: ({ reason, inMs }) => setDeadlines((d) => ({ ...d, [reason]: Date.now() + inMs })),
      // Only ever `idle` — the lifetime deadline cannot be cancelled, so
      // there is deliberately no per-reason branch here to get wrong.
      onSessionWarningCleared: () => setDeadlines((d) => d.idle === null ? d : { ...d, idle: null }),
      onRoster: (r) => setRoster(r),
      onSelfHeal: () => setSelfHeals((n) => n + 1),
      // A submit that never left the client. Counted, never swallowed.
      onSubmitDropped: () => setDroppedPreReady((n) => n + 1),
      // Async failures with nowhere to return (seal, transport send). Surfaced
      // rather than swallowed — this callback exists because three bugs this
      // session lived inside empty catches. console.error is the floor: an app
      // embedding this should report it properly.
      onError: (info) => {
        if (info.where === "enc.submit") setSendFailures((n) => n + 1);
        console.error(`[wordinweb] ${info.where}`, info.error);
      },
      onEpochChange: (from, to) => {
        if (store) {
          void store.get(docId).then((old) => {
            if (old && old.genesisId === from) {
              return store.put({ ...old, docId: draftKey(docId, from) });
            }
          });
        }
        setEpochChanged({ from, to });
      },
      onFastForward: (from, _to) => {
        if (store) {
          void store.get(docId).then((old) => {
            if (old && old.genesisId === from) {
              return store.put({ ...old, docId: supersededKey(docId, from) });
            }
          });
        }
      }
    };
    const flush = () => {
      void persister?.flush();
      flushOfflinePersist();
    };
    const rejoining = reconnectNonce > 0;
    void (async () => {
      const stretched = docKey && shareCode ? await stretchShareCode(shareCode, docId) : void 0;
      const codeProof = stretched ? btoa(String.fromCharCode(...stretched)) : void 0;
      if (disposed) return;
      const mediaOpts = httpBase ? { httpBase } : void 0;
      const conn = docKey ? new EncryptedCollabConnection(transport, clientId, docKey, callbacks, stretched, void 0, mediaOpts) : new CollabConnection(transport, clientId, callbacks, mediaOpts);
      connRef.current = conn;
      if (store) {
        persister = new BundlePersister(conn, store, docId, {
          // The browser's bundle IS the durable copy in a zero-custody design,
          // so a failed write (quota, blocked storage, private mode) means
          // "saved" is a lie. Never silent.
          onError: (err) => {
            setPersistErrors((n) => n + 1);
            console.error("[wordinweb] bundle-persist", err);
          },
          // The offline tail rides every bundle write (doc 15 §4.3 "tail
          // hygiene"): without this, the first post-resume write would
          // ERASE a stored tail the user had not reconciled yet.
          offlineTail: () => {
            const tail = offlineTailRef.current ?? [];
            return tail.length ? { tail: [...tail], epoch: offlineEpochRef.current ?? void 0 } : null;
          }
        });
        persisterRef.current = persister;
        if (typeof window !== "undefined") {
          window.addEventListener("pagehide", flush);
          document.addEventListener("visibilitychange", flush);
        }
        await withDeadline(flushDoneRef.current, STORE_DEADLINE_MS, void 0);
        if (disposed || connRef.current !== conn) return;
        const read = await withDeadline(
          store.get(docId),
          STORE_DEADLINE_MS,
          TIMED_OUT
        );
        if (disposed || connRef.current !== conn) return;
        const storeTimedOut = read === TIMED_OUT;
        const bundle = storeTimedOut ? null : read;
        if (storeTimedOut) setStoreSlow(true);
        if (!offlineTailRef.current?.length) {
          offlineTailRef.current = bundle?.offlineTail;
          offlineEpochRef.current = bundle?.offlineTailEpoch ?? null;
          setOfflineHeld(bundle?.offlineTail?.length ?? 0);
        }
        if (bundle) conn.resume(bundle, token, { profile, codeProof, ownerToken });
        else conn.join(docId, token, { profile, takeover: takeover || rejoining, codeProof, ownerToken });
      } else {
        conn.join(docId, token, { profile, takeover: takeover || rejoining, codeProof });
      }
    })();
    return () => {
      disposed = true;
      monitor?.stop();
      if (typeof document !== "undefined") document.removeEventListener("visibilitychange", onVisible);
      if (store && typeof window !== "undefined") {
        window.removeEventListener("pagehide", flush);
        document.removeEventListener("visibilitychange", flush);
      }
      flushOfflinePersist();
      flushDoneRef.current = Promise.all([flushDoneRef.current, persister?.flush() ?? Promise.resolve()]).then(() => {
      });
      persister?.stop();
      if (persisterRef.current === persister) persisterRef.current = null;
      connRef.current = null;
    };
  }, [url, docId, clientId, token, createSocket, store, takeover, docKey, shareCode, ownerToken, httpBase, reconnectNonce]);
  useEffect(() => () => {
    if (retryTimerRef.current !== null) clearTimeout(retryTimerRef.current);
    retryTimerRef.current = null;
  }, []);
  const myWrite = roster.find((r) => r.clientId === clientId)?.write;
  const refusalBlocked = refused !== null || (myWrite !== void 0 ? myWrite !== "allowed" : readOnlyBlocked);
  const offlineActive = !refusalBlocked && connection !== "live" && ready;
  const offlineCapped = offlineActive && offlineHeld >= (opts.offlineTailCap ?? OFFLINE_TAIL_CAP);
  return {
    doc,
    version,
    renderVersion,
    // Union of what the connection applied (its replica tracks per-intent
    // scopes) and any local offline applies, drained together at the repaint.
    takeRenderScope: () => {
      const local = renderScopeRef.current;
      renderScopeRef.current = null;
      const conn = connRef.current?.takeRenderScope() ?? null;
      if (local && conn) return unionScopes(local, conn);
      return conn ?? local;
    },
    docEpoch,
    ready,
    // DocxView's editor applies every command to the live doc BEFORE emitting
    // its intent, so the connection must not optimistically re-apply it (that
    // doubled each keystroke: "Hello" rendered as "Hello" + its reversal, and
    // the corrupted offsets got everything after the first char rejected
    // server-side). Pre-applied semantics track pending + send only.
    //
    // OFFLINE, the same intent is CAPTURED instead (doc 15 §2): the edit is
    // already in the doc, so it only needs recording. Past the cap nothing
    // routes here — writesBlocked has withdrawn the editor — and anything
    // that slips to the connection anyway lands in the loud droppedPreReady
    // backstop rather than vanishing.
    submit: (intent) => {
      if (offlineActive && !offlineCapped) return captureOffline(
        intent,
        /*preApplied*/
        true
      );
      connRef.current?.submitPreApplied(intent);
    },
    // Toolbar/API ops: canonical-apply optimistic path (see CollabSession).
    // Offline, the capture applies through the SAME canonical applyIntent
    // code before recording, so the local result stays byte-identical to
    // what every replica will derive when the tail replays.
    submitOp: (intent) => {
      if (offlineActive && !offlineCapped) return captureOffline(
        intent,
        /*preApplied*/
        false
      );
      connRef.current?.submit(intent);
    },
    setPresence: (pos) => connRef.current?.setPresence(pos),
    allocIds: (n) => connRef.current?.allocIds(n) ?? [],
    uploadMedia: async (bytes) => await connRef.current?.uploadMedia(bytes) ?? null,
    // Read straight off the connection rather than mirrored into state: it is
    // set once when the welcome lands and never changes for the session, and
    // in the encrypted connection it is assigned OUTSIDE the serial rehydrate
    // chain, so it is readable before `ready` — which is what lets the picker
    // open and the pre-check run while the document is still rehydrating.
    mediaMaxBlobBytes: connRef.current?.mediaMaxBlobBytes ?? null,
    // Collaborative undo lives on the ENCRYPTED connection: it needs the
    // mirror (a local re-derivation of the canonical log) to compute the
    // inverse. A plaintext room's authority is the server, so undo there
    // needs a wire message and is not wired yet — the optional call reports
    // "unavailable" rather than pretending.
    // The ref is TYPED as the plaintext class (the encrypted one is cast in
    // at construction — a pre-existing modelling wrinkle, not this arc's), so
    // the access is structural. The TYPE is the real one now, so a change to
    // UndoOutcome breaks here rather than silently disagreeing.
    undoLast: () => connRef.current?.undoLast?.() ?? "unavailable",
    presence,
    refused,
    sessionWarning,
    readOnlyBlocked,
    notOwner,
    // Derived, not stored: one place decides "writes won't land", so the gate
    // and the banner can never disagree about whether the user may type.
    //
    // THE SERVER'S ROSTER SIGNAL WINS WHEN PRESENT, and that is the fix for a
    // real user-visible bug: `readOnlyBlocked` is refusal-driven, so it can
    // only ever say "blocked" — nothing tells it a lock was LIFTED. A blocked
    // participant therefore sat in viewer mode until reload, and needed the
    // "Try editing again" button to escape. The roster status is POSITIVE: it
    // says `allowed` too, and it is re-fanned on every transition, so a lift
    // reaches this client without anyone attempting an edit to discover it.
    //
    // `undefined` means an older server that publishes no status. Fall back to
    // the refusal-driven flag rather than assuming `allowed` — a permissive
    // default would put the user's first keystrokes straight back in the
    // appear-then-vanish loop this exists to remove.
    //
    // A NON-LIVE CONNECTION DOES NOT BLOCK: unreachable is not refused (see
    // refusalBlocked above, and the field's own doc). What DOES fold in is
    // the offline tail CAP — past it an edit cannot be durably kept, which
    // is the same "cannot land" fact this predicate exists to gate, and
    // folding it here is what keeps `editable: … && !writesBlocked` the one
    // gate every editor surface already reads.
    //
    // A REFUSAL BLOCKS WRITES FIRST OF ALL: the session is dead, and note the
    // hub does NOT close the socket after a kick, so `connection` can still
    // read "live" and the roster can still hold a stale `allowed` — neither
    // may win over a refusal. This is what keeps a document rendered behind a
    // "session ended" dialog read-only through the same single predicate.
    writesBlocked: refusalBlocked || offlineCapped,
    // Offline editing state (doc 15 §2) — see the CollabSession doc. Null
    // whenever a refusal already owns the story: the two must never both
    // claim the banner.
    offline: offlineActive ? { editsHeld: offlineHeld, capped: offlineCapped } : null,
    // Left UNSET by a disconnect on purpose. These four values describe what
    // the SERVER decided about this participant, and a client that cannot
    // reach the server has not been told anything new — reporting a stale
    // `allowed` beside `writesBlocked: true` would be a straight
    // contradiction, and inventing a fifth value would put a transport fault
    // into a permissions vocabulary. `connection` is where that story lives.
    writeStatus: myWrite,
    connection,
    reconnect,
    retryWrites: () => setReadOnlyBlocked(false),
    epochChanged,
    roster,
    setProfile: (p) => connRef.current?.setProfile(p),
    activity: connRef.current?.activity ?? [],
    admin: (action) => connRef.current?.admin(action),
    selfHeals,
    droppedPreReady,
    persistErrors,
    storeSlow,
    sendFailures,
    arrival,
    reconcile: (mode) => {
      const conn = connRef.current;
      const tail = offlineTailRef.current ?? [];
      if (!conn || !tail.length) {
        offlineTailRef.current = void 0;
        offlineEpochRef.current = null;
        setOfflineHeld(0);
        setArrival(null);
        return;
      }
      setArrival(null);
      if (mode === "suggest") {
        runTailReplay(conn, "suggest");
      } else {
        offlineTailRef.current = void 0;
        offlineEpochRef.current = null;
        setOfflineHeld(0);
        void persisterRef.current?.flush();
      }
    }
  };
}
var COLLAB_TOOLBAR_DEFAULTS = {
  // Images ride the doc-16 media relay (bytes out of band, address committed
  // in the intent). The rest of the upload surface stays closed per the demo
  // threat model — no object upload, no stored-payload surface.
  icon: false,
  screenshot: false,
  model3D: false,
  media: false,
  object: false,
  history: false,
  // toolbar undo/redo drives the LOCAL history; collaborative undo is server-side
  drawing: false,
  // ink strokes have no intent yet
  arrange: false
  // selected-object arrange ops are not collab-anchored yet
};
function CollabEditor(opts) {
  const session = useCollab(opts);
  useEffect(() => {
    opts.onSession?.(session);
  }, [session.version, session.ready, session.refused, session.roster, session.epochChanged, session.sessionWarning, session.writesBlocked, session.connection]);
  const [api, setApi] = useState(null);
  const pendingCaretRef = useRef(null);
  const seenEpochRef = useRef(session.docEpoch);
  if (session.docEpoch !== seenEpochRef.current) {
    seenEpochRef.current = session.docEpoch;
    pendingCaretRef.current = api?.getEncodedCaret() ?? null;
  }
  const bytes = useMemo(
    () => session.refused && session.doc ? session.doc.save() : null,
    [session.refused, session.doc, session.docEpoch]
  );
  const liveSource = useMemo(() => new Uint8Array(0), []);
  if (session.refused) {
    const docVisible = !!(session.doc && bytes);
    const content = opts.refusedContent ? opts.refusedContent(session.refused, { docVisible }) : createElement(
      "div",
      // Legible over a rendered page in the overlay case; inert extra
      // styling on the plain full-screen path costs nothing.
      docVisible ? { style: { margin: 16, padding: "10px 14px", width: "fit-content", background: "rgba(255,255,255,.95)", borderRadius: 8, font: "14px system-ui", boxShadow: "0 2px 8px rgba(0,0,0,.25)" } } : void 0,
      `Please refresh \u2014 ${session.refused}.`
    );
    if (!docVisible) {
      return createElement("div", { className: "dxw-collab-refused" }, content);
    }
    return createElement(
      "div",
      { className: "dxw-collab-refused-host", style: { position: "relative", height: "100%" } },
      createElement(DocxView, {
        source: bytes,
        // NO `collab` prop: a refused session is dead — no socket, nothing to
        // submit to — so this is a plain document view, not a live editor.
        // And the SAME single gate as the live path below: refusal folds into
        // writesBlocked (see CollabSession), so this is always false here —
        // an editable surface over a dead session would apply keystrokes
        // locally and lose them.
        editable: (opts.editable ?? true) && !session.writesBlocked,
        style: { height: "100%" }
      }),
      // The overlay covers the document (blocking pointer interaction with a
      // dead page); the host's content styles itself on top of it. The
      // load-bearing `dxw-collab-refused` class keeps its meaning — "the
      // element holding the refusal content" — with the overlay placement
      // carried by the additional class. DELIBERATELY NO z-index: it would
      // create a stacking context that traps a host's `position: fixed`
      // dialog/scrim at this layer, silently re-scoping the host's own
      // z-values. DOM order (after the DocxView) already paints it above the
      // document.
      createElement(
        "div",
        { className: "dxw-collab-refused dxw-collab-refused-overlay", style: { position: "absolute", inset: 0 } },
        content
      )
    );
  }
  if (!session.ready || !session.doc) {
    return createElement(ConnectingNotice, null);
  }
  const view = createElement(DocxView, {
    source: liveSource,
    onReady: (a) => {
      setApi(a);
      opts.onReady?.(a);
      const caret = pendingCaretRef.current;
      if (caret) {
        pendingCaretRef.current = null;
        a.setCaretFromEncoded(caret);
      }
    },
    // Render the live doc object directly; repaint in place on each version
    // bump. submit + presence + id allocator flow out; DocxView draws carets.
    collab: {
      submit: session.submit,
      submitOp: session.submitOp,
      presence: session.presence,
      allocIds: session.allocIds,
      // Image bytes go out of band over the relay (doc 16); its presence is
      // what makes the toolbar's image button real in a room.
      uploadMedia: session.uploadMedia,
      mediaMaxBlobBytes: session.mediaMaxBlobBytes,
      // Cmd+Z: reverse my last sequenced action over the wire.
      undoLast: () => {
        session.undoLast();
      },
      doc: session.doc,
      // The repaint signal is renderVersion, NOT version: version moves on
      // every onChange (including the editor's own submit + echo, which the
      // editor already painted), and DocxView answers each renderSignal move
      // with a whole-document relayout — behind an inert container past the
      // background-layout page threshold. Riding version made every keystroke
      // in a 500-page room queue that relayout: the "collab editor is
      // unusable on a big document" report.
      renderSignal: session.renderVersion,
      // The dirty scope behind renderSignal: a remote text edit repaints as
      // ONE incremental paragraph relayout instead of the whole document.
      takeRenderScope: session.takeRenderScope,
      // Outbound presence: the editor reports caret moves; remote tabs draw
      // this user's cursor (inbound presence above draws theirs here).
      setPresence: session.setPresence,
      // Name flags on remote carets (doc 14 §2): presence and roster share
      // the bound-clientId keyspace, so this join is exact.
      participantNames: Object.fromEntries(session.roster.map((r) => [r.clientId, r.profile.name]))
    },
    // VIEWER MODE when the server will not take this client's writes. Not a
    // banner over a live editor: that shape applied every keystroke locally
    // and lost it, which is why the read-only banner could be on screen while
    // the user's text appeared and then vanished.
    editable: (opts.editable ?? true) && !session.writesBlocked,
    // Re-key only on docEpoch (a true-conflict reload) — NOT on every change.
    // Between reloads the live doc mutates in place and the key stays stable,
    // so DocxView repaints in place instead of re-mounting (no flash/jump).
    key: session.docEpoch
  });
  if (opts.toolbar === false) return view;
  return createElement(
    "div",
    { className: "dxw-collab-shell", style: { display: "flex", flexDirection: "column", height: "100%" } },
    createElement(DocxToolbar, {
      api,
      mode: opts.toolbarMode ?? "advanced",
      features: { ...COLLAB_TOOLBAR_DEFAULTS, ...opts.toolbarFeatures }
    }),
    createElement("div", { style: { flex: 1, minHeight: 0 } }, view)
  );
}
function ConnectingNotice() {
  const [slow, setSlow] = useState(false);
  useEffect(() => {
    const t = setTimeout(() => setSlow(true), 5e3);
    return () => clearTimeout(t);
  }, []);
  return createElement(
    "div",
    {
      className: "dxw-collab-connecting",
      style: { padding: 16, font: "14px system-ui" },
      role: "status",
      "aria-live": "polite",
      "aria-busy": "true"
    },
    createElement("span", { className: "dxw-collab-spinner", "aria-hidden": "true", key: "s" }),
    slow ? "Still connecting \u2014 is the collab server running? (start it, then reload this page)" : "Connecting\u2026"
  );
}
export {
  BundlePersister,
  COLLAB_TOOLBAR_DEFAULTS,
  CollabEditor,
  InMemoryBundleStore,
  IndexedDbBundleStore,
  bytesToB64,
  deriveEpochKeys,
  docHash,
  docKeyFromFragment,
  draftKey,
  mediaAddressesOf,
  mintDocKey,
  parseBundleKey,
  sealCheckpoint,
  stretchShareCode,
  supersededKey,
  useCollab,
  versionKey
};
//# sourceMappingURL=collab.js.map