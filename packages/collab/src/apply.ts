import {
  DocxDocument,
  StableIds,
  deleteSuggestedRange,
  markParagraphGlyph,
  suggestMeta,
  applyInsertText,
  applySplitParagraph,
  applyDeleteRange,
  applyInsertSeparator,
  applyDeleteSeparator,
  separatorAtWireOffset,
  applyRunFormat,
  setParagraphAlignment,
  setParagraphStyle,
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
  setDrawingWordArtStyle,
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
  moveMath,
  deleteComment,
  setCommentResolved,
  editCommentText,
  insertBookmarkAroundSelection,
  collectRevisions,
  acceptRevision,
  rejectRevision,
  acceptAllRevisions,
  rejectAllRevisions,
  isSafeUrl,
  applyTableOp,
  runWireLength,
  resolveRunOffset,
  resizeDrawing,
  resizeTableColumn,
  moveTableTo,
  removeDrawingRun,
  mergeParagraphBackward,
  siblingParagraph,
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
  ADDRESS_WIRE_FIELD,
  applyRegisteredOperation,
  registeredOperation,
  type OperationAddress,
  type OperationTarget,
  type EditCaret,
  type Run,
  type Block,
  type XmlElement,
  type SelectionSegment,
} from "@wordinweb/core";
import { Intent, Position, isRegisteredIntent, type RegisteredIntent } from "./intents.js";

/**
 * How much of the document an applied intent disturbed — the input to
 * dirty-scoped reconciliation (perf B9/B10). Reconciling an edit used to cost
 * O(document) on every path (full model reparse + full id walk), so per-op
 * cost grew with the document; the scope lets the caller pay for THE EDIT.
 *
 * `doc` is the conservative default and means exactly today's behavior: full
 * refresh() + assignFromRoots(). Only intents whose blast radius is provably
 * one paragraph (or the two sides of a split) report a narrower scope; every
 * structural, document-level, or exotic intent keeps the document scope.
 */
export type Scope =
  | { kind: "doc" }
  /** Only these paragraphs' contents changed. */
  | { kind: "block"; blocks: XmlElement[] }
  /** `before` was split; `after` is the new sibling that follows it. */
  | { kind: "split"; before: XmlElement; after: XmlElement };

export type ApplyScope = Scope & { applied: boolean };

const DOC_SCOPE: Scope = { kind: "doc" };

/**
 * Union of two dirty scopes — the batching rule for a coalesced repaint: a
 * receive batch (or one rAF window) can apply several intents, and the single
 * repaint that follows must relayout everything any of them touched, once.
 * Doc absorbs everything. Two block scopes merge their block lists. A split
 * stays a split only while every other dirty block is one of its two halves
 * (the layout engine's insertion fast path re-lays exactly those); any
 * combination the narrow forms cannot express verifiably widens to doc scope,
 * which is always correct.
 */
export function unionScopes(a: Scope | null, b: Scope): Scope {
  if (!a) return b;
  if (a.kind === "doc" || b.kind === "doc") return DOC_SCOPE;
  if (a.kind === "split" && b.kind === "split") {
    return a.before === b.before && a.after === b.after ? a : DOC_SCOPE;
  }
  if (a.kind === "split" || b.kind === "split") {
    const split = a.kind === "split" ? a : (b as Extract<Scope, { kind: "split" }>);
    const block = a.kind === "block" ? a : (b as Extract<Scope, { kind: "block" }>);
    return block.blocks.every((el) => el === split.before || el === split.after) ? split : DOC_SCOPE;
  }
  const blocks = [...a.blocks];
  for (const el of b.blocks) if (!blocks.includes(el)) blocks.push(el);
  return { kind: "block", blocks };
}

/**
 * Bridge from wire intents to the core mutation functions. Resolves stable
 * ids to model positions, applies the mutation headlessly (no DOM), and
 * assigns carried ids to newly created nodes so every replica agrees on ids
 * (plan doc 03). Callers reconcile the parsed model afterward — see
 * resyncScope; this function performs the tree mutation only.
 *
 * Returns true if the intent applied, false for a clean no-op / unresolvable
 * position (the caller records a rejection, doc 03).
 */
export function applyIntent(doc: DocxDocument, ids: StableIds, intent: Intent): boolean {
  return applyIntentScoped(doc, ids, intent).applied;
}

/** applyIntent plus the dirty scope of what it touched. */
export function applyIntentScoped(doc: DocxDocument, ids: StableIds, intent: Intent): ApplyScope {
  const out = { scope: DOC_SCOPE };
  try {
    const applied = applyIntentInner(doc, ids, intent, out);
    return { ...out.scope, applied };
  } catch {
    // Any throw (e.g. a carried-id collision from a buggy/hostile client)
    // becomes a clean rejection: the session logs a reject entry and every
    // replica's receive loop keeps running instead of dying mid-broadcast.
    // A throw can leave a half-applied tree, so the scope is not trustworthy —
    // report document scope, which is what a recovering caller must assume.
    return { kind: "doc", applied: false };
  }
}

/**
 * Reconcile the parsed model and the id table with a tree the given intent
 * just mutated, doing work proportional to the intent's scope. Falls back to
 * the full refresh whenever a targeted reparse cannot handle the container
 * (paragraph outside the body story, a paragraph carrying section breaks or
 * bookmark ranges, a split whose halves aren't plain adjacent siblings) —
 * correctness beats speed, and the fallback is exactly the old behavior.
 */
export function resyncScope(doc: DocxDocument, ids: StableIds, scope: Scope): void {
  if (scope.kind === "block") {
    let ok = true;
    for (const block of scope.blocks) ok = doc.reparseBodyParagraph(block) !== null && ok;
    if (ok) {
      // Scoped assignment is id-table-identical to the full walk here: a
      // block-scoped intent creates tracked nodes only inside these
      // paragraphs, and everything outside already carries an id. See
      // StableIds.assignFromSubtrees.
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

function applyIntentInner(
  doc: DocxDocument,
  ids: StableIds,
  intent: Intent,
  out: { scope: Scope },
): boolean {
  const runOf = makeRunLookup(doc, ids);
  // Suggesting mode rides on the intent itself: kinds with a tracked form
  // carry `suggest` (author/date) and build their own suggesting ctx below.
  // This default ctx serves the rest, which never suggest.
  const ctx = { suggesting: false, revMeta: () => { throw new Error("suggesting mode unsupported headlessly"); } };

  // Registered operations carry their mutation in the core registry. Narrowing
  // here leaves the switch below exhaustive over what is still hand-written,
  // so its `never` gate keeps its full force.
  if (isRegisteredIntent(intent)) return applyRegistered(doc, ids, runOf, intent);

  switch (intent.kind) {
    case "insertText": {
      const caret = resolveCaret(ids, runOf, intent.at);
      if (!caret) return false;
      if (intent.suggest) {
        // Tracked change: record the insertion as a w:ins with carried
        // author/date; the revision w:id is scan-based (deterministic).
        const s = intent.suggest;
        const suggestCtx = { suggesting: true, revMeta: () => ({ author: s.author, date: s.date, nextId: () => doc.nextRevisionId() }) };
        // Verify the scope BEFORE mutating: the insertion may split the
        // addressed run into fresh pieces, after which the original run
        // element is no longer in the tree and containment can't be checked.
        const scope = blockScope(ids, intent.at.blockId, intent.at.runId);
        applyInsertText(doc, caret, intent.text, suggestCtx);
        // The insertion (w:ins wrapper, split run pieces) stays inside the
        // addressed paragraph, so reconcile at the PARAGRAPH's cost: scoped
        // reparse + assignFromSubtrees produces the id table the old full
        // refresh()+assignFromRoots() did (the documented equivalence in
        // StableIds.assignFromSubtrees — new nodes only inside the block).
        // This ran a full-document refresh per remote suggest KEYSTROKE, and
        // the doc-wide scope it reported forced a whole-document relayout per
        // keystroke on every peer — on a 500-page document, seconds of inert
        // editor per keystroke (the livelock's fuel). Unverifiable block ids
        // fall back to doc scope; resyncScope's own fallback is the old full
        // refresh, so anything the scoped reparse cannot handle is unchanged.
        if (doc.stableIds) resyncScope(doc, ids, scope);
        out.scope = scope;
        return true;
      }
      applyInsertText(doc, caret, intent.text, ctx);
      // Plain typing splices characters into one w:t — the whole blast radius
      // is the addressed paragraph (perf B9/B10: this is THE hot path).
      out.scope = blockScope(ids, intent.at.blockId, intent.at.runId);
      return true;
    }
    case "insertSeparator": {
      const caret = resolveCaret(ids, runOf, intent.at);
      if (!caret) return false;
      if (intent.suggest) {
        // Tracked separator: a w:ins sibling with carried author/date (the
        // same scoped-reparse discipline the suggest insertText apply uses —
        // run splits without carried ids resync by parse order).
        const s = intent.suggest;
        const suggestCtx = { suggesting: true, revMeta: () => ({ author: s.author, date: s.date, nextId: () => doc.nextRevisionId() }) };
        const scope = blockScope(ids, intent.at.blockId, intent.at.runId);
        if (!applyInsertSeparator(doc, caret, intent.separator, suggestCtx)) return false;
        if (doc.stableIds) resyncScope(doc, ids, scope);
        out.scope = scope;
        return true;
      }
      // Plain: an in-run w:t split around the separator — no run is created
      // or removed, so no id changes and the paragraph is the blast radius.
      if (!applyInsertSeparator(doc, caret, intent.separator, ctx)) return false;
      out.scope = blockScope(ids, intent.at.blockId, intent.at.runId);
      return true;
    }
    case "deleteSeparator": {
      const runEl = ids.elOf(intent.at.runId);
      if (!runEl) return false;
      // Resolve the addressed wire unit to the concrete separator element; a
      // unit that no longer holds one (concurrent edit) is a clean no-op.
      const sepEl = separatorAtWireOffset(runEl, intent.at.offset);
      if (!sepEl) return false;
      if (intent.suggest) {
        // Tracked separator deletion: the host run splits around a w:del
        // (same scoped-reparse discipline the suggest text strike uses —
        // run splits without carried ids resync by parse order).
        const s = intent.suggest;
        const suggestCtx = { suggesting: true, revMeta: () => ({ author: s.author, date: s.date, nextId: () => doc.nextRevisionId() }) };
        const scope = blockScope(ids, intent.at.blockId, intent.at.runId);
        if (!applyDeleteSeparator(doc, sepEl, suggestCtx)) return false;
        if (doc.stableIds) resyncScope(doc, ids, scope);
        out.scope = scope;
        return true;
      }
      // Plain: the separator element leaves the run and its two w:t halves
      // join — the insert's exact inverse; the paragraph is the whole blast
      // radius. Scope resolves BEFORE the mutation: an emptied
      // separator-only run is dropped by the mutation, and its retired id
      // is pruned.
      const scope = blockScope(ids, intent.at.blockId, intent.at.runId);
      const res = applyDeleteSeparator(doc, sepEl, ctx);
      if (!res) return false;
      if (res.removedRun) ids.prune(doc.editableRoots());
      out.scope = scope;
      return true;
    }
    case "deleteText": {
      if (intent.end <= intent.start) return false;
      const caret = resolveCaret(ids, runOf, { blockId: intent.blockId, runId: intent.runId, offset: intent.start });
      if (!caret) return false;
      // Wire offsets are cumulative within the run; the resolved caret is
      // local to its w:t. A range that would cross into the NEXT w:t is not
      // emitted by any sender (deletes are per-grapheme within one w:t) —
      // reject it as a clean no-op rather than corrupting a neighbor.
      const localEnd = caret.offset + (intent.end - intent.start);
      if (localEnd > caret.t.text.length) return false;
      // Engine-independent splice via the shared core mutation (offsets were
      // resolved client-side, so no Intl.Segmenter dependency here).
      applyDeleteRange(caret, caret.offset, localEnd);
      // A delete only shortens one w:t's text: no tracked node is removed, so
      // no id retires and the paragraph is the whole blast radius.
      out.scope = blockScope(ids, intent.blockId, intent.runId);
      return true;
    }
    case "splitParagraph": {
      const caret = resolveCaret(ids, runOf, intent.at);
      if (!caret) return false;
      const splitCtx = intent.suggest
        ? {
            suggesting: true,
            revMeta: () => ({
              author: intent.suggest!.author,
              date: intent.suggest!.date,
              nextId: () => doc.nextRevisionId(),
            }),
          }
        : ctx;
      const res = applySplitParagraph(doc, caret, splitCtx);
      if (!res) return false;
      ids.assign(res.after, intent.newBlockId);
      const newRun = res.after.children.find((c) => localRun(c));
      if (newRun) ids.assign(newRun, intent.newRunId);
      // Both halves of the split carry their ids explicitly (above), so the
      // scoped reparse below has only the two paragraphs' MODEL to rebuild.
      out.scope = { kind: "split", before: res.before, after: res.after };
      return true;
    }
    case "formatRun": {
      const runEl = ids.elOf(intent.runId);
      if (!runEl) return false;
      const entry = runOf(runEl);
      if (!entry) return false;
      // Whole-run format: a segment with t=null covers the entire run, so
      // applyRunFormat takes the in-place (no-split) path — run id preserved.
      const seg: SelectionSegment = { run: entry.run, t: null, start: 0, end: 0, props: entry.run.props };
      applyRunFormat(doc, [seg], intent.patch as never, suggestMeta(doc, intent.suggest));
      return true;
    }
    case "formatParagraph": {
      const blockEl = ids.elOf(intent.blockId);
      if (!blockEl) return false;
      // setParagraphAlignment/Style resolve the paragraph by walking UP from a
      // target, so pass a descendant w:t (or the block itself as a fallback).
      const target = firstTextDescendant(blockEl) ?? blockEl;
      const meta = suggestMeta(doc, intent.suggest);
      let changed = false;
      if (intent.align) changed = setParagraphAlignment(doc, [target], intent.align, meta) || changed;
      if (intent.styleId !== undefined) changed = setParagraphStyle(doc, [target], intent.styleId, meta) || changed;
      return changed;
    }
    case "formatRange": {
      if (intent.end <= intent.start) return false;
      const runEl = ids.elOf(intent.runId);
      if (!runEl) return false;
      const entry = runOf(runEl);
      if (!entry || !entry.firstT) return false;
      const len = runWireLength(runEl);
      if (intent.start < 0 || intent.end > len) return false;
      // Wire offsets are cumulative within the run; the segment applyRunFormat
      // takes is local to one w:t. A range crossing a w:t boundary has no
      // single-segment form — reject it as a clean no-op (senders emit
      // per-segment ranges, each within one w:t).
      const hit = resolveRunOffset(runEl, intent.start);
      if (!hit) return false;
      const localEnd = hit.offset + (intent.end - intent.start);
      if (localEnd > hit.t.text.length) return false;
      // Sub-range format: splits the run into before/middle/after (all new).
      const seg: SelectionSegment = { run: entry.run, t: hit.t, start: hit.offset, end: localEnd, props: entry.run.props };
      const formatted = applyRunFormat(doc, [seg], intent.patch as never, suggestMeta(doc, intent.suggest));
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
      // mergeDown mints a fresh continuation paragraph, splitCell fresh cells;
      // both also retire ids (the merged/replaced content), so they take
      // carried ids like the inserts AND prune like the deletes.
      const createsNodes = isInsert || intent.op === "mergeDown" || intent.op === "splitCell";
      // For node-creating ops, snapshot the tracked-node set so we can find
      // the new nodes afterward and give them the carried ids.
      const before = createsNodes ? trackedSet(ids, doc) : null;
      const ok = applyTableOp(doc, target, intent.op as never, suggestMeta(doc, intent.suggest));
      if (!ok) return false;
      if (before && intent.nodeIds) {
        // Assign carried ids to the newly created tracked nodes in doc order.
        const fresh: XmlElement[] = [];
        walkTracked(doc, (el) => { if (!before.has(el)) fresh.push(el); });
        for (let k = 0; k < fresh.length && k < intent.nodeIds.length; k++) {
          ids.reassign(fresh[k], intent.nodeIds[k]);
        }
      }
      if (!isInsert) {
        // Delete/merge/split/shading ops: retire stale ids for removed content.
        ids.prune(doc.editableRoots());
      }
      return true;
    }
    case "mergeParagraph": {
      const pEl = ids.elOf(intent.blockId);
      if (!pEl) return false;
      const meta = suggestMeta(doc, intent.suggest);
      if (meta) {
        // Tracked merge: strike the pilcrow between the two paragraphs — the
        // PREVIOUS paragraph's mark — and leave both paragraphs standing. The
        // editor's Backspace-at-paragraph-start does exactly this locally.
        const previous = siblingParagraph(doc, pEl, -1);
        if (!previous) return false;
        markParagraphGlyph(previous, "del", meta);
        doc.refresh();
        return true;
      }
      const ok = mergeParagraphBackward(doc, pEl);
      if (ok) ids.prune(doc.editableRoots()); // retire the merged paragraph's id
      return ok;
    }
    case "commentRun": {
      const runEl = ids.elOf(intent.runId);
      if (!runEl) return false;
      const entry = runOf(runEl);
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
    case "suggestRevision": {
      // Resolve ranges to text elements via stable ids (offsets were
      // client-resolved — rule b); drop unresolvable/empty ranges as clean
      // no-ops. Marks resolve block elements directly.
      const meta = {
        author: intent.suggest.author,
        date: intent.suggest.date,
        nextId: () => doc.nextRevisionId(),
      };
      const resolved: { t: XmlElement; start: number; end: number; block: XmlElement | null }[] = [];
      for (const r of intent.ranges ?? []) {
        const caret = resolveCaret(ids, runOf, { blockId: r.blockId, runId: r.runId, offset: r.start });
        if (!caret) continue;
        const localEnd = caret.offset + (r.end - r.start);
        if (localEnd > caret.t.text.length) continue;
        // Verify the range's own blockId names the paragraph really holding
        // the run — the same rule blockScope enforces everywhere else. A
        // verified block lets the reconcile below stay scoped; any failure
        // widens to the full-document path.
        const sc = blockScope(ids, r.blockId, r.runId);
        resolved.push({ t: caret.t, start: caret.offset, end: localEnd, block: sc.kind === "block" ? sc.blocks[0] : null });
      }
      const markEls: { el: XmlElement; glyph: "ins" | "del"; block: XmlElement | null }[] = [];
      for (const m of intent.marks ?? []) {
        const el = ids.elOf(m.blockId);
        if (el) markEls.push({ el, glyph: m.glyph, block: localName(el.name) === "p" ? el : null });
      }
      if (resolved.length === 0 && markEls.length === 0) return false;
      if (resolved.length) deleteSuggestedRange(doc, resolved, meta);
      for (const m of markEls) markParagraphGlyph(m.el, m.glyph, meta);
      // Strikes wrap runs in w:del and split partially-struck runs — new
      // nodes, but all INSIDE the addressed paragraphs. When every touched
      // paragraph verified, reconcile at those paragraphs' cost (scoped
      // reparse + assignFromSubtrees — id-table-identical to the full walk,
      // see StableIds.assignFromSubtrees) and report them as the dirty
      // scope. The old unconditional refresh()+assignFromRoots() here ran a
      // full-document rebuild per remote suggest-DELETE keystroke and forced
      // a whole-document relayout on every peer; anything unverifiable keeps
      // exactly that behavior via the doc-scope fallback.
      const parts = [...resolved.map((x) => x.block), ...markEls.map((x) => x.block)];
      if (parts.every((b): b is XmlElement => b !== null)) {
        const blocks: XmlElement[] = [];
        for (const b of parts) if (!blocks.includes(b)) blocks.push(b);
        const scope: Scope = { kind: "block", blocks };
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
      // Doc 16 §2: apply performs the REGISTRATION half only — part + rel +
      // content-type + the drawing run with client-measured extents (layout
      // reserves exact space from XML alone, doc 05). Bytes arrive out of
      // band and install via installMedia when fetched; until then the part
      // is pending and renders as a placeholder. The sha/iv metadata is
      // recorded on the doc for verification + E2EE re-supply (doc 16 §5.3).
      const relId = doc.registerPendingImage(intent.blobSha, intent.ext, { iv: intent.iv }, runEl);
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
      const drawing = insertChartAt(doc, entry.firstT, intent.chart as never);
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
      const drawing = drawingIn(runEl, intent.objectIndex, runOf);
      if (!drawing) return false;
      return setDrawingRotation(doc, drawing, intent.degrees);
    }
    case "setDrawingFill": {
      const runEl = ids.elOf(intent.runId);
      if (!runEl) return false;
      const drawing = drawingIn(runEl, intent.objectIndex, runOf);
      if (!drawing) return false;
      return setDrawingFill(doc, drawing, intent.color);
    }
    case "setChartData": {
      const runEl = ids.elOf(intent.runId);
      if (!runEl) return false;
      const drawing = drawingIn(runEl, intent.objectIndex, runOf);
      if (!drawing) return false;
      return setChartData(doc, drawing, intent.chart as never);
    }
    case "setSmartArtNodeText": {
      const runEl = ids.elOf(intent.runId);
      if (!runEl) return false;
      const drawing = drawingIn(runEl, intent.objectIndex, runOf);
      if (!drawing) return false;
      return setSmartArtNodeText(doc, drawing, intent.index, intent.text);
    }
    case "setDrawingWordArtText": {
      const runEl = ids.elOf(intent.runId);
      if (!runEl) return false;
      const drawing = drawingIn(runEl, intent.objectIndex, runOf);
      if (!drawing) return false;
      return setDrawingWordArtText(doc, drawing, intent.text);
    }
    case "setDrawingWordArtStyle": {
      const runEl = ids.elOf(intent.runId);
      if (!runEl) return false;
      const drawing = drawingIn(runEl, intent.objectIndex, runOf);
      if (!drawing) return false;
      return setDrawingWordArtStyle(doc, drawing, intent.color, intent.opacity);
    }
    case "setDrawingLineStyle": {
      const runEl = ids.elOf(intent.runId);
      if (!runEl) return false;
      const drawing = drawingIn(runEl, intent.objectIndex, runOf);
      if (!drawing) return false;
      return setDrawingLineStyle(doc, drawing, intent.color, intent.widthPx, intent.dash);
    }
    case "setImageAltText": {
      const runEl = ids.elOf(intent.runId);
      if (!runEl) return false;
      const drawing = drawingIn(runEl, intent.objectIndex, runOf);
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
      const drawing = drawingIn(runEl, intent.objectIndex, runOf);
      if (!drawing) return false;
      return setImageWrap(doc, drawing, intent.mode);
    }
    case "setDrawingOrder": {
      const runEl = ids.elOf(intent.runId);
      if (!runEl) return false;
      const drawing = drawingIn(runEl, intent.objectIndex, runOf);
      if (!drawing) return false;
      return setDrawingOrder(doc, drawing, intent.order);
    }
    case "setSmartArtData": {
      const runEl = ids.elOf(intent.runId);
      if (!runEl) return false;
      const drawing = drawingIn(runEl, intent.objectIndex, runOf);
      if (!drawing) return false;
      return setSmartArtData(doc, drawing, intent.smartArt as never);
    }
    case "setSmartArtFill": {
      const runEl = ids.elOf(intent.runId);
      if (!runEl) return false;
      const drawing = drawingIn(runEl, intent.objectIndex, runOf);
      if (!drawing) return false;
      return setSmartArtFill(doc, drawing, intent.color, intent.nodeIndex);
    }
    case "setSmartArtTextFormat": {
      const runEl = ids.elOf(intent.runId);
      if (!runEl) return false;
      const drawing = drawingIn(runEl, intent.objectIndex, runOf);
      if (!drawing) return false;
      return setSmartArtTextFormat(doc, drawing, intent.format as never, intent.nodeIndex);
    }
    case "setFloatingPagePosition": {
      const runEl = ids.elOf(intent.runId);
      if (!runEl) return false;
      const drawing = drawingIn(runEl, intent.objectIndex, runOf);
      if (!drawing) return false;
      return setFloatingPagePosition(doc, drawing, intent.xPx, intent.yPx);
    }
    case "resizeDrawing": {
      const runEl = ids.elOf(intent.runId);
      if (!runEl) return false;
      const drawing = drawingIn(runEl, intent.objectIndex, runOf);
      if (!drawing) return false;
      return resizeDrawing(doc, drawing, intent.widthPx, intent.heightPx);
    }
    case "resizeTableColumn": {
      const tbl = tableOfParagraph(doc, ids, intent.cellParagraphId);
      if (!tbl) return false;
      return resizeTableColumn(doc, tbl, intent.boundary, intent.deltaPx, intent.renderedWidths);
    }
    case "moveTable": {
      const tbl = tableOfParagraph(doc, ids, intent.cellParagraphId);
      if (!tbl) return false;
      return moveTableTo(doc, tbl, intent.xPx, intent.yPx, intent.preservePageStart, intent.pageDelta);
    }
    case "removeDrawing": {
      const runEl = ids.elOf(intent.runId);
      if (!runEl) return false;
      const drawing = drawingIn(runEl, intent.objectIndex, runOf);
      if (!drawing) return false;
      const ok = removeDrawingRun(doc, drawing);
      if (ok) ids.prune(doc.editableRoots()); // retire the removed run's ids
      return ok;
    }
    case "setMathLinear": {
      const blockEl = ids.elOf(intent.blockId);
      if (!blockEl) return false;
      const math = mathIn(blockEl, intent.mathIndex);
      if (!math) return false;
      return setMathLinear(doc, math, intent.mathText);
    }
    case "deleteMath": {
      const blockEl = ids.elOf(intent.blockId);
      if (!blockEl) return false;
      const math = mathIn(blockEl, intent.mathIndex);
      if (!math) return false;
      return deleteMath(doc, math);
    }
    case "moveMath": {
      const blockEl = ids.elOf(intent.blockId);
      if (!blockEl) return false;
      const math = mathIn(blockEl, intent.mathIndex);
      if (!math) return false;
      const dest = resolveCaret(ids, runOf, intent.at);
      if (!dest) return false;
      const before = trackedSet(ids, doc);
      if (!moveMath(doc, math, dest.t, dest.offset)) return false;
      // A mid-text drop splits the destination run; the tail takes the
      // carried id so every replica addresses it identically. (reassign, not
      // assign: moveMath's internal refresh() may already have auto-numbered
      // it in parse order.)
      const fresh: XmlElement[] = [];
      walkTracked(doc, (el) => { if (!before.has(el)) fresh.push(el); });
      for (let k = 0; k < fresh.length && k < intent.nodeIds.length; k++) ids.reassign(fresh[k], intent.nodeIds[k]);
      return true;
    }
    case "ensureHeaderFooter": {
      // Already there: a clean no-op, so two participants opening the header
      // at the same moment can't create two parts.
      if (doc.hasHfPart(intent.hfKind)) return false;
      const before = trackedSet(ids, doc);
      doc.ensureHfPart(intent.hfKind);
      const fresh: XmlElement[] = [];
      walkTracked(doc, (el) => { if (!before.has(el)) fresh.push(el); });
      for (let k = 0; k < fresh.length && k < intent.nodeIds.length; k++) ids.reassign(fresh[k], intent.nodeIds[k]);
      return true;
    }
    case "deleteComment":
      return deleteComment(doc, intent.commentId);
    case "resolveComment":
      // The carried paraId candidate is consumed only when the thread parent's
      // body paragraph lacks one (recordedProvenance draws at most one value).
      return setCommentResolved(doc, intent.commentId, intent.resolved, recordedProvenance({ paraIds: [intent.paraId] }));
    case "editComment":
      return editCommentText(doc, intent.commentId, intent.text);
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
      const seg: SelectionSegment = { run: entry.run, t: hit.t, start: hit.offset, end: localEnd, props: entry.run.props };
      return insertBookmarkAroundSelection(doc, [seg], intent.name);
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
    case "rejectAllRevisions":
      return rejectAllRevisions(doc) > 0;
    case "setLink": {
      if (!isSafeUrl(intent.url)) return false; // reject javascript:/data: etc.
      const runEl = ids.elOf(intent.runId);
      if (!runEl) return false;
      const entry = runOf(runEl);
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
      const entry = runOf(runEl);
      if (!entry || !entry.firstT) return false;
      return insertBookmarkAt(doc, entry.firstT, entry.firstT.text.length, intent.name);
    }
    case "adjustIndent": {
      const blockEl = ids.elOf(intent.blockId);
      if (!blockEl) return false;
      const target = firstTextDescendant(blockEl) ?? blockEl;
      return adjustIndent(doc, [target], intent.direction, suggestMeta(doc, intent.suggest));
    }
    case "setSpacing": {
      const blockEl = ids.elOf(intent.blockId);
      if (!blockEl) return false;
      const target = firstTextDescendant(blockEl) ?? blockEl;
      return setParagraphSpacing(doc, [target], intent.patch as never, suggestMeta(doc, intent.suggest));
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
    default: {
      const exhaustive: never = intent;
      return exhaustive;
    }
  }
}

/**
 * Resolve a registered operation's stable-id address to a document target.
 * A null result IS the collaborative "honest no-op": the address named
 * content this replica cannot see, so nothing is mutated and every replica
 * records the same rejection.
 */
function resolveOperationTarget(
  doc: DocxDocument,
  ids: StableIds,
  runOf: RunLookup,
  address: OperationAddress,
  intent: RegisteredIntent,
): OperationTarget | null {
  // A document-scoped operation names no node: the document IS the target, so
  // this address always resolves and the operation's own payload has to carry
  // its rejection predicate (see the registry's OperationAddress comment).
  const empty = { t: null, run: null, cellParagraph: null, drawing: null };
  if (address === "document") return { ...empty, el: doc.docRoot };
  const addressId = (intent as unknown as Record<string, number>)[ADDRESS_WIRE_FIELD[address]];
  switch (address) {
    case "run": {
      const el = ids.elOf(addressId);
      if (!el) return null;
      const entry = runOf(el);
      if (!entry) return null;
      return { ...empty, el, t: entry.firstT ?? null, run: entry.run };
    }
    case "block": {
      const el = ids.elOf(addressId);
      if (!el) return null;
      return { ...empty, el, t: firstTextDescendant(el) };
    }
    case "cell": {
      const paraEl = ids.elOf(addressId) ?? null;
      const tbl = paraEl ? tableOf(doc, paraEl) : null;
      return tbl ? { ...empty, el: tbl, cellParagraph: paraEl } : null;
    }
    case "object": {
      // BOTH halves have to resolve. A run whose contents moved under a
      // concurrent edit rejects on every replica rather than mutating
      // whatever now sits at that index — the honest no-op an unresolvable
      // id already gives, extended to the index that narrows it.
      const el = ids.elOf(addressId);
      if (!el) return null;
      const objectIndex = (intent as unknown as { objectIndex?: number }).objectIndex;
      const drawing = drawingIn(el, objectIndex, runOf);
      return drawing ? { ...empty, el, drawing } : null;
    }
  }
}

/** Apply a registered operation: resolve its address, run the declared
 * mutation, retire the ids of anything it removed, and hand the nodes it
 * created their carried ids. */
function applyRegistered(
  doc: DocxDocument,
  ids: StableIds,
  runOf: RunLookup,
  intent: RegisteredIntent,
): boolean {
  const definition = registeredOperation(intent.kind);
  if (!definition) return false;
  const target = resolveOperationTarget(doc, ids, runOf, definition.address, intent);
  if (!target) return false;
  const nodeIds = "nodeIds" in intent ? intent.nodeIds : undefined;
  const before = nodeIds ? trackedSet(ids, doc) : null;
  if (!applyRegisteredOperation(doc, target, intent)) return false;
  // Prune BEFORE handing out carried ids, so a replacing operation
  // (insertWatermark) leaves no id pointing at a detached run.
  if (definition.prunesIds) ids.prune(doc.editableRoots());
  if (before && nodeIds) assignFreshTracked(ids, doc, before, nodeIds);
  return true;
}

/**
 * Assign carried ids to tracked nodes created since `before`, in doc order.
 *
 * The fresh nodes' AUTO ids are dropped first. A refresh inside the mutation
 * auto-assigns sequential ids starting at the table's `next`, and after an
 * earlier intent whose batch was only partly consumed, `next` sits INSIDE
 * that batch's range — so an auto id here can equal a LATER id of THIS
 * intent's carried batch, and reassign would reject a perfectly good intent
 * as a collision (found by the index wave: mark an XE entry, then
 * insertIndex). Dropping the autos first makes the batch land regardless;
 * fresh nodes beyond the batch re-auto-assign ABOVE the carried maximum
 * (reassign bumps `next`), which every replica derives identically.
 */
function assignFreshTracked(ids: StableIds, doc: DocxDocument, before: Set<XmlElement>, nodeIds: number[]): void {
  const fresh: XmlElement[] = [];
  walkTracked(doc, (el) => { if (!before.has(el)) fresh.push(el); });
  for (const el of fresh) ids.unassign(el);
  for (let k = 0; k < fresh.length; k++) {
    if (k < nodeIds.length) ids.reassign(fresh[k], nodeIds[k]);
    else ids.assign(fresh[k]);
  }
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

/** The table containing an id-addressed cell paragraph — the NEAREST tbl
 * ancestor, so a nested table's paragraph addresses the inner table (senders
 * pick a paragraph whose nearest tbl is the intended one). */
function tableOfParagraph(doc: DocxDocument, ids: StableIds, cellParagraphId: number): XmlElement | null {
  const paraEl = ids.elOf(cellParagraphId);
  return paraEl ? tableOf(doc, paraEl) : null;
}

/** The nearest w:tbl ancestor of (or equal to) an element. */
function tableOf(doc: DocxDocument, el: XmlElement): XmlElement | null {
  for (let cur: XmlElement | null = el; cur; cur = doc.findParentOf(cur) ?? null) {
    if (localName(cur.name) === "tbl") return cur;
  }
  return null;
}

/** The first DrawingML or legacy VML drawing inside a run's subtree. */
function firstDrawingIn(el: XmlElement): XmlElement | null {
  if (["drawing", "shape", "rect", "line"].includes(localName(el.name))) return el;
  for (const c of el.children) {
    const found = firstDrawingIn(c);
    if (found) return found;
  }
  return null;
}

function drawingSource(content: Run["content"][number]): XmlElement | undefined {
  if (content.kind === "image" || content.kind === "drawing") return content.srcDrawing;
  if (content.kind !== "anchor") return undefined;
  const shape = content.shape;
  if (shape.type === "wordart") return shape.src;
  if (shape.type === "line") return shape.src;
  return shape.srcDrawing;
}

function drawingIn(runEl: XmlElement, objectIndex: number | undefined, runOf: RunLookup): XmlElement | null {
  if (objectIndex === undefined) return firstDrawingIn(runEl);
  const content = runOf(runEl)?.run.content[objectIndex];
  return content ? drawingSource(content) ?? null : null;
}

/** Every m:oMath under `el`, in document order. Must stay identical to the
 * editor's helper of the same name: both sides have to count the same
 * equations in the same order for an edit to land on the one the author
 * touched. */
function mathsIn(el: XmlElement): XmlElement[] {
  if (localName(el.name) === "oMath") return [el];
  return el.children.flatMap(mathsIn);
}

/** The equation a math intent names: the `index`th under the block. */
function mathIn(el: XmlElement, index = 0): XmlElement | null {
  return mathsIn(el)[index] ?? null;
}

interface RunEntry {
  run: Run;
  firstT: XmlElement | undefined;
}

/** Look up a run's source XML element in the model (needed for the checkbox
 * guard in applyInsertText and to locate the run's first w:t). */
type RunLookup = (runEl: XmlElement, blockIdHint?: number) => RunEntry | undefined;

function runEntry(run: Run): RunEntry {
  return { run, firstT: run.content.find((c) => c.kind === "text")?.srcT };
}

/**
 * A run lookup that resolves ONE run instead of indexing the whole model.
 *
 * The old form built a Map over every run in the document on every apply — an
 * O(document) walk to answer a single-character insert (perf B9). Nearly every
 * intent addresses a run whose paragraph the intent also names, so the hinted
 * path indexes just that paragraph. Anything else (drawing/table/comment
 * intents that carry only a run id, or a hint that doesn't hold the run) falls
 * back to the full index, built at most once per apply — same answers as
 * before, off the same model, just not paid for by every keystroke.
 */
function makeRunLookup(doc: DocxDocument, ids: StableIds): RunLookup {
  let full: Map<XmlElement, RunEntry> | null = null;
  const scoped = new Map<XmlElement, RunEntry>();
  return (runEl, blockIdHint) => {
    if (!full) {
      const cached = scoped.get(runEl);
      if (cached) return cached;
      if (blockIdHint !== undefined) {
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

/** True when `needle` is `el` or sits somewhere under it. */
function contains(el: XmlElement, needle: XmlElement): boolean {
  if (el === needle) return true;
  for (const c of el.children) if (contains(c, needle)) return true;
  return false;
}

/**
 * Block scope for a text-level intent: the paragraph it addressed, but only
 * once that paragraph is CONFIRMED to hold the addressed run. The two ids
 * arrive together on the wire from a sender that may be buggy or hostile, and
 * a scope naming the wrong paragraph would leave the real one's model stale —
 * a silent, permanent divergence. Anything unverifiable (the block id resolves
 * to a table, to nothing, or to a paragraph that doesn't hold the run) falls
 * back to document scope, which is always correct.
 */
function blockScope(ids: StableIds, blockId: number, runId: number): Scope {
  const blockEl = ids.elOf(blockId);
  const runEl = ids.elOf(runId);
  if (!blockEl || !runEl || localName(blockEl.name) !== "p") return DOC_SCOPE;
  return contains(blockEl, runEl) ? { kind: "block", blocks: [blockEl] } : DOC_SCOPE;
}

/** Index every run in the model by its source element — body story AND the
 * header/footer parts, including the text-box stories hanging off either.
 *
 * The hf half is not an optimization: header paragraphs carry stable ids (they
 * ride in doc.editableRoots(), so the id table and the checkpoint sidecar have
 * always covered them) but they are parsed into doc.headers/doc.footers, NOT
 * into doc.sections. Walking sections alone left every run in a header
 * unresolvable, so resolveCaret returned null and EVERY text/format intent
 * aimed at a header was a clean reject on the server and on every peer — while
 * the originator, which applies locally before emitting, kept the edit. That is
 * the silent fork the header/footer toolbar group was gated off to avoid.
 *
 * The text-box half is the SAME bug one level deeper, and it is how pleading
 * paper draws its margin line numbers: the number column is a w:txbxContent
 * story in the header, not w:lnNumType. Those story blocks hang off the RUN
 * that carries the drawing (run.content anchor -> shape.blocks), which neither
 * the section walk nor the hf walk descends into. The editor lets a caret into
 * that story (parse marks it textboxStory, and its runs are id-tracked), so a
 * user could type in the gutter and watch every keystroke get rejected by the
 * server and every peer — "I can't change the number column". */
function buildRunMap(doc: DocxDocument): Map<XmlElement, RunEntry> {
  const map = new Map<XmlElement, RunEntry>();
  const visitRun = (run: Run): void => {
    if (run.src) map.set(run.src, runEntry(run));
    // A text box's story is a little document of its own. Recursing through
    // visitBlocks also covers a text box nested inside another one's story.
    for (const c of run.content) {
      if (c.kind === "anchor" && c.shape.type === "textbox") visitBlocks(c.shape.blocks);
    }
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
  for (const hf of doc.headers.values()) visitBlocks(hf.blocks);
  for (const hf of doc.footers.values()) visitBlocks(hf.blocks);
  return map;
}

/** Resolve a wire position (offset in the run's WIRE space — cumulative with
 * separators counted, checkpoint B1 rev 2) to the concrete (w:t, local
 * offset) it falls in. Strict bounds: a position past the run's wire length
 * is unresolvable (clean rejection), unlike the caret-restore decodeCaret
 * which clamps. */
function resolveCaret(ids: StableIds, runOf: RunLookup, pos: Position): EditCaret | null {
  const runEl = ids.elOf(pos.runId);
  if (!runEl) return null;
  const entry = runOf(runEl, pos.blockId);
  if (!entry || !entry.firstT) return null;
  if (pos.offset < 0 || pos.offset > runWireLength(runEl)) return null;
  const hit = resolveRunOffset(runEl, pos.offset);
  if (!hit) return null;
  return { t: hit.t, run: entry.run, offset: hit.offset };
}
