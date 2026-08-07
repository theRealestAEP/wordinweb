import { XmlElement, cloneXml, localName } from "../xml.js";
import { Run } from "../model.js";
import { DocxDocument } from "../docx.js";
import { checkboxStateElement } from "../checkbox.js";
import { insertSuggestedText, markParagraphGlyph, RevisionMeta } from "./suggest.js";

/**
 * Editor-level document mutations, extracted from DocxEditor so they operate
 * on an explicit caret rather than ambient editor state. Replicated editing
 * applies these on the server and on peers with the same code the local
 * editor runs (plan doc 02): they touch only the XML tree and the caret —
 * never the DOM, layout, or view state — so they run headlessly in Node.
 *
 * Each returns the post-edit caret (or null / the unchanged caret for a
 * no-op). Callers reassign their caret from the result and handle
 * commit/rerender themselves.
 */

/** A caret as a model position: the w:t element it sits in, the cached run,
 * a char offset within `t.text`, and a wrap-boundary affinity. Structurally
 * identical to DocxEditor's private Caret. */
export interface EditCaret {
  t: XmlElement;
  run: Run;
  offset: number;
  bias?: "end";
}

/** Context a mutation needs from the editor beyond the caret. */
export interface MutationCtx {
  /** Suggesting (tracked-changes) mode: edits are recorded as w:ins/w:del. */
  suggesting: boolean;
  /** Revision metadata factory (author/date/id) for suggested edits. Called
   * at most once per mutation, when a revision element is created. */
  revMeta: () => RevisionMeta;
}

/**
 * Insert text at the caret (suggesting-aware). Mirrors the former
 * DocxEditor.insertTextCore: a checkbox glyph is atomic (no-op); suggesting
 * mode wraps the insertion in a revision; otherwise the text is spliced into
 * the caret's w:t. Returns the caret after the insertion (unchanged on no-op).
 */
export function applyInsertText(
  doc: DocxDocument,
  caret: EditCaret,
  text: string,
  ctx: MutationCtx,
): EditCaret {
  // A checkbox content control's glyph is atomic: typing never edits it.
  if (checkboxStateElement(caret.run, caret.t)) return caret;
  if (ctx.suggesting) {
    const nc = insertSuggestedText(doc, caret.t, caret.offset, text, ctx.revMeta());
    if (nc) return { t: nc.t, run: caret.run, offset: nc.offset, bias: "end" };
    return caret;
  }
  const t = caret.t;
  t.text = t.text.slice(0, caret.offset) + text + t.text.slice(caret.offset);
  return { t, run: caret.run, offset: caret.offset + text.length, bias: "end" };
}

/**
 * Delete `[from, to)` characters within the caret's single w:t. Mirrors the
 * text-removal half of DocxEditor.deleteContents (the grapheme-cluster
 * boundary is resolved by the caller and passed as explicit offsets — so the
 * replicated form carries engine-independent offsets and never re-derives
 * `Intl.Segmenter` boundaries that vary across runtimes, plan doc 05 M2).
 * Returns the caret at `from`. A checkbox glyph is atomic (no-op).
 */
export function applyDeleteRange(
  caret: EditCaret,
  from: number,
  to: number,
): EditCaret {
  if (checkboxStateElement(caret.run, caret.t)) return caret;
  const lo = Math.max(0, Math.min(from, to));
  const hi = Math.min(caret.t.text.length, Math.max(from, to));
  if (lo >= hi) return caret;
  caret.t.text = caret.t.text.slice(0, lo) + caret.t.text.slice(hi);
  return { t: caret.t, run: caret.run, offset: lo, bias: "end" };
}

/** Result of a paragraph split: the two sibling paragraph elements and the
 * caret placed at the start of the second. */
export interface SplitResult {
  before: XmlElement;
  after: XmlElement;
  caret: EditCaret;
}

/**
 * Split the caret's paragraph at the caret (Enter). Mirrors the former
 * DocxEditor.splitParagraphCore: text after the caret and any following runs
 * move to a new sibling paragraph with cloned pPr (section break stripped);
 * suggesting mode records the introduced paragraph mark. Returns null when
 * the caret is not inside a run inside a paragraph.
 */
export function applySplitParagraph(
  doc: DocxDocument,
  caret: EditCaret,
  ctx: MutationCtx,
): SplitResult | null {
  // Resolve containers from the w:t itself — cached run/model objects go
  // stale after any refresh, but the t element's identity is durable.
  const rEl = doc.findParentOf(caret.t);
  if (!rEl || localName(rEl.name) !== "r") return null;
  let pEl: XmlElement | undefined = doc.findParentOf(rEl);
  while (pEl && localName(pEl.name) !== "p") pEl = doc.findParentOf(pEl);
  if (!pEl) return null;
  const pParent = doc.findParentOf(pEl);
  if (!pParent) return null;
  const runIdx = pEl.children.indexOf(rEl);
  const tIdx = rEl.children.indexOf(caret.t);
  if (tIdx === -1) return null;

  const prefix = pEl.name.includes(":") ? pEl.name.slice(0, pEl.name.indexOf(":") + 1) : "";
  const rPr = rEl.children.find((c) => localName(c.name) === "rPr");

  let afterT: XmlElement;
  let afterRun: XmlElement;
  /** Index in pEl.children after which content moves to the new paragraph. */
  let splitIdx = runIdx;
  if (runIdx !== -1) {
    // Split the caret run: text after the caret moves to a new run.
    afterT = {
      name: caret.t.name,
      attrs: { ...caret.t.attrs, "xml:space": "preserve" },
      text: caret.t.text.slice(caret.offset),
      children: [],
    };
    afterRun = {
      name: rEl.name,
      attrs: { ...rEl.attrs },
      text: "",
      children: [...(rPr ? [cloneXml(rPr)] : []), afterT, ...rEl.children.slice(tIdx + 1)],
    };
    caret.t.text = caret.t.text.slice(0, caret.offset);
    rEl.children = rEl.children.slice(0, tIdx + 1);
  } else {
    // The caret run rides inside a tracked-insert wrapper (w:ins) — the user
    // splits right after a suggested insertion (Enter, or a multi-line paste
    // in suggesting mode). Only the END-of-wrapper split has a simple form:
    // nothing after the caret needs to leave the wrapper, so the new
    // paragraph seeds with a plain empty run and takes the following
    // siblings. A mid-suggestion split stays a no-op, as before.
    const wrapper = doc.findParentOf(rEl);
    const wrapperIdx = wrapper ? pEl.children.indexOf(wrapper) : -1;
    if (!wrapper || localName(wrapper.name) !== "ins" || wrapperIdx === -1) return null;
    if (caret.offset < caret.t.text.length || tIdx < rEl.children.length - 1) return null;
    if (wrapper.children.indexOf(rEl) !== wrapper.children.length - 1) return null;
    afterT = {
      name: caret.t.name,
      attrs: { ...caret.t.attrs, "xml:space": "preserve" },
      text: "",
      children: [],
    };
    afterRun = {
      name: rEl.name,
      attrs: { ...rEl.attrs },
      text: "",
      children: [...(rPr ? [cloneXml(rPr)] : []), afterT],
    };
    splitIdx = wrapperIdx;
  }

  // New paragraph: cloned pPr (minus any section break!) + moved content.
  const pPrEl = pEl.children.find((c) => localName(c.name) === "pPr");
  const newPPr = pPrEl ? cloneXml(pPrEl) : undefined;
  if (newPPr) {
    newPPr.children = newPPr.children.filter((c) => localName(c.name) !== "sectPr");
  }
  const moved = pEl.children.slice(splitIdx + 1);
  pEl.children = pEl.children.slice(0, splitIdx + 1);
  const newP: XmlElement = {
    name: prefix + "p",
    attrs: {},
    text: "",
    children: [...(newPPr ? [newPPr] : []), afterRun, ...moved],
  };
  const pIdx = pParent.children.indexOf(pEl);
  pParent.children.splice(pIdx + 1, 0, newP);
  // Hand the tree's new links to the parent memo. Enter at the end of the
  // document is the seeding shape (each split addresses the paragraph the
  // previous one created), and without this every one of those splits pays a
  // full-document walk to rediscover what we just spliced (perf B9).
  doc.noteParent(newP, pParent);
  doc.noteParent(afterRun, newP);
  doc.noteParent(afterT, afterRun);
  for (const child of moved) doc.noteParent(child, newP);

  // Suggesting mode: the split introduces a new paragraph mark at the end of
  // the FIRST paragraph — record it as an inserted glyph (pPr/rPr/w:ins).
  if (ctx.suggesting) markParagraphGlyph(pEl, "ins", ctx.revMeta());

  return { before: pEl, after: newP, caret: { t: afterT, run: caret.run, offset: 0 } };
}
