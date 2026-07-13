import { DocxDocument } from "../docx.js";
import { XmlElement, cloneXml, localName } from "../xml.js";

/**
 * Suggesting mode — "my own tracked changes".
 *
 * When the editor is in suggesting mode, edits are recorded as OOXML revisions
 * instead of mutating text in place: inserted text is wrapped in `w:ins`,
 * deleted content is converted to `w:del`/`w:delText` (never actually removed),
 * and paragraph splits/merges mark the paragraph glyph via `pPr/rPr/w:ins` or
 * `pPr/rPr/w:del`. The markup renderer (parse/document.ts) already colors and
 * underlines/strikes these, so pending suggestions show live as you type.
 *
 * The XML tree stays the source of truth; every helper mutates it and returns
 * the new caret target so the editor can re-place its own caret after refresh.
 * Accept/reject unwrap or remove a single revision (see acceptRevision).
 */

export interface RevisionMeta {
  author: string;
  /** ISO 8601 timestamp for the w:date attribute. */
  date: string;
  /** Allocates a fresh, unique w:id. */
  nextId: () => number;
}

/** Where an edit left the caret (the editor re-resolves this after refresh). */
export interface CaretTarget {
  t: XmlElement;
  offset: number;
}

/** A tracked-change element the review UI can accept or reject. */
export type RevisionKind = "insertion" | "deletion" | "markInsertion" | "markDeletion";
export interface RevisionRef {
  /** The w:ins / w:del element (run-level, or the mark ins/del inside rPr). */
  el: XmlElement;
  kind: RevisionKind;
  /** For mark revisions: the owning w:p. */
  paragraph?: XmlElement;
  author?: string;
}

// ---------- small XML builders ----------

function prefixOf(e: XmlElement): string {
  return e.name.includes(":") ? e.name.slice(0, e.name.indexOf(":") + 1) : "";
}

function el(name: string, attrs: Record<string, string> = {}, children: XmlElement[] = [], text = ""): XmlElement {
  return { name, attrs, children, text };
}

function attrByLocal(e: XmlElement, local: string): string | undefined {
  const key = Object.keys(e.attrs).find((k) => localName(k) === local);
  return key ? e.attrs[key] : undefined;
}

function revAttrs(w: string, meta: RevisionMeta): Record<string, string> {
  return {
    [`${w}id`]: String(meta.nextId()),
    [`${w}author`]: meta.author,
    [`${w}date`]: meta.date,
  };
}

/** A w:r cloning rEl's prefix/attrs/rPr, carrying the given content children. */
function runLike(rEl: XmlElement, rPr: XmlElement | undefined, content: XmlElement[]): XmlElement {
  return {
    name: rEl.name,
    attrs: { ...rEl.attrs },
    text: "",
    children: [...(rPr ? [cloneXml(rPr)] : []), ...content],
  };
}

/** A w:t / w:delText element cloning t's name/attrs with new text. */
function textLike(t: XmlElement, name: string, s: string): XmlElement {
  return { name, attrs: { ...t.attrs, "xml:space": "preserve" }, text: s, children: [] };
}

function findRun(doc: DocxDocument, t: XmlElement): { rEl: XmlElement; parent: XmlElement } | null {
  const rEl = doc.findParentOf(t);
  if (!rEl || localName(rEl.name) !== "r") return null;
  const parent = doc.findParentOf(rEl);
  if (!parent) return null;
  return { rEl, parent };
}

/** True when rEl sits directly inside a w:ins/w:del authored by `author`. */
function ownRevisionWrapper(parent: XmlElement, kind: "ins" | "del", author: string): boolean {
  return localName(parent.name) === kind && attrByLocal(parent, "author") === author;
}

// ---------- insertion ----------

/**
 * Record `text` as a tracked insertion at (t, offset). Contiguous typing inside
 * one's own pending w:ins just extends its w:t (so a keystroke burst is a single
 * revision); otherwise the containing run is split and a new w:ins is threaded
 * between the halves. Returns the caret position after the inserted text.
 */
export function insertSuggestedText(
  doc: DocxDocument,
  t: XmlElement,
  offset: number,
  text: string,
  meta: RevisionMeta,
): CaretTarget | null {
  if (text.length === 0) return { t, offset };
  const found = findRun(doc, t);
  if (!found) return null;
  const { rEl, parent } = found;

  // Fast path: already typing inside my own insertion — extend it in place.
  if (ownRevisionWrapper(parent, "ins", meta.author)) {
    t.text = t.text.slice(0, offset) + text + t.text.slice(offset);
    return { t, offset: offset + text.length };
  }

  const w = prefixOf(rEl);
  const rPr = rEl.children.find((c) => localName(c.name) === "rPr");

  // Caret inside a w:del (typing right after deleting a run of text): an
  // insertion may not nest inside a deletion, so drop the new w:ins as a
  // sibling just before (offset 0) or after (offset at end) the w:del.
  if (localName(parent.name) === "del") {
    const grand = doc.findParentOf(parent);
    if (!grand) return null;
    const delIdx = grand.children.indexOf(parent);
    const newT = textLike(t, t.name, text);
    const ins = el(`${w}ins`, revAttrs(w, meta), [runLike(rEl, rPr, [newT])]);
    const at = offset <= 0 ? delIdx : delIdx + 1;
    grand.children.splice(at, 0, ins);
    return { t: newT, offset: text.length };
  }

  const tIdx = rEl.children.indexOf(t);
  const idx = parent.children.indexOf(rEl);
  if (tIdx === -1 || idx === -1) return null;

  const preceding = rEl.children.slice(0, tIdx).filter((c) => localName(c.name) !== "rPr");
  const following = rEl.children.slice(tIdx + 1);

  const replacement: XmlElement[] = [];
  const beforeChildren = [...preceding];
  if (offset > 0) beforeChildren.push(textLike(t, t.name, t.text.slice(0, offset)));
  if (beforeChildren.length > 0) replacement.push(runLike(rEl, rPr, beforeChildren));

  const newT = textLike(t, t.name, text);
  const ins = el(`${w}ins`, revAttrs(w, meta), [runLike(rEl, rPr, [newT])]);
  replacement.push(ins);

  const afterChildren: XmlElement[] = [];
  if (offset < t.text.length) afterChildren.push(textLike(t, t.name, t.text.slice(offset)));
  afterChildren.push(...following);
  if (afterChildren.length > 0) replacement.push(runLike(rEl, rPr, afterChildren));

  parent.children.splice(idx, 1, ...replacement);
  return { t: newT, offset: text.length };
}

// ---------- deletion ----------

/** A (t, start, end) slice of text to record as deleted. */
export interface DeleteRange {
  t: XmlElement;
  start: number;
  end: number;
}

/**
 * Record a set of text ranges as tracked deletions. Text inside one's own
 * pending w:ins is physically removed instead (you are un-suggesting your own
 * not-yet-accepted insertion, which Word also does). Everything else is split
 * out of its run and wrapped in w:del with its w:t retyped as w:delText.
 *
 * Ranges are processed right-to-left within each run so offsets stay valid.
 * Returns the caret position at the left edge of the affected region.
 */
export function deleteSuggestedRange(
  doc: DocxDocument,
  ranges: DeleteRange[],
  meta: RevisionMeta,
): CaretTarget | null {
  // Group by w:t, delete from the end so earlier offsets stay valid.
  const byT = new Map<XmlElement, { start: number; end: number }[]>();
  for (const r of ranges) {
    if (r.end <= r.start) continue;
    const list = byT.get(r.t) ?? [];
    list.push({ start: r.start, end: r.end });
    byT.set(r.t, list);
  }
  // Segments arrive in document order, so the FIRST w:t is the leftmost. Its
  // left boundary is where the caret belongs after the deletion; later w:t
  // groups never move it rightward.
  let caret: CaretTarget | null = null;
  for (const [t, rs] of byT) {
    // Right-to-left within the run keeps offsets valid; the last (leftmost)
    // range resolved for the first run gives the caret.
    rs.sort((a, b) => b.start - a.start);
    let runCaret: CaretTarget | null = null;
    for (const r of rs) {
      const c = deleteOneRange(doc, t, r.start, r.end, meta);
      if (c) runCaret = c;
    }
    if (caret === null && runCaret) caret = runCaret;
  }
  return caret;
}

function deleteOneRange(
  doc: DocxDocument,
  t: XmlElement,
  start: number,
  end: number,
  meta: RevisionMeta,
): CaretTarget | null {
  start = Math.max(0, Math.min(start, t.text.length));
  end = Math.max(start, Math.min(end, t.text.length));
  if (start === end) return { t, offset: start };
  const found = findRun(doc, t);
  if (!found) return null;
  const { rEl, parent } = found;

  // Removing my own pending insertion: just splice the text out.
  if (ownRevisionWrapper(parent, "ins", meta.author)) {
    t.text = t.text.slice(0, start) + t.text.slice(end);
    // An emptied w:t leaves an empty run/ins; harmless (renders nothing) and
    // cheaper than restructuring. Caret sits where the text was.
    return { t, offset: start };
  }

  const w = prefixOf(rEl);
  const rPr = rEl.children.find((c) => localName(c.name) === "rPr");
  const tIdx = rEl.children.indexOf(t);
  const idx = parent.children.indexOf(rEl);
  if (tIdx === -1 || idx === -1) return null;

  const preceding = rEl.children.slice(0, tIdx).filter((c) => localName(c.name) !== "rPr");
  const following = rEl.children.slice(tIdx + 1);
  const delTextName = w + "delText";

  const replacement: XmlElement[] = [];
  let beforeT: XmlElement | null = null;
  const beforeChildren = [...preceding];
  if (start > 0) {
    beforeT = textLike(t, t.name, t.text.slice(0, start));
    beforeChildren.push(beforeT);
  }
  if (beforeChildren.length > 0) replacement.push(runLike(rEl, rPr, beforeChildren));

  const deletedT = textLike(t, delTextName, t.text.slice(start, end));
  const del = el(`${w}del`, revAttrs(w, meta), [runLike(rEl, rPr, [deletedT])]);
  replacement.push(del);

  const afterChildren: XmlElement[] = [];
  if (end < t.text.length) afterChildren.push(textLike(t, t.name, t.text.slice(end)));
  afterChildren.push(...following);
  if (afterChildren.length > 0) replacement.push(runLike(rEl, rPr, afterChildren));

  parent.children.splice(idx, 1, ...replacement);
  coalesceDeletions(parent, del, meta.author);

  // Caret at the left edge of the struck text so the next Backspace keeps
  // eating leftward. Prefer the end of the surviving "before" run.
  if (beforeT) return { t: beforeT, offset: beforeT.text.length };
  return { t: deletedT, offset: 0 };
}

/** Merge a fresh w:del with adjacent same-author w:del siblings (keeps the XML
 * tidy across a Backspace burst — multiple <w:del> are valid but Word coalesces). */
function coalesceDeletions(parent: XmlElement, del: XmlElement, author: string): void {
  const isMyDel = (e: XmlElement | undefined): boolean =>
    !!e && localName(e.name) === "del" && attrByLocal(e, "author") === author;
  let i = parent.children.indexOf(del);
  if (i === -1) return;
  // Absorb the next sibling into `del`.
  while (isMyDel(parent.children[i + 1])) {
    const next = parent.children[i + 1];
    del.children.push(...next.children);
    parent.children.splice(i + 1, 1);
  }
  // Absorb `del` into a previous sibling.
  while (isMyDel(parent.children[i - 1])) {
    const prev = parent.children[i - 1];
    prev.children.push(...del.children);
    parent.children.splice(i, 1);
    del = prev;
    i -= 1;
  }
}

// ---------- paragraph mark revisions ----------

function ensurePPrRPr(pEl: XmlElement): XmlElement {
  const w = prefixOf(pEl);
  let pPr = pEl.children.find((c) => localName(c.name) === "pPr");
  if (!pPr) {
    pPr = el(`${w}pPr`);
    pEl.children.unshift(pPr);
  }
  let rPr = pPr.children.find((c) => localName(c.name) === "rPr");
  if (!rPr) {
    rPr = el(`${w}rPr`);
    // rPr is the last child of pPr in the schema; append is safe.
    pPr.children.push(rPr);
  }
  return rPr;
}

/** Mark a paragraph's end glyph as an inserted (w:ins) or deleted (w:del)
 * revision. Used for Enter (split → the first paragraph's new mark is inserted)
 * and Backspace/Delete across a paragraph boundary (the joined mark is deleted). */
export function markParagraphGlyph(
  pEl: XmlElement,
  kind: "ins" | "del",
  meta: RevisionMeta,
): XmlElement {
  const w = prefixOf(pEl);
  const rPr = ensurePPrRPr(pEl);
  // Replace any existing mark revision of this kind (idempotent).
  const existingIdx = rPr.children.findIndex((c) => localName(c.name) === kind);
  const mark = el(`${w}${kind}`, revAttrs(w, meta));
  if (existingIdx !== -1) rPr.children[existingIdx] = mark;
  else rPr.children.unshift(mark); // ins/del lead rPr in the schema
  return mark;
}

/** Whether a paragraph's mark already carries a w:ins/w:del of `kind`. */
export function paragraphGlyphRevision(pEl: XmlElement, kind: "ins" | "del"): XmlElement | null {
  const pPr = pEl.children.find((c) => localName(c.name) === "pPr");
  const rPr = pPr?.children.find((c) => localName(c.name) === "rPr");
  return rPr?.children.find((c) => localName(c.name) === kind) ?? null;
}

// ---------- review: locate / accept / reject ----------

/** The run-level revision (w:ins/w:del) enclosing a w:t, if any. */
export function revisionForText(doc: DocxDocument, t: XmlElement): RevisionRef | null {
  let cur: XmlElement | undefined = doc.findParentOf(t);
  while (cur) {
    const ln = localName(cur.name);
    if (ln === "ins" || ln === "del") {
      return {
        el: cur,
        kind: ln === "ins" ? "insertion" : "deletion",
        author: attrByLocal(cur, "author"),
      };
    }
    if (ln === "p" || ln === "body" || ln === "tc") break;
    cur = doc.findParentOf(cur);
  }
  return null;
}

function convertDelTextToText(el: XmlElement): void {
  const walk = (e: XmlElement): void => {
    if (localName(e.name) === "delText") {
      e.name = e.name.replace(/delText$/, "t");
    }
    for (const c of e.children) walk(c);
  };
  walk(el);
}

/** Replace `el` in its parent with its own children (unwrap). */
function unwrap(doc: DocxDocument, el: XmlElement): boolean {
  const parent = doc.findParentOf(el);
  if (!parent) return false;
  const idx = parent.children.indexOf(el);
  if (idx === -1) return false;
  parent.children.splice(idx, 1, ...el.children);
  return true;
}

function removeEl(doc: DocxDocument, el: XmlElement): boolean {
  const parent = doc.findParentOf(el);
  if (!parent) return false;
  const idx = parent.children.indexOf(el);
  if (idx === -1) return false;
  parent.children.splice(idx, 1);
  return true;
}

/** Join pEl with the paragraph that follows it (accept a deleted mark / reject
 * an inserted mark): move the next paragraph's runs into pEl, drop the next. */
function joinWithNext(doc: DocxDocument, pEl: XmlElement): boolean {
  const parent = doc.findParentOf(pEl);
  if (!parent) return false;
  const idx = parent.children.indexOf(pEl);
  const next = parent.children[idx + 1];
  if (!next || localName(next.name) !== "p") return false;
  const moved = next.children.filter((c) => localName(c.name) !== "pPr");
  pEl.children.push(...moved);
  parent.children.splice(idx + 1, 1);
  return true;
}

/**
 * Accept a single revision: an insertion becomes permanent text, a deletion
 * removes its content, an inserted mark stays split, a deleted mark merges the
 * two paragraphs. Returns false if the element could not be located.
 */
export function acceptRevision(doc: DocxDocument, ref: RevisionRef): boolean {
  let ok = false;
  switch (ref.kind) {
    case "insertion":
      ok = unwrap(doc, ref.el);
      break;
    case "deletion":
      ok = removeEl(doc, ref.el);
      break;
    case "markInsertion":
      ok = removeMarkRevision(ref);
      break;
    case "markDeletion":
      if (ref.paragraph) {
        removeMarkRevision(ref);
        ok = joinWithNext(doc, ref.paragraph);
      }
      break;
  }
  if (ok) doc.refresh();
  return ok;
}

/**
 * Reject a single revision: an insertion disappears, a deletion's content is
 * restored, an inserted mark un-splits (the paragraphs rejoin), a deleted mark
 * is removed (the paragraphs stay separate). Returns false if not located.
 */
export function rejectRevision(doc: DocxDocument, ref: RevisionRef): boolean {
  let ok = false;
  switch (ref.kind) {
    case "insertion":
      ok = removeEl(doc, ref.el);
      break;
    case "deletion":
      convertDelTextToText(ref.el);
      ok = unwrap(doc, ref.el);
      break;
    case "markInsertion":
      if (ref.paragraph) {
        removeMarkRevision(ref);
        ok = joinWithNext(doc, ref.paragraph);
      }
      break;
    case "markDeletion":
      ok = removeMarkRevision(ref);
      break;
  }
  if (ok) doc.refresh();
  return ok;
}

/** Remove a mark revision element (the w:ins/w:del inside pPr/rPr). */
function removeMarkRevision(ref: RevisionRef): boolean {
  const rPr = ref.paragraph
    ? ref.paragraph.children.find((c) => localName(c.name) === "pPr")?.children.find((c) => localName(c.name) === "rPr")
    : undefined;
  if (!rPr) return false;
  const idx = rPr.children.indexOf(ref.el);
  if (idx === -1) return false;
  rPr.children.splice(idx, 1);
  return true;
}
