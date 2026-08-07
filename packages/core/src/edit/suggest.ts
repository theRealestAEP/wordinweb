import { DocxDocument, NotePart } from "../docx.js";
import { XmlElement, cloneXml, localName } from "../xml.js";

/**
 * Suggesting mode — "my own tracked changes".
 *
 * When the editor is in suggesting mode, edits are recorded as OOXML revisions
 * instead of mutating text in place: inserted text is wrapped in `w:ins`,
 * deleted content is converted to `w:del`/`w:delText` (never actually removed),
 * paragraph splits/merges mark the paragraph glyph via `pPr/rPr/w:ins` or
 * `pPr/rPr/w:del`, and formatting changes stash the properties they replace in
 * a `*PrChange` — `w:rPrChange` and `w:pPrChange` for a run and a paragraph,
 * `w:tblPrChange`, `w:trPrChange` and `w:tcPrChange` for a table, a row and a
 * cell. The markup renderer (parse/document.ts) already
 * colors and underlines/strikes the text ones, so pending suggestions show live
 * as you type; the format ones show as the new formatting, like Word, because
 * the renderer reads only direct property children and never descends into a
 * change record.
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

/**
 * RevisionMeta for the author/date a wire payload carries. Nondeterministic
 * values are generated once by the originating client and travel with the
 * edit; the w:id is scan-derived from identical tree state, so every replica
 * writes byte-identical XML. Undefined `suggest` means a direct (untracked)
 * edit, which every suggest-aware mutation takes as "no meta".
 */
export function suggestMeta(
  doc: DocxDocument,
  suggest: { author: string; date: string } | undefined,
): RevisionMeta | undefined {
  if (!suggest) return undefined;
  return { author: suggest.author, date: suggest.date, nextId: () => doc.nextRevisionId() };
}

/** Where an edit left the caret (the editor re-resolves this after refresh). */
export interface CaretTarget {
  t: XmlElement;
  offset: number;
}

/**
 * A tracked FORMATTING change: the five properties elements that can carry a
 * `*PrChange` record of the properties they replaced.
 */
export type FormatRevisionKind =
  | "runFormat"
  | "paragraphFormat"
  | "tableFormat"
  | "rowFormat"
  | "cellFormat";

/** A tracked-change element the review UI can accept or reject. */
export type RevisionKind =
  | "insertion"
  | "deletion"
  | "markInsertion"
  | "markDeletion"
  | FormatRevisionKind;

export interface RevisionRef {
  /** The w:ins / w:del element (run-level, or the mark ins/del inside rPr),
   * or the w:rPrChange / w:pPrChange / w:tblPrChange / w:trPrChange /
   * w:tcPrChange element for a format revision. */
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

  const newT = textLike(t, t.name, text);
  const ins = el(`${w}ins`, revAttrs(w, meta), [runLike(rEl, rPr, [newT])]);

  // The addressed run element stays in the tree, so its stable id — and any
  // reference a caller still holds — survives the insertion. When the caret
  // is at the run's start the whole run moves after the w:ins untouched;
  // otherwise the run keeps the content before the caret (t truncated in
  // place, the way applySplitParagraph splits a run) and a new run carries
  // the content after.
  const replacement: XmlElement[] = [];
  if (offset === 0 && preceding.length === 0) {
    replacement.push(ins, rEl);
  } else {
    const afterChildren: XmlElement[] = [];
    if (offset < t.text.length) afterChildren.push(textLike(t, t.name, t.text.slice(offset)));
    afterChildren.push(...following);
    t.text = t.text.slice(0, offset);
    t.attrs["xml:space"] = "preserve";
    rEl.children = rEl.children.slice(0, tIdx + 1);
    replacement.push(rEl, ins);
    if (afterChildren.length > 0) replacement.push(runLike(rEl, rPr, afterChildren));
  }

  // A second author can type inside another author's pending insertion. OOXML
  // revisions must stay as siblings, so split the outer insertion around the
  // new one instead of nesting w:ins elements.
  if (localName(parent.name) === "ins") {
    const grand = doc.findParentOf(parent);
    if (!grand) return null;
    const parentIdx = grand.children.indexOf(parent);
    if (parentIdx === -1) return null;

    const insIdx = replacement.indexOf(ins);
    const before = [...parent.children.slice(0, idx), ...replacement.slice(0, insIdx)];
    const after = [...replacement.slice(insIdx + 1), ...parent.children.slice(idx + 1)];
    const outerParts: XmlElement[] = [];
    if (before.length > 0) outerParts.push({ ...cloneXml(parent), children: before });
    outerParts.push(ins);
    if (after.length > 0) {
      const afterOuter = { ...cloneXml(parent), children: after };
      const idKey = Object.keys(afterOuter.attrs).find((key) => localName(key) === "id");
      if (before.length > 0 && idKey) afterOuter.attrs[idKey] = String(meta.nextId());
      outerParts.push(afterOuter);
    }
    grand.children.splice(parentIdx, 1, ...outerParts);
    return { t: newT, offset: text.length };
  }

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

// ---------- formatting revisions (w:rPrChange / w:pPrChange) ----------

/**
 * Tracked FORMATTING changes. Word records them by keeping the new properties
 * where they always live — directly in w:rPr / w:pPr / w:tblPr / w:trPr /
 * w:tcPr, so every renderer shows the new look with no special casing — and
 * stashing the PREVIOUS properties in a `*PrChange` as the last child of that
 * same element. Accept drops the change element; reject puts the recorded
 * properties back.
 *
 * All five are recorded BEFORE the mutation runs, and only once per pending
 * suggestion: a second edit to the same run, paragraph, table, row or cell
 * refines the same suggestion, so the recorded properties stay the ones the
 * reviewer would get back on reject — the state before ANY of it was
 * suggested.
 */

/**
 * Per-kind shape of a tracked formatting revision.
 *
 * `excluded` is what the change element's payload type may NOT carry, which
 * ECMA-376 states as the difference between the properties element's own type
 * and the base type the record holds:
 *
 *  - CT_RPr  minus CT_RPrOriginal  → the change record itself.
 *  - CT_PPr  minus CT_PPrBase      → the paragraph mark's run properties and
 *                                    the section break.
 *  - CT_TblPr minus CT_TblPrBase   → the change record itself.
 *  - CT_TrPr minus CT_TrPrBase     → the row's own w:ins / w:del structural
 *                                    revisions, which are not formatting.
 *  - CT_TcPr minus CT_TcPrInner    → the change record itself.
 *
 * Excluded children stay LIVE: they are neither recorded nor disturbed by a
 * reject, which is what keeps a row's structural revision and a paragraph's
 * sectPr in place while its formatting is reviewed.
 */
const FORMAT_CHANGE: Record<
  FormatRevisionKind,
  { props: string; change: string; excluded: readonly string[] }
> = {
  runFormat: { props: "rPr", change: "rPrChange", excluded: ["rPrChange"] },
  paragraphFormat: { props: "pPr", change: "pPrChange", excluded: ["rPr", "sectPr", "pPrChange"] },
  tableFormat: { props: "tblPr", change: "tblPrChange", excluded: ["tblPrChange"] },
  rowFormat: { props: "trPr", change: "trPrChange", excluded: ["ins", "del", "trPrChange"] },
  cellFormat: { props: "tcPr", change: "tcPrChange", excluded: ["tcPrChange"] },
};

/**
 * Record `owner`'s current properties as the "previous formatting" of a
 * tracked format change. Call BEFORE mutating them.
 *
 * The properties element leads its owner in every one of the five schemas
 * (w:rPr, w:pPr, w:tblPr, w:trPr and w:tcPr are all the first child), so one
 * unshift creates a missing one correctly; the change record closes it, so one
 * push places it correctly.
 */
function recordFormatChange(owner: XmlElement, kind: FormatRevisionKind, meta: RevisionMeta): void {
  const { props: propsName, change: changeName, excluded } = FORMAT_CHANGE[kind];
  const w = prefixOf(owner);
  let props = owner.children.find((c) => localName(c.name) === propsName);
  if (props?.children.some((c) => localName(c.name) === changeName)) return;
  if (!props) {
    props = el(`${w}${propsName}`);
    owner.children.unshift(props);
  }
  const previous = el(
    `${w}${propsName}`,
    {},
    props.children.filter((c) => !excluded.includes(localName(c.name))).map(cloneXml),
  );
  props.children.push(el(`${w}${changeName}`, revAttrs(w, meta), [previous]));
}

/** Record a run's w:rPr as the previous formatting (w:rPrChange). */
export function recordRunFormatChange(rEl: XmlElement, meta: RevisionMeta): void {
  recordFormatChange(rEl, "runFormat", meta);
}

/** Record a paragraph's w:pPr as the previous formatting (w:pPrChange). */
export function recordParagraphFormatChange(pEl: XmlElement, meta: RevisionMeta): void {
  recordFormatChange(pEl, "paragraphFormat", meta);
}

/** Record a table's w:tblPr as the previous formatting (w:tblPrChange). */
export function recordTableFormatChange(tblEl: XmlElement, meta: RevisionMeta): void {
  recordFormatChange(tblEl, "tableFormat", meta);
}

/** Record a row's w:trPr as the previous formatting (w:trPrChange). */
export function recordRowFormatChange(trEl: XmlElement, meta: RevisionMeta): void {
  recordFormatChange(trEl, "rowFormat", meta);
}

/** Record a cell's w:tcPr as the previous formatting (w:tcPrChange). */
export function recordCellFormatChange(tcEl: XmlElement, meta: RevisionMeta): void {
  recordFormatChange(tcEl, "cellFormat", meta);
}

/**
 * Record a table's w:tblGrid as the previous column grid (w:tblGridChange).
 *
 * The grid is NOT part of w:tblPr, so a column-width suggestion that only
 * restored the table properties would put back a width the grid still
 * contradicts. CT_TblGridChange carries a w:id and nothing else — no author,
 * no date — so it is not a revision a reviewer can address on its own; it
 * rides with the table's w:tblPrChange, which every caller here records too,
 * and is accepted or rejected with it.
 */
export function recordTableGridChange(tblEl: XmlElement, meta: RevisionMeta): void {
  const w = prefixOf(tblEl);
  const grid = tblEl.children.find((c) => localName(c.name) === "tblGrid");
  if (!grid || grid.children.some((c) => localName(c.name) === "tblGridChange")) return;
  const previous = el(
    `${w}tblGrid`,
    {},
    grid.children.filter((c) => localName(c.name) !== "tblGridChange").map(cloneXml),
  );
  grid.children.push(el(`${w}tblGridChange`, { [`${w}id`]: String(meta.nextId()) }, [previous]));
}

/** Append a CT_PPrBase property to a w:pPr, ahead of the three children that
 * close the element. A paragraph command that appends has to step over them,
 * or it writes a pPr Word will not read. */
export function appendParagraphProp(pPr: XmlElement, prop: XmlElement): void {
  const trailing = FORMAT_CHANGE.paragraphFormat.excluded;
  const at = pPr.children.findIndex((c) => trailing.includes(localName(c.name)));
  pPr.children.splice(at === -1 ? pPr.children.length : at, 0, prop);
}

// ---------- review: locate / accept / reject ----------

/** Which element owns each format revision, by local name. */
const FORMAT_KIND_BY_OWNER: Record<string, FormatRevisionKind | undefined> = {
  r: "runFormat",
  p: "paragraphFormat",
  tbl: "tableFormat",
  tr: "rowFormat",
  tc: "cellFormat",
};

/** The `*PrChange` a format revision's owner currently holds, if any. */
function formatChangeIn(owner: XmlElement, kind: FormatRevisionKind): XmlElement | null {
  const { props: propsName, change: changeName } = FORMAT_CHANGE[kind];
  const props = owner.children.find((c) => localName(c.name) === propsName);
  return props?.children.find((c) => localName(c.name) === changeName) ?? null;
}

function isFormatKind(kind: RevisionKind): kind is FormatRevisionKind {
  return kind in FORMAT_CHANGE;
}

/**
 * The revision at a w:t, for the accept/reject popover. Resolved innermost
 * out, and in this order of preference:
 *
 *  1. A w:ins / w:del the text sits inside — the thing the caret is literally
 *     within, and what the review UI has always offered. It cannot cross a
 *     paragraph, so that search stops at the w:p.
 *  2. The paragraph's own mark revision, an inserted or deleted pilcrow.
 *  3. A FORMAT revision on an ancestor: the run's w:rPrChange, then the
 *     paragraph's, cell's, row's and table's records. Without this pass a
 *     tracked formatting change is countable, and acceptable in bulk, but
 *     unreachable by clicking the text it formats.
 *
 * Structural revisions outrank formatting ones because they decide whether
 * content exists at all; a paragraph can carry both at once.
 */
export function revisionForText(doc: DocxDocument, t: XmlElement): RevisionRef | null {
  let format: RevisionRef | null = null;
  let insideParagraph = true;
  let cur: XmlElement | undefined = doc.findParentOf(t);
  while (cur) {
    const ln = localName(cur.name);
    if (insideParagraph && (ln === "ins" || ln === "del")) {
      return {
        el: cur,
        kind: ln === "ins" ? "insertion" : "deletion",
        author: attrByLocal(cur, "author"),
      };
    }
    if (ln === "p") {
      const insMark = paragraphGlyphRevision(cur, "ins");
      if (insMark) return { el: insMark, kind: "markInsertion", paragraph: cur };
      const delMark = paragraphGlyphRevision(cur, "del");
      if (delMark) return { el: delMark, kind: "markDeletion", paragraph: cur };
      insideParagraph = false;
    }
    const kind = FORMAT_KIND_BY_OWNER[ln];
    if (kind && !format) {
      const change = formatChangeIn(cur, kind);
      if (change) {
        format = {
          el: change,
          kind,
          ...(kind === "paragraphFormat" ? { paragraph: cur } : {}),
          author: attrByLocal(change, "author"),
        };
      }
    }
    if (ln === "body") break;
    cur = doc.findParentOf(cur);
  }
  return format;
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
 * Every tracked change in the document, in document order: run-level
 * w:ins/w:del wrappers, paragraph-mark revisions (pPr/rPr ins/del), and the
 * five formatting records. Covers the body, headers/footers and footnotes —
 * everything the editor can suggest into.
 */
export function collectRevisions(doc: DocxDocument): RevisionRef[] {
  const out: RevisionRef[] = [];
  const walk = (el: XmlElement): void => {
    const ln = localName(el.name);
    // Paragraph-mark revisions live in pPr/rPr; collect them at the w:p so
    // they carry their owning paragraph, and never double-collect them as
    // run-level wrappers.
    if (ln === "pPr") return;
    if (ln === "ins" || ln === "del") {
      out.push({
        el,
        kind: ln === "ins" ? "insertion" : "deletion",
        author: attrByLocal(el, "author"),
      });
    }
    // Run, table, row and cell formatting revisions are collected AT their
    // owner, ahead of its content, because properties precede what they
    // format. The paragraph is the exception and is skipped here: the pilcrow
    // its pPrChange formats ENDS the paragraph, so document order puts it
    // after the paragraph's runs, and it is collected below.
    const formatKind = ln === "p" ? undefined : FORMAT_KIND_BY_OWNER[ln];
    if (formatKind) {
      const change = formatChangeIn(el, formatKind);
      if (change) out.push({ el: change, kind: formatKind, author: attrByLocal(change, "author") });
    }
    // A row's w:trPr may hold its own structural w:ins / w:del. Those are not
    // run-level revisions, and unwrapping one would splice its children into
    // the trPr, so never descend into a trPr — its trPrChange is already
    // collected.
    if (ln === "trPr") return;
    for (const c of el.children) walk(c);
    // The paragraph-mark revision is the pilcrow at the paragraph's END, so
    // it follows the paragraph's own runs in document order.
    if (ln === "p") {
      const change = formatChangeIn(el, "paragraphFormat");
      if (change) {
        out.push({ el: change, kind: "paragraphFormat", paragraph: el, author: attrByLocal(change, "author") });
      }
      const pPr = el.children.find((c) => localName(c.name) === "pPr");
      const rPr = pPr?.children.find((c) => localName(c.name) === "rPr");
      for (const c of rPr?.children ?? []) {
        const cn = localName(c.name);
        if (cn === "ins" || cn === "del") {
          out.push({
            el: c,
            kind: cn === "ins" ? "markInsertion" : "markDeletion",
            paragraph: el,
            author: attrByLocal(c, "author"),
          });
        }
      }
    }
  };
  for (const root of doc.revisionRoots()) walk(root);
  return out;
}

function applyAccept(doc: DocxDocument, ref: RevisionRef): boolean {
  switch (ref.kind) {
    case "insertion":
      return unwrap(doc, ref.el);
    case "deletion":
      return removeEl(doc, ref.el);
    case "markInsertion":
      return removeMarkRevision(ref);
    case "markDeletion":
      if (!ref.paragraph) return false;
      removeMarkRevision(ref);
      return joinWithNext(doc, ref.paragraph);
    case "runFormat":
    case "paragraphFormat":
    case "tableFormat":
    case "rowFormat":
    case "cellFormat":
      // The new formatting is already the live one; accepting just retires the
      // record of what it replaced.
      return removeFormatChange(doc, ref);
  }
}

/**
 * The w:tblGrid change record riding with a table's w:tblPrChange, if any.
 * `props` is the w:tblPr holding the change; its parent is the table.
 */
function gridChangeBeside(doc: DocxDocument, props: XmlElement): XmlElement | null {
  const tbl = doc.findParentOf(props);
  const grid = tbl?.children.find((c) => localName(c.name) === "tblGrid");
  return grid?.children.find((c) => localName(c.name) === "tblGridChange") ?? null;
}

/** Drop a `*PrChange` from the properties element holding it. */
function removeFormatChange(doc: DocxDocument, ref: RevisionRef): boolean {
  const props = doc.findParentOf(ref.el);
  if (!props) return false;
  const idx = props.children.indexOf(ref.el);
  if (idx === -1) return false;
  props.children.splice(idx, 1);
  if (ref.kind === "tableFormat") {
    const gridChange = gridChangeBeside(doc, props);
    if (gridChange) removeEl(doc, gridChange);
  }
  if (props.children.length === 0) removeEl(doc, props);
  return true;
}

/**
 * Reject a formatting revision: the recorded previous properties become the
 * live ones again, and the children the record could not carry (see
 * FORMAT_CHANGE.excluded) stay exactly as they are, keeping their schema
 * position at the end. That is what leaves a paragraph's w:sectPr, a
 * paragraph mark's w:rPr and a row's structural w:ins / w:del untouched while
 * their formatting is reviewed.
 */
function restoreFormatChange(doc: DocxDocument, ref: RevisionRef): boolean {
  if (!isFormatKind(ref.kind)) return false;
  const props = doc.findParentOf(ref.el);
  if (!props || props.children.indexOf(ref.el) === -1) return false;
  const { change, excluded } = FORMAT_CHANGE[ref.kind];
  const recorded = ref.el.children[0];
  const restored = (recorded?.children ?? []).map(cloneXml);
  const kept = props.children.filter((c) => {
    const ln = localName(c.name);
    if (ln === change || !excluded.includes(ln)) return false;
    // An emptied paragraph-mark rPr — its own revision already reviewed — is
    // not a property, and keeping it would leave the paragraph in a shape it
    // was never in.
    if (ln === "rPr") return c.children.length > 0;
    return true;
  });
  if (ref.kind === "tableFormat") restoreGridChange(doc, props);
  props.children = [...restored, ...kept];
  if (props.children.length === 0) removeEl(doc, props);
  return true;
}

/** Put a table's recorded w:tblGrid back and drop the record. */
function restoreGridChange(doc: DocxDocument, props: XmlElement): void {
  const gridChange = gridChangeBeside(doc, props);
  if (!gridChange) return;
  const grid = doc.findParentOf(gridChange);
  if (!grid) return;
  grid.children = (gridChange.children[0]?.children ?? []).map(cloneXml);
}

function applyReject(doc: DocxDocument, ref: RevisionRef): boolean {
  switch (ref.kind) {
    case "insertion":
      return removeEl(doc, ref.el);
    case "deletion":
      convertDelTextToText(ref.el);
      return unwrap(doc, ref.el);
    case "markInsertion":
      if (!ref.paragraph) return false;
      removeMarkRevision(ref);
      return joinWithNext(doc, ref.paragraph);
    case "markDeletion":
      return removeMarkRevision(ref);
    case "runFormat":
    case "paragraphFormat":
    case "tableFormat":
    case "rowFormat":
    case "cellFormat":
      return restoreFormatChange(doc, ref);
  }
}

/**
 * Accept a single revision: an insertion becomes permanent text, a deletion
 * removes its content, an inserted mark stays split, a deleted mark merges the
 * two paragraphs, and a format change keeps the new formatting. Returns false
 * if the element could not be located.
 */
export function acceptRevision(doc: DocxDocument, ref: RevisionRef): boolean {
  const parts = doc.notePartsHolding(ref.el);
  const ok = applyAccept(doc, ref);
  if (ok) {
    doc.markNotePartsChanged(parts);
    doc.refresh();
  }
  return ok;
}

/**
 * Reject a single revision: an insertion disappears, a deletion's content is
 * restored, an inserted mark un-splits (the paragraphs rejoin), a deleted mark
 * is removed (the paragraphs stay separate), and a format change puts the
 * previous properties back. Returns false if not located.
 */
export function rejectRevision(doc: DocxDocument, ref: RevisionRef): boolean {
  const parts = doc.notePartsHolding(ref.el);
  const ok = applyReject(doc, ref);
  if (ok) {
    doc.markNotePartsChanged(parts);
    doc.refresh();
  }
  return ok;
}

/**
 * Accept (or reject) every tracked change in one pass, refreshing the model
 * once at the end. Refs are applied in REVERSE document order: a paragraph-
 * mark accept/reject joins its paragraph with the NEXT one, so processing
 * back-to-front keeps every earlier ref's elements attached to the tree
 * (front-to-back, the second of two consecutive mark revisions would already
 * be absorbed and silently skipped). Returns how many revisions were applied.
 */
export function acceptAllRevisions(doc: DocxDocument): number {
  return applyToAll(doc, applyAccept);
}

export function rejectAllRevisions(doc: DocxDocument): number {
  return applyToAll(doc, applyReject);
}

function applyToAll(doc: DocxDocument, apply: (doc: DocxDocument, ref: RevisionRef) => boolean): number {
  const refs = collectRevisions(doc);
  const notes = new Set<NotePart>();
  let n = 0;
  for (let i = refs.length - 1; i >= 0; i--) {
    // Ask which part holds the revision while it is still attached; keep the
    // answer only when the mutation lands, so a part nothing changed in is
    // never re-serialized.
    const held = doc.notePartsHolding(refs[i].el);
    if (!apply(doc, refs[i])) continue;
    n++;
    for (const part of held) notes.add(part);
  }
  if (n > 0) {
    doc.markNotePartsChanged([...notes]);
    doc.refresh();
  }
  return n;
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
