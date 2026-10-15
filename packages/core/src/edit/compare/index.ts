import { diffArrays } from "diff";
import { DocxDocument } from "../../docx.js";
import { XmlElement, child, cloneXml, localName, sameXml, serializeXml } from "../../xml.js";
import {
  RevisionMeta,
  acceptAllRevisions,
  collectRevisions,
  markDeletedText,
  markParagraphGlyph,
  markRowRevision,
  recordCellFormatChange,
  recordParagraphFormatChange,
  recordRowFormatChange,
  recordRunFormatChange,
  recordSectionFormatChange,
  recordTableFormatChange,
  recordTableGridChange,
  revisionWrapper,
} from "../suggest.js";
import { CONTAINERS, ContentToken, Placed, ReadOptions, isPassenger, paragraphText, readParagraph, runPropsKey, sliceItems, textBoxesIn } from "./content.js";
import { Key, histogramMatch } from "./histogram.js";
import { COMPARE_TUNING, normalizeText, similarity, similarityFloor } from "./text.js";

/**
 * Compare Documents — Word's "legal blackline".
 *
 * Produces a THIRD document: `revised`, with every difference from `original`
 * expressed as an ordinary tracked change attributed to one reviewer. Nothing
 * here invents a revision format; it drives the same w:ins / w:del / *PrChange
 * vocabulary the suggesting editor writes, so the result is reviewable with
 * the accept/reject surface that already exists, and round-trips through it:
 *
 *   accept every revision in the result  →  `revised`
 *   reject every revision in the result  →  `original`
 *
 * (Equal modulo run coalescing — accepting a revision unwraps it without
 * re-merging the runs it split, which is true of Word too. The gate test in
 * test/compare-roundtrip.test.ts states the canonical form exactly.)
 *
 * ## Tiers, stated plainly
 *
 * **Text — full.** Paragraphs align by histogram diff over a text/style/
 * numbering fingerprint, then by Dice similarity inside the gaps; matched
 * paragraphs are diffed word by word. Insertions, deletions, paragraph splits
 * and merges, paragraph-property and run-property changes are all expressed
 * faithfully.
 *
 * **Tables — rows and cells.** A matched table pair aligns its ROWS by the
 * same histogram pass, so a mid-table insert does not cascade. Matched rows
 * with the same cell count recurse into each cell's paragraphs. A row that
 * matches nothing becomes a row insertion or deletion (w:trPr/w:ins, w:trPr/
 * w:del). A row pair whose cell COUNTS differ is coarse: the whole row is
 * struck and the whole new row inserted.
 *
 * **Drawings, fields, footnote references and other non-text run content —
 * atomic.** They match themselves byte for byte or they are a change. A
 * changed image is a struck image plus an inserted one, never a patched one.
 *
 * **Paragraphs holding a hyperlink, content control, smart tag or inline
 * math — coarse.** Word-level diffing rebuilds a paragraph's runs, and a
 * revision wrapper may not sit around a `w:hyperlink` (the schema puts w:ins
 * INSIDE it). Such a paragraph is never word-diffed: if it changed at all it
 * is struck whole and the new one inserted whole, and `onNote` says so.
 *
 * **Headers, footers, footnotes and endnotes — compared as the body is.** A
 * header or footer the last section has on one side alone is compared against
 * nothing, so it comes out wholly inserted or struck. No revision can carry
 * the section's reference to it, so `onNote` reports that change.
 *
 * **Styles and numbering — NOT compared.** The result carries `revised`'s.
 * `onNote` reports when `original`'s differ, so a reviewer is never told "no
 * change" about something that was never looked at. The result also carries
 * `revised`'s comments part as it is.
 *
 * **Moves — not detected, by architecture.** `RevisionKind` has no moveFrom /
 * moveTo member and nothing in suggest.ts writes those elements, so a move is
 * a new output primitive, not a smarter aligner. A moved paragraph comes out
 * as a deletion plus an insertion, which is what Word usually produces too.
 */

export interface CompareNote {
  code:
    /** A paragraph pair too structured to word-diff; struck and reinserted whole. */
    | "coarse-paragraph"
    /** A table row pair whose cell counts differ; struck and reinserted whole. */
    | "coarse-row"
    /** A block that is neither a paragraph nor a table, added or removed. */
    | "untracked-block"
    /** A paragraph mark whose run properties differ across the swap; see below. */
    | "paragraph-mark-props"
    /** A part outside the body differs and was not compared. */
    | "part-not-compared"
    /** The documents share no paragraph at all; one paragraph mark goes untracked. */
    | "no-common-paragraph"
    /** Original-only content referencing a relationship the result does not have. */
    | "unresolved-relationship"
    /** The last section gained or lost a header or footer; the reference is not a revision, so one review leaves it empty. */
    | "header-footer-reference"
    /** An input held tracked changes; it was compared as if they were accepted. */
    | "existing-revisions-accepted";
  detail: string;
}

export interface CompareOptions {
  /** Who the revisions are attributed to. */
  author: string;
  /** ISO 8601 timestamp for every w:date written. Default: the current time. */
  date?: string;
  /**
   * Report formatting-only differences as tracked format revisions. Default
   * true; Word exposes the same switch.
   */
  formatting?: boolean;
  /**
   * Report whitespace-only differences. Default true; Word exposes the same
   * switch. Turning it OFF is lossy on purpose — a difference that is not
   * recorded cannot be rejected back either, so the accept/reject round trip
   * does not hold for whitespace with this off.
   */
  whitespace?: boolean;
  /** Called once per difference that could not be expressed faithfully. */
  onNote?: (note: CompareNote) => void;
  /**
   * Write the result into `revised` instead of a copy of it, and return
   * `revised`. A copy is a whole second document in memory (114 MB for a
   * 217-page book), so a caller that does not need `revised` afterwards
   * should pass this. `original` is never changed either way.
   */
  reuseRevised?: boolean;
}

export interface Ctx {
  doc: DocxDocument;
  /** The document deleted content is cloned OUT of; see adoptFromOriginal. */
  original: DocxDocument;
  meta: RevisionMeta;
  formatting: boolean;
  whitespace: boolean;
  note: (note: CompareNote) => void;
  /**
   * The parts whose relationships struck content resolves through: the
   * header or footer part on each side, or null for the document part.
   */
  scope: { from: string | null; into: string | null };
  /** Original note id -> the result's id for the same note (compareNotes). */
  noteIds: Record<"footnote" | "endnote", Map<string, string>>;
  /** Original comment id -> the result's id for that comment (adoptComment). */
  commentIds: Map<string, string>;
}

/** Where a block ended up, which is what decides the paragraph-mark revisions. */
export type BlockClass = "kept" | "inserted" | "deleted";

export interface Entry {
  el: XmlElement;
  cls: BlockClass;
}

/**
 * Compare two documents. Returns a NEW document; neither input is modified.
 */
export function compareDocuments(
  original: DocxDocument,
  revised: DocxDocument,
  options: CompareOptions,
): DocxDocument {
  // Word compares documents "as if the tracked changes have been accepted",
  // and so does this: carried through, an earlier review's revisions mixed
  // with this comparison's, and reject-all undid both.
  const note = options.onNote ?? (() => {});
  const accepted = (doc: DocxDocument, which: string, inPlace: boolean): DocxDocument => {
    if (collectRevisions(doc).length === 0) return doc;
    const copy = inPlace ? doc : doc.clone();
    acceptAllRevisions(copy);
    note({ code: "existing-revisions-accepted", detail: `the ${which} document held tracked changes; it was compared as if they were accepted` });
    return copy;
  };
  const reuse = options.reuseRevised === true;
  const merged = accepted(revised, "revised", reuse);
  // Unless the caller gave it up, the result is a copy: it must never share
  // the revised input's tree.
  return compareInto(accepted(original, "original", false), merged === revised && !reuse ? revised.clone() : merged, options);
}

function compareInto(original: DocxDocument, merged: DocxDocument, options: CompareOptions): DocxDocument {
  // Every revision takes an id; with none issued the body is the clone's own,
  // untouched, and its model needs no rebuild (refresh was a sixth of the time
  // on a document compared with itself).
  let issued = 0;
  const ctx = compareContext(original, merged, options, () => issued++);

  const originalBody = child(original.docRoot, "body");
  const mergedBody = child(merged.docRoot, "body");
  if (!originalBody || !mergedBody) throw new Error("document.xml has no w:body");

  // Before the body is touched, merged holds exactly the revised parts.
  reportUncomparedParts(original, merged, ctx);
  const headerFooterPairs = pairHeadersFooters(original, merged);
  // Notes first: struck body text carries the ORIGINAL's note references,
  // and adoptFromOriginal re-points them through the ids this settles.
  compareNotes("footnote", original, merged, ctx);
  compareNotes("endnote", original, merged, ctx);

  const tail = mergedBody.children.filter((c) => localName(c.name) === "sectPr");
  const entries = alignBlocks(blockList(originalBody, "sectPr"), blockList(mergedBody, "sectPr"), ctx);
  applyParagraphMarks(entries, ctx);
  mergedBody.children = [...entries.map((e) => e.el), ...tail];
  giveStruckReferencesTheirOwnNotes("footnote", merged, ctx);
  giveStruckReferencesTheirOwnNotes("endnote", merged, ctx);

  compareHeadersFooters(original, merged, headerFooterPairs, ctx);
  compareLastSection(original, merged, ctx);

  if (issued > 0) merged.refresh();
  return merged;
}

/**
 * What one compare carries through every step. Revisions are written into
 * `merged`; `onIssue` is called for each revision id taken.
 */
export function compareContext(original: DocxDocument, merged: DocxDocument, options: CompareOptions, onIssue: () => void = () => {}): Ctx {
  return {
    doc: merged,
    original,
    meta: {
      author: options.author,
      date: options.date ?? new Date().toISOString().replace(/\.\d{3}Z$/, "Z"),
      nextId: () => {
        onIssue();
        return merged.nextRevisionId();
      },
    },
    formatting: options.formatting !== false,
    whitespace: options.whitespace !== false,
    note: options.onNote ?? (() => {}),
    scope: { from: null, into: null },
    noteIds: { footnote: new Map(), endnote: new Map() },
    commentIds: new Map(),
  };
}

/** Align and diff one container's blocks in place (a header, a note). */
export function compareContainer(oRoot: XmlElement, mRoot: XmlElement, ctx: Ctx): void {
  const o = blockList(oRoot);
  const m = blockList(mRoot);
  if (o.length === m.length && o.every((el, i) => sameXml(el, m[i]))) return;
  const entries = alignBlocks(o, m, ctx);
  applyParagraphMarks(entries, ctx);
  // Keep the container's own non-block children (a note's properties) ahead.
  const blocks = new Set(m);
  mRoot.children = [...mRoot.children.filter((c) => !blocks.has(c) && !["p", "tbl", "sdt"].includes(localName(c.name))), ...entries.map((e) => e.el)];
}

/**
 * A header or footer part on each side of the compare, or the last
 * section's on one side alone when the other side's last section shows
 * none of that kind and type.
 */
export interface HeaderFooterPair {
  kind: "header" | "footer";
  type: "default" | "first" | "even";
  from: string | null;
  into: string | null;
}

/**
 * Header and footer parts paired across the documents: section by section,
 * kind by kind (default, first, even), each pair once. A section without a
 * reference of a kind shows the previous section's part, as Word does, so
 * it pairs that part. `sections` gives each side's section breaks that are
 * not in its tree, in order, ahead of those that are.
 *
 * The last section's references with no counterpart come out one-sided. An
 * earlier section's do not: a section-break paragraph whose references
 * differ is struck and reinserted whole (pairable), which carries them.
 */
export function pairHeadersFooters(
  original: DocxDocument,
  merged: DocxDocument,
  sections?: { original: XmlElement[]; merged: XmlElement[] },
): HeaderFooterPair[] {
  const refs = (doc: DocxDocument): Map<string, string> => {
    const out = new Map<string, string>();
    const body = child(doc.docRoot, "body");
    const sects = [
      ...((doc === original ? sections?.original : sections?.merged) ?? []),
      ...(body?.children ?? []).flatMap((c) =>
        localName(c.name) === "sectPr" ? [c] : localName(c.name) === "p" ? [child(child(c, "pPr"), "sectPr")].filter((x): x is XmlElement => !!x) : [],
      ),
    ];
    const shown = new Map<string, string>();
    sects.forEach((sp, i) => {
      for (const ref of sp.children) {
        const ln = localName(ref.name);
        if (ln !== "headerReference" && ln !== "footerReference") continue;
        const target = doc.documentRels.get(ref.attrs["r:id"])?.target;
        if (target) shown.set(`${ln}:${ref.attrs["w:type"] ?? "default"}`, target);
      }
      // The body's own section properties are the last section's.
      const section = body?.children.includes(sp) ? "last" : String(i);
      for (const [kindType, target] of shown) out.set(`${section}:${kindType}`, target);
    });
    return out;
  };
  const theirs = refs(original);
  const mine = refs(merged);
  const pairs: HeaderFooterPair[] = [];
  const paired = { from: new Set<string>(), into: new Set<string>() };
  const parse = (key: string): Pick<HeaderFooterPair, "kind" | "type"> => {
    const [, ln, type] = key.split(":");
    return { kind: ln === "headerReference" ? "header" : "footer", type: type as HeaderFooterPair["type"] };
  };
  const pair = (key: string, from: string, into: string): void => {
    paired.from.add(from);
    paired.into.add(into);
    pairs.push({ ...parse(key), from, into });
  };
  const same = (from: string, into: string): boolean => {
    const oRoot = original.headerFooterRoot(from);
    const mRoot = merged.headerFooterRoot(into);
    return !!oRoot && !!mRoot && sameXml(oRoot, mRoot);
  };
  // A part with the same content on both sides is the same part wherever
  // its section now sits (a removed section shifts the ones after it): at
  // its own section first, then anywhere, in section order. The rest pair
  // section by section.
  for (const [key, from] of theirs) {
    const into = mine.get(key);
    if (into && !paired.into.has(into) && same(from, into)) pair(key, from, into);
  }
  for (const [key, from] of theirs) {
    if (paired.from.has(from)) continue;
    const into = [...mine.values()].find((target) => !paired.into.has(target) && same(from, target));
    if (into) pair(key, from, into);
  }
  for (const [key, from] of theirs) {
    const into = mine.get(key);
    if (into && !paired.into.has(into)) pair(key, from, into);
    else if (!into && key.startsWith("last:")) pairs.push({ ...parse(key), from, into: null });
  }
  for (const [key, into] of mine) {
    if (key.startsWith("last:") && !theirs.has(key) && !paired.into.has(into)) pairs.push({ ...parse(key), from: null, into });
  }
  return pairs;
}

/**
 * The last section's page setup (size, orientation, margins, columns and the
 * rest of w:sectPr), tracked as a w:sectPrChange holding the original's when
 * it differs. The header and footer references are compareHeadersFooters'.
 */
export function compareLastSection(original: DocxDocument, merged: DocxDocument, ctx: Ctx): void {
  if (!ctx.formatting) return;
  const before = child(child(original.docRoot, "body"), "sectPr");
  const after = child(child(merged.docRoot, "body"), "sectPr");
  if (!after) return;
  const contents = (sectPr: XmlElement | undefined): string =>
    (sectPr?.children ?? []).filter((c) => !["headerReference", "footerReference", "sectPrChange"].includes(localName(c.name))).map((c) => serializeXml(c)).join("");
  if (contents(before) === contents(after)) return;
  recordSectionFormatChange(after, before ?? { name: after.name, attrs: {}, children: [], text: "" }, ctx.meta);
}

/**
 * Compare each paired header or footer part in place. A part the last
 * section has on one side alone is compared against nothing: one the
 * revised document lost gets a new part in the result, referenced from the
 * last section, with the original's content struck; one it gained has its
 * content marked inserted. A sectPrChange cannot carry the reference
 * itself, so the review that removes the content leaves an empty part, and
 * `onNote` says so.
 */
export function compareHeadersFooters(original: DocxDocument, merged: DocxDocument, pairs: HeaderFooterPair[], ctx: Ctx): void {
  for (const { kind, type, from, into } of pairs) {
    const oRoot = from === null ? null : original.headerFooterRoot(from);
    let mRoot = into === null ? null : merged.headerFooterRoot(into);
    if (from !== null && into !== null) {
      if (oRoot && mRoot) compareContainer(oRoot, mRoot, { ...ctx, scope: { from, into } });
      continue;
    }
    const name = `the last section's ${type === "default" ? "" : `${type}-page `}${kind}`;
    if (from !== null) {
      const sectPr = child(child(merged.docRoot, "body"), "sectPr");
      if (!oRoot || !sectPr) continue;
      const target = merged.addHfPart(kind, type, sectPr);
      mRoot = merged.headerFooterRoot(target) as XmlElement;
      mRoot.children = [];
      compareContainer(oRoot, mRoot, { ...ctx, scope: { from, into: target } });
      ctx.note({ code: "header-footer-reference", detail: `${name} was removed; its reference is not a revision, so accept-all leaves the ${kind} empty` });
    } else if (mRoot) {
      compareContainer({ ...mRoot, children: [] }, mRoot, { ...ctx, scope: { from: null, into } });
      ctx.note({ code: "header-footer-reference", detail: `${name} was added; its reference is not a revision, so reject-all leaves the ${kind} empty` });
    }
  }
}

/**
 * A footnotes or endnotes part, as compareNotes and
 * copyNotesForStruckReferences read and change it. compareDocuments reads it
 * from the document's tree (treeNotes); a caller holding the notes elsewhere
 * can answer from there.
 */
export interface NotesPart {
  /** The note with this id that holds content (no w:type): the last one, when ids repeat. */
  content(id: string): XmlElement | undefined;
  /** What two equal notes share: normalizeText of the content note's paragraphText. */
  key(id: string): Key;
  /** The part's first note with this id, of any type. */
  first(id: string): XmlElement | undefined;
  /** Keep the changes made to a content note. A tree's note is changed in place. */
  update(id: string, note: XmlElement): void;
  /** Add a note at the end of the part. */
  append(note: XmlElement): void;
  /** One more than the highest id of any child of the part; at least 1. */
  nextId(): number;
}

/** The notes part in a document's tree, or null when it has none (and `create` is false). */
function treeNotes(kind: "footnote" | "endnote", doc: DocxDocument, create = false): NotesPart | null {
  const root = kind === "footnote" ? doc.footnotesTree(create) : doc.endnotesTree(create);
  if (!root) return null;
  const byId = new Map(
    root.children.filter((c) => localName(c.name) === kind && !c.attrs["w:type"]).map((c) => [c.attrs["w:id"], c] as const),
  );
  return {
    content: (id) => byId.get(id),
    key: (id) => normalizeText(paragraphText(byId.get(id) as XmlElement)),
    first: (id) => root.children.find((c) => localName(c.name) === kind && c.attrs["w:id"] === id),
    update: () => {},
    append: (note) => root.children.push(note),
    nextId: () => {
      let next = 1;
      for (const c of root.children) next = Math.max(next, 1 + (Number(c.attrs["w:id"]) || 0));
      return next;
    },
  };
}

/**
 * Footnotes or endnotes, paired in the order their references appear and
 * compared inside. A note only the revised document has gets its content
 * marked inserted; one only the original had is copied into the result
 * under a fresh id with its content struck, and struck body text pointing
 * at it is re-pointed (adoptFromOriginal, through ctx.noteIds). `refs` gives
 * each side's reference ids that are not in its tree, in order, ahead of
 * those that are; `parts` gives each side's notes part when it is not in
 * the tree.
 */
export function compareNotes(
  kind: "footnote" | "endnote",
  original: DocxDocument,
  merged: DocxDocument,
  ctx: Ctx,
  refs?: { original: string[]; merged: string[] },
  parts?: { original: NotesPart | null; merged: NotesPart | null },
): void {
  const oPart = parts ? parts.original : treeNotes(kind, original);
  let mPart = parts ? parts.merged : treeNotes(kind, merged);
  if (!oPart && !mPart) return;
  const refName = `${kind}Reference`;
  const inOrder = (doc: DocxDocument, part: NotesPart | null): string[] => {
    if (!part) return [];
    const out: string[] = [];
    const seen = new Set<string>();
    const take = (id: string): void => {
      if (part.content(id) && !seen.has(id)) {
        seen.add(id);
        out.push(id);
      }
    };
    const walk = (e: XmlElement): void => {
      if (localName(e.name) === refName) take(e.attrs["w:id"]);
      for (const c of e.children) walk(c);
    };
    for (const id of (doc === original ? refs?.original : refs?.merged) ?? []) take(id);
    walk(doc.docRoot);
    return out;
  };
  // Each side's content notes, by id, in reference order.
  const oNotes = inOrder(original, oPart);
  const rNotes = inOrder(merged, mPart);
  const oKey = (i: number): Key => (oPart as NotesPart).key(oNotes[i]);
  const rKey = (j: number): Key => (mPart as NotesPart).key(rNotes[j]);
  // A note with the same text on both sides is the same note wherever its
  // reference now sits: a note's place is its reference's, so moving the
  // paragraph that holds it must not unpair it. Equal texts pair first, in
  // reference order.
  const pairs: [number, number][] = [];
  const waiting = new Map<Key, number[]>();
  rNotes.forEach((_, j) => {
    const list = waiting.get(rKey(j));
    if (list) list.push(j);
    else waiting.set(rKey(j), [j]);
  });
  oNotes.forEach((_, i) => {
    const j = waiting.get(oKey(i))?.shift();
    if (j !== undefined) pairs.push([i, j]);
  });
  // The rest anchor in reference order, and a single leftover on each side of
  // a slot is that note, edited (the same rule as paragraphs).
  const exactO = new Set(pairs.map((p) => p[0]));
  const exactR = new Set(pairs.map((p) => p[1]));
  const oRest = oNotes.map((_, i) => i).filter((i) => !exactO.has(i));
  const rRest = rNotes.map((_, j) => j).filter((j) => !exactR.has(j));
  let oAt = 0;
  let rAt = 0;
  const rest = histogramMatch(oRest.map(oKey), rRest.map(rKey));
  for (const a of [...rest, { a: oRest.length, b: rRest.length }]) {
    if (a.a - oAt === 1 && a.b - rAt === 1) pairs.push([oRest[oAt], rRest[rAt]]);
    if (a.a < oRest.length && a.b < rRest.length) pairs.push([oRest[a.a], rRest[a.b]]);
    oAt = a.a + 1;
    rAt = a.b + 1;
  }
  const ids = ctx.noteIds[kind];
  const pairedO = new Set(pairs.map((p) => p[0]));
  const pairedR = new Set(pairs.map((p) => p[1]));
  let changed = false;
  for (const [i, j] of pairs) {
    ids.set(oNotes[i], rNotes[j]);
    const before = ctx.meta.nextId;
    let touched = false;
    const note = (mPart as NotesPart).content(rNotes[j]) as XmlElement;
    compareContainer((oPart as NotesPart).content(oNotes[i]) as XmlElement, note, { ...ctx, meta: { ...ctx.meta, nextId: () => ((touched = true), before()) } });
    (mPart as NotesPart).update(rNotes[j], note);
    changed ||= touched;
  }
  rNotes.forEach((id, j) => {
    if (pairedR.has(j)) return;
    const note = (mPart as NotesPart).content(id) as XmlElement;
    for (const block of blockList(note)) if (localName(block.name) === "p") wrapContainerContent(block, "ins", ctx);
    (mPart as NotesPart).update(id, note);
    changed = true;
  });
  const removed = oNotes.filter((_, i) => !pairedO.has(i));
  if (removed.length > 0) {
    mPart ??= treeNotes(kind, merged, true);
    if (mPart) {
      let next = mPart.nextId();
      for (const id of removed) {
        const copy = cloneXml((oPart as NotesPart).content(id) as XmlElement);
        copy.attrs["w:id"] = String(next);
        ids.set(id, String(next++));
        for (const block of blockList(copy)) {
          if (localName(block.name) !== "p") continue;
          adoptFromOriginal(block, { ...ctx, scope: { from: null, into: null } });
          wrapContainerContent(block, "del", ctx);
        }
        mPart.append(copy);
      }
      changed = true;
    }
  }
  if (changed) merged.markNotePartsChanged([kind === "footnote" ? "footnotes" : "endnotes"]);
}

/**
 * One note, one reference. A paragraph that moved is struck at one place and
 * inserted at another, and both copies referenced the same note, which Word
 * refuses to open (parity2-notes with blocks moved). A struck reference gets
 * a struck copy of its note instead, as Word gives one when it tracks the
 * deletion of a reference.
 */
function giveStruckReferencesTheirOwnNotes(kind: "footnote" | "endnote", merged: DocxDocument, ctx: Ctx): void {
  const refs: NoteRef[] = [];
  noteRefsIn(merged.docRoot, kind, false, refs);
  for (const [i, id] of copyNotesForStruckReferences(kind, merged, ctx, refs)) refs[i].el.attrs["w:id"] = id;
}

/** A note reference, and whether a revision strikes it. */
export interface NoteRef {
  el: XmlElement;
  struck: boolean;
}

/** The `kind` note references inside `el`, in document order, onto `out`. */
export function noteRefsIn(el: XmlElement, kind: "footnote" | "endnote", struck: boolean, out: NoteRef[]): void {
  if (localName(el.name) === `${kind}Reference`) out.push({ el, struck });
  // A deleted row records it in its trPr, and its cells (a nested table's
  // too) may hold references with no w:del around them.
  const rowDeleted = (tr: XmlElement): boolean =>
    localName(tr.name) === "tr" && !!child(child(tr, "trPr"), "del");
  for (const c of el.children) noteRefsIn(c, kind, struck || localName(c.name) === "del" || rowDeleted(c), out);
}

/**
 * The note side of giveStruckReferencesTheirOwnNotes. `refs` are every
 * `kind` reference in the body, in order. Each struck reference beyond the
 * one that keeps its note gets a struck copy of the note; the result maps
 * such a reference's index in `refs` to its copy's id.
 */
export function copyNotesForStruckReferences(
  kind: "footnote" | "endnote",
  merged: DocxDocument,
  ctx: Ctx,
  refs: readonly { el: { attrs: Record<string, string> }; struck: boolean }[],
  part: NotesPart | null = treeNotes(kind, merged),
): Map<number, string> {
  const ids = new Map<number, string>();
  if (!part) return ids;
  const byId = new Map<string, number[]>();
  refs.forEach((ref, i) => {
    const list = byId.get(ref.el.attrs["w:id"]);
    if (list) list.push(i);
    else byId.set(ref.el.attrs["w:id"], [i]);
  });
  let next = part.nextId();
  let changed = false;
  for (const [id, list] of byId) {
    if (list.length < 2) continue;
    const note = part.first(id);
    if (!note) continue;
    // The live reference keeps the note; every struck one beyond it gets a copy.
    const keep = list.find((i) => !refs[i].struck) ?? list[0];
    for (const i of list) {
      if (i === keep || !refs[i].struck) continue;
      const copy = cloneXml(note);
      copy.attrs["w:id"] = String(next);
      ids.set(i, String(next++));
      for (const block of blockList(copy)) if (localName(block.name) === "p") wrapContainerContent(block, "del", ctx);
      part.append(copy);
      changed = true;
    }
  }
  if (changed) merged.markNotePartsChanged([kind === "footnote" ? "footnotes" : "endnotes"]);
  return ids;
}

export function blockList(container: XmlElement, ...skip: string[]): XmlElement[] {
  return container.children.filter((c) => !skip.includes(localName(c.name)));
}

/** Say so when a part outside the body differs, rather than imply it matched. */
export function reportUncomparedParts(original: DocxDocument, revised: DocxDocument, ctx: Ctx): void {
  // Headers, footers and notes are compared (compareHeadersFooters,
  // compareNotes). Style and numbering DEFINITIONS are not: Word's compare
  // does not track them either, and their effect shows on the paragraphs.
  const parts: [string, XmlElement | null][] = [
    ["styles.xml", original.stylesTree()],
    ["numbering.xml", original.numberingTree()],
  ];
  const revisedTrees: Record<string, XmlElement | null> = {
    "styles.xml": revised.stylesTree(),
    "numbering.xml": revised.numberingTree(),
  };
  for (const [name, tree] of parts) {
    const other = revisedTrees[name];
    const a = tree ? serializeXml(tree) : "";
    const b = other ? serializeXml(other) : "";
    if (a !== b) {
      ctx.note({ code: "part-not-compared", detail: `${name} differs; the result carries the revised document's` });
    }
  }
}

// ---------------------------------------------------------------------------
// Block alignment
// ---------------------------------------------------------------------------

/**
 * Align two block sequences and emit the merged sequence, tracked.
 *
 * Two passes, following GumTree's shape:
 *
 *  1. **Anchor.** Fingerprint each block and run histogram diff over the
 *     fingerprints. Every matched pair is an anchor. In a typical revision
 *     this settles the large majority of the document with no similarity
 *     computation at all.
 *  2. **Fill the gaps.** Between consecutive anchors sits a small block of
 *     unmatched paragraphs on each side. Only THERE is pairwise similarity
 *     computed, and only there can a paragraph be paired. Bounding the search
 *     to a gap is what makes it structurally impossible for a paragraph on
 *     page 40 to match one on page 2 — the mis-alignment that cascades into
 *     unreadable output.
 */
function alignBlocks(oBlocks: XmlElement[], rBlocks: XmlElement[], ctx: Ctx): Entry[] {
  const anchors = histogramMatch(
    oBlocks.map((el) => blockFingerprint(el, ctx.original)),
    rBlocks.map((el) => blockFingerprint(el, ctx.doc)),
  );
  const entries: Entry[] = [];
  let oAt = 0;
  let rAt = 0;
  for (const anchor of [...anchors, { a: oBlocks.length, b: rBlocks.length }]) {
    emitGap(oBlocks, rBlocks, oAt, anchor.a, rAt, anchor.b, entries, ctx);
    if (anchor.a < oBlocks.length && anchor.b < rBlocks.length) {
      entries.push(emitMatched(oBlocks[anchor.a], rBlocks[anchor.b], ctx));
    }
    oAt = anchor.a + 1;
    rAt = anchor.b + 1;
  }
  return entries;
}

/**
 * A block's identity for the anchor pass: its normalized text plus the two
 * things that distinguish otherwise identical paragraphs. Three chapters that
 * each open with "Introduction" usually differ in outline level or list
 * numbering, and the fingerprint carries both, so histogram's lowest-
 * occurrence rule has something to separate them by.
 */
export function blockFingerprint(el: XmlElement, doc: DocxDocument): string {
  const ln = localName(el.name);
  if (ln === "p") {
    const pPr = child(el, "pPr");
    const style = child(pPr, "pStyle")?.attrs["w:val"] ?? "";
    const numPr = child(pPr, "numPr");
    // Alignment too: probe3-thai holds the same paragraph justified and then
    // left-aligned, and without it the matcher anchored on the wrong copy.
    const jc = child(pPr, "jc")?.attrs["w:val"] ?? "";
    // And the section it ends, if any: see pairable.
    const sect = sectionKey(el, doc);
    return `P\u0000${normalizeText(paragraphText(el))}\u0000${style}\u0000${numPr ? serializeXml(numPr) : ""}\u0000${jc}\u0000${sect}`;
  }
  if (ln === "tbl") return tableFingerprint(paragraphText(el));
  return `X\u0000${serializeXml(el)}`;
}

/** blockFingerprint of a table, from its paragraphText. */
export function tableFingerprint(text: string): string {
  return `T\u0000${normalizeText(text)}`;
}

export interface Pairing {
  kind: "pair" | "split" | "merge";
  o: number;
  r: number;
  /** "split": the second revised paragraph. "merge": the second original one. */
  other: number;
}

/**
 * Resolve one gap between anchors, then emit it in document order: at each
 * step the unmatched originals come first (struck), then the unmatched
 * revisions (inserted), then the matched pair.
 */
function emitGap(
  oBlocks: XmlElement[],
  rBlocks: XmlElement[],
  oLo: number,
  oHi: number,
  rLo: number,
  rHi: number,
  entries: Entry[],
  ctx: Ctx,
): void {
  // A gap with no paragraph on either side has no mark to stand on: an added
  // or removed paragraph there can only join a neighbour that is a table, and
  // review cannot do that (a heading rewritten between two tables came back
  // from reject-all with an empty paragraph after it). Such a gap, like one
  // opening the container, must pair something and diff it in place.
  const before = entries[entries.length - 1];
  const unanchored = (!before || !isParagraph(before.el)) && (rHi >= rBlocks.length || !isParagraph(rBlocks[rHi]));
  const pairings = resolveGap(
    treeSide(oBlocks, ctx.original, true, ctx),
    treeSide(rBlocks, ctx.doc, false, ctx),
    oLo,
    oHi,
    rLo,
    rHi,
    (oLo === 0 && rLo === 0) || unanchored,
  );
  for (const op of gapOps(pairings, oLo, oHi, rLo, rHi)) {
    for (const entry of emitGapOp(op, (i) => oBlocks[i], (j) => rBlocks[j], ctx)) entries.push(entry);
  }
}

/** One step of a gap's emission, in document order. */
export type GapOp =
  | { kind: "deleted"; o: number }
  | { kind: "inserted"; r: number }
  | Pairing;

/** A gap's pairings as the steps that emit it, in document order. */
export function* gapOps(pairings: readonly Pairing[], oLo: number, oHi: number, rLo: number, rHi: number): Generator<GapOp> {
  let oAt = oLo;
  let rAt = rLo;
  for (const pairing of [...pairings, { kind: "pair" as const, o: oHi, r: rHi, other: -1 }]) {
    for (let i = oAt; i < pairing.o; i++) yield { kind: "deleted", o: i };
    for (let j = rAt; j < pairing.r; j++) yield { kind: "inserted", r: j };
    if (pairing.o >= oHi || pairing.r >= rHi) return;
    yield pairing;
    oAt = pairing.kind === "merge" ? pairing.other + 1 : pairing.o + 1;
    rAt = pairing.kind === "split" ? pairing.other + 1 : pairing.r + 1;
  }
}

/** Emit one gap step: the entries it adds, given each side's blocks by index. */
export function emitGapOp(op: GapOp, o: (i: number) => XmlElement, r: (j: number) => XmlElement, ctx: Ctx): Entry[] {
  if (op.kind === "deleted") return [emitDeleted(o(op.o), ctx)];
  if (op.kind === "inserted") return [emitInserted(r(op.r), ctx)];
  if (op.kind === "split") {
    // One original paragraph became two. Their content already IS the
    // original's, word for word and property for property, so there is no
    // text revision to make — only the new pilcrow between them, which the
    // mark pass records against the first paragraph.
    const kept = r(op.r);
    alignParagraphProps(kept, child(kept, "pPr"), child(o(op.o), "pPr"), ctx);
    return [
      { el: kept, cls: "kept" },
      { el: r(op.other), cls: "inserted" },
    ];
  }
  if (op.kind === "merge") {
    // Two originals became one: both are emitted, and the pilcrow between
    // them is struck. Their content is already what the revised paragraph
    // holds, so there is no text revision to make.
    const first = cloneXml(o(op.o));
    adoptFromOriginal(first, ctx);
    alignParagraphProps(first, child(r(op.r), "pPr"), child(o(op.o), "pPr"), ctx);
    const second = cloneXml(o(op.other));
    adoptFromOriginal(second, ctx);
    return [
      { el: first, cls: "kept" },
      { el: second, cls: "deleted" },
    ];
  }
  return [emitMatched(o(op.o), r(op.r), ctx)];
}

/**
 * One side of a gap, read by index: what resolveGap needs from each block.
 * compareDocuments reads it from the tree; a caller holding blocks elsewhere
 * can answer from there.
 */
export interface GapSide {
  /** The block's local name: p, tbl, sdt, tr... */
  kind(i: number): string;
  /** paragraphText of the block. */
  text(i: number): string;
  /** sectionKey of the block, against its own document. */
  section(i: number): string;
  /** A table row's cell count. */
  cells(i: number): number;
  /** paragraphKeyParts of the block. */
  keyParts(i: number): string[] | null;
}

/** A gap side over blocks in the tree. Its key parts are kept for the one gap that reads them. */
function treeSide(els: XmlElement[], doc: DocxDocument, original: boolean, ctx: Ctx): GapSide {
  const keys = new Map<number, string[] | null>();
  return {
    kind: (i) => localName(els[i].name),
    text: (i) => paragraphText(els[i]),
    section: (i) => sectionKey(els[i], doc),
    cells: (i) => els[i].children.filter((c) => localName(c.name) === "tc").length,
    keyParts: (i) => {
      let parts = keys.get(i);
      if (parts === undefined) {
        parts = paragraphKeyParts(els[i], original, ctx);
        keys.set(i, parts);
      }
      return parts;
    },
  };
}

/** Gap size (original x revised blocks) past which pairing turns windowed. */
const GAP_PAIR_CAP = 200 * 200;
/** How far either side of its own relative position a block looks, windowed. */
const GAP_WINDOW = 50;

export function resolveGap(
  O: GapSide,
  R: GapSide,
  oLo: number,
  oHi: number,
  rLo: number,
  rHi: number,
  needsAnchor = false,
): Pairing[] {
  // One block on each side: paired when they are the same shape, as the
  // single-leftover rule would pair them, or else when scoring would.
  // Settled first, so two long tables are not tokenized to be scored.
  if (oHi - oLo === 1 && rHi - rLo === 1) {
    const pair =
      sameShape(O, oLo, R, rLo) ||
      (pairable(O, oLo, R, rLo) && (needsAnchor || similarity(O.text(oLo), R.text(rLo)) >= similarityFloor(O.text(oLo), R.text(rLo))));
    return pair ? [{ kind: "pair", o: oLo, r: rLo, other: -1 }] : [];
  }

  const out: Pairing[] = [];
  const oUsed = new Set<number>();
  const rUsed = new Set<number>();

  // Splits and merges first: they are exact (the content keys concatenate),
  // and they are the shape a similarity pass would otherwise turn into
  // delete-one-insert-two, which reads as a rewrite of something nobody
  // rewrote.
  // Accepted pairings are monotone, so a candidate crosses one exactly when
  // it crosses a neighbour in o order: a binary search, not a scan of all.
  const placed: Pairing[] = [];
  const crosses = (o: number, r: number): boolean => {
    let lo = 0;
    let hi = placed.length;
    while (lo < hi) {
      const mid = (lo + hi) >> 1;
      if (placed[mid].o < o) lo = mid + 1;
      else hi = mid;
    }
    return (lo > 0 && placed[lo - 1].r >= r) || (lo < placed.length && placed[lo].r <= r);
  };
  const accept = (p: Pairing): void => {
    out.push(p);
    let at = placed.length;
    while (at > 0 && placed[at - 1].o > p.o) at--;
    placed.splice(at, 0, p);
  };

  // A long gap with nothing anchoring it (a stretch rewritten end to end) made
  // every original block meet every revised one: quadratic in the gap. Past
  // GAP_PAIR_CAP pairs, each original block meets only the revised blocks
  // within GAP_WINDOW of its own relative position, which still finds every
  // local edit and keeps the cost linear in the gap.
  const oLen = oHi - oLo;
  const rLen = rHi - rLo;
  const windowed = oLen * rLen > GAP_PAIR_CAP;
  const span = (i: number): [number, number] => {
    if (!windowed) return [rLo, rHi];
    const centre = rLo + Math.round(((i - oLo) * rLen) / Math.max(1, oLen));
    return [Math.max(rLo, centre - GAP_WINDOW), Math.min(rHi, centre + GAP_WINDOW + 1)];
  };

  for (let i = oLo; i < oHi; i++) {
    const [jLo, jHi] = span(i);
    for (let j = jLo; j + 1 < jHi; j++) {
      if (oUsed.has(i) || rUsed.has(j) || rUsed.has(j + 1) || crosses(i, j)) continue;
      if (concatMatches(O, [i], R, [j, j + 1])) {
        accept({ kind: "split", o: i, r: j, other: j + 1 });
        oUsed.add(i);
        rUsed.add(j);
        rUsed.add(j + 1);
      }
    }
  }
  for (let i = oLo; i + 1 < oHi; i++) {
    const [jLo, jHi] = span(i);
    for (let j = jLo; j < jHi; j++) {
      if (oUsed.has(i) || oUsed.has(i + 1) || rUsed.has(j) || crosses(i, j)) continue;
      if (concatMatches(O, [i, i + 1], R, [j])) {
        accept({ kind: "merge", o: i, r: j, other: i + 1 });
        oUsed.add(i);
        oUsed.add(i + 1);
        rUsed.add(j);
      }
    }
  }

  // Then similarity, best score first, over the gap's cross-product only.
  // The candidates are three columns, not objects: a windowed gap holds up
  // to 2 × GAP_WINDOW of them per original block, and a long rewritten list
  // held millions.
  const scored = new Candidates();
  for (let i = oLo; i < oHi; i++) {
    if (oUsed.has(i)) continue;
    const [jLo, jHi] = span(i);
    for (let j = jLo; j < jHi; j++) {
      if (rUsed.has(j)) continue;
      if (!pairable(O, i, R, j)) continue;
      const a = O.text(i);
      const b = R.text(j);
      const score = similarity(a, b);
      if (score >= similarityFloor(a, b)) scored.push(i, j, score);
    }
  }
  for (const k of scored.byScore()) {
    const o = scored.o[k];
    const r = scored.r[k];
    if (oUsed.has(o) || rUsed.has(r)) continue;
    // Pairings must stay monotone: a crossing pair would put a deletion after
    // the insertion that replaced it and read as two unrelated edits.
    if (crosses(o, r)) continue;
    accept({ kind: "pair", o, r, other: -1 });
    oUsed.add(o);
    rUsed.add(r);
  }

  // A slot between two settled pairings (or a gap's ends) that holds exactly
  // one leftover paragraph on each side is that paragraph, edited where it
  // stands. Similarity cannot see it when the paragraph is short: one new word
  // in a 3-word heading misses the 0.8 short-paragraph floor, and the heading
  // came out struck and reinserted whole (91% of the compare benchmark's
  // misses). A MOVED item never lands here: it leaves a removal in one slot
  // and an insertion in another, so a reordered list still reads as a move.
  // Tables pair the same way, and table rows when their cell counts agree: a
  // two-row table whose one word changed was struck and reinserted whole.
  const settled = [...out].sort((a, b) => a.o - b.o);
  let oFrom = oLo;
  let rFrom = rLo;
  for (const p of [...settled, { kind: "pair" as const, o: oHi, r: rHi, other: -1 }]) {
    const os: number[] = [];
    const rs: number[] = [];
    for (let i = oFrom; i < p.o; i++) if (!oUsed.has(i)) os.push(i);
    for (let j = rFrom; j < p.r; j++) if (!rUsed.has(j)) rs.push(j);
    if (os.length === 1 && rs.length === 1 && sameShape(O, os[0], R, rs[0])) {
      accept({ kind: "pair", o: os[0], r: rs[0], other: -1 });
      oUsed.add(os[0]);
      rUsed.add(rs[0]);
    }
    oFrom = p.kind === "merge" ? p.other + 1 : p.o + 1;
    rFrom = p.kind === "split" ? p.other + 1 : p.r + 1;
  }

  // A gap that OPENS the container and pairs nothing has to pair something
  // anyway. Every added or removed paragraph is recorded by striking the
  // pilcrow of the paragraph BEFORE it (see applyParagraphMarks), so a
  // container whose every block is added or removed has no pilcrow left to
  // stand on and its last paragraph mark cannot be tracked at all.
  //
  // Pairing the first leftover on each side gives that anchor. It costs
  // nothing in fidelity: the pair is then word-diffed, and two paragraphs with
  // no words in common come out as everything struck followed by everything
  // inserted — the same content a delete-plus-insert shows, in one paragraph
  // instead of two, which is also what Word shows for a rewritten paragraph.
  if (needsAnchor && out.length === 0) {
    for (let i = oLo; i < oHi && out.length === 0; i++) {
      if (oUsed.has(i)) continue;
      for (let j = rLo; j < rHi; j++) {
        if (rUsed.has(j) || !pairable(O, i, R, j)) continue;
        accept({ kind: "pair", o: i, r: j, other: -1 });
        break;
      }
    }
  }
  return out.sort((a, b) => a.o - b.o);
}

/** The scored pairs of a gap, as three columns. */
class Candidates {
  o = new Int32Array(1024);
  r = new Int32Array(1024);
  score = new Float64Array(1024);
  n = 0;

  push(o: number, r: number, score: number): void {
    if (this.n === this.o.length) {
      const o2 = new Int32Array(this.n * 2);
      o2.set(this.o);
      this.o = o2;
      const r2 = new Int32Array(this.n * 2);
      r2.set(this.r);
      this.r = r2;
      const score2 = new Float64Array(this.n * 2);
      score2.set(this.score);
      this.score = score2;
    }
    this.o[this.n] = o;
    this.r[this.n] = r;
    this.score[this.n] = score;
    this.n++;
  }

  /** The candidates' indexes, best score first, then by original block, then as pushed. */
  byScore(): Uint32Array {
    const order = new Uint32Array(this.n);
    for (let k = 0; k < this.n; k++) order[k] = k;
    return order.sort((x, y) => this.score[y] - this.score[x] || this.o[x] - this.o[y] || x - y);
  }
}

/**
 * Two blocks of one kind that may be paired and diffed in place. A paragraph
 * that ends a section pairs only with one ending the SAME section: the break
 * lives in the paragraph mark, a pPrChange cannot carry it, and a sectPrChange
 * cannot carry header or footer references. Paired anyway, the original's
 * section was lost on reject-all (probe-headerheight with blocks moved: six
 * empty section-break paragraphs shared one fingerprint, the H721 break was
 * anchored to the H720 one, and its header went missing). A changed break
 * shows as struck and reinserted, which reviews back to either side.
 */
function pairable(O: GapSide, i: number, R: GapSide, j: number): boolean {
  if (O.kind(i) !== R.kind(j)) return false;
  if (O.kind(i) !== "p") return true;
  return O.section(i) === R.section(j);
}

/**
 * The section a paragraph ends, as its w:sectPr with each relationship id
 * replaced by the part it points to, so the two documents' ids compare; "" for
 * a paragraph that ends no section.
 */
export function sectionKey(el: XmlElement, doc: DocxDocument): string {
  const sectPr = child(child(el, "pPr"), "sectPr");
  if (!sectPr) return "";
  return serializeXml(sectPr).replace(/r:id="([^"]*)"/g, (_, id: string) => `r:id="${doc.documentRels.get(id)?.target ?? id}"`);
}

/**
 * Two paragraphs, two tables, two block-level content controls (or custom XML
 * elements), or two table rows with the same number of cells.
 */
function sameShape(O: GapSide, i: number, R: GapSide, j: number): boolean {
  if (O.kind(i) === "p" && R.kind(j) === "p") return pairable(O, i, R, j);
  if (O.kind(i) === R.kind(j) && ["tbl", "sdt", "customXml"].includes(O.kind(i))) return true;
  return O.kind(i) === "tr" && R.kind(j) === "tr" && O.cells(i) === R.cells(j);
}

/**
 * True when the two sides carry exactly the same content — same words, same
 * atoms, same run properties — once concatenated. That exactness is what makes
 * a split or a merge emit ONE paragraph-mark revision and no text revision,
 * and what makes rejecting it restore the original byte for byte.
 */
/**
 * How to read a paragraph from one side. A note reference keys on the revised
 * id its note maps to: the original's through compareNotes' pairing (a removed
 * note's copy has an id no revised reference carries), the revised's as is.
 */
export function readOptions(ctx: Ctx, original: boolean): ReadOptions {
  const noteKey = (ref: XmlElement): string => {
    const kind = localName(ref.name) === "footnoteReference" ? "footnote" : "endnote";
    const id = ref.attrs["w:id"];
    return original ? `${ctx.noteIds[kind].get(id) ?? `original ${id}`}` : id;
  };
  return { whitespace: ctx.whitespace, noteKey };
}

function concatMatches(O: GapSide, oAt: number[], R: GapSide, rAt: number[]): boolean {
  const a = concatKey(O, oAt);
  if (!a) return false;
  const b = concatKey(R, rAt);
  return !!b && a.length === b.length && a.every((part, i) => part === b[i]);
}

function concatKey(side: GapSide, at: number[]): string[] | null {
  const out: string[] = [];
  for (const i of at) {
    const parts = side.keyParts(i);
    if (!parts) return null;
    for (const part of parts) out.push(part);
  }
  return out;
}

/**
 * One paragraph's tokens with their run properties and containers, or null
 * when it cannot take part in a split or merge. A gap side reads it once per
 * gap: the split and merge passes try every nearby pair, and re-reading both
 * paragraphs for each try was half the time of a long rewritten document
 * (the Yiddish novel, 90% of paragraphs edited: 3.4 of 6.9 s). Kept for one
 * gap only, so a long document's keys are not all alive at once.
 */
export function paragraphKeyParts(el: XmlElement, original: boolean, ctx: Ctx): string[] | null {
  let parts: string[] | null = null;
  // A section break is no split or merge: its mark is the break, and the
  // join a review makes would move it (wild-gatech with blocks moved: an
  // empty paragraph and an empty break merged into one, losing a footer).
  if (localName(el.name) === "p" && !child(child(el, "pPr"), "sectPr")) {
    const content = readParagraph(el, readOptions(ctx, original));
    if (content.inlineDiffable) {
      parts = content.tokens.map((token) => {
        const props = token.pieces ? token.pieces.map((p) => `${p.text.length}:${runPropsKey(p.run)}`).join("|") : runPropsKey(token.run);
        return token.key + "\u0002" + props + "\u0002" + token.path.map((c) => serializeXml({ ...c, children: c.children.filter((k) => /Pr$/.test(localName(k.name))) })).join(">");
      });
    }
  }
  return parts;
}

// ---------------------------------------------------------------------------
// Emission
// ---------------------------------------------------------------------------

/**
 * A matched pair. Both sides are always the same kind of block: fingerprints
 * are namespaced by kind, and the gap pass refuses to pair across kinds.
 */
export function emitMatched(oEl: XmlElement, rEl: XmlElement, ctx: Ctx): Entry {
  const ln = localName(oEl.name);
  if (ln === "p") comparePair(oEl, rEl, ctx);
  else if (ln === "tbl") compareTable(oEl, rEl, ctx);
  // A block-level content control or custom XML element: compare what it
  // holds. Passing it through let an edited paragraph inside a content
  // control come out with no tracked change at all.
  else if (ln === "sdt") {
    const o = child(oEl, "sdtContent");
    const r = child(rEl, "sdtContent");
    if (o && r) compareContainer(o, r, ctx);
  } else if (ln === "customXml") compareContainer(oEl, rEl, ctx);
  return { el: rEl, cls: "kept" };
}

/**
 * Relationship ids in `r:` attributes, which point INTO a package's own
 * relationship part.
 */
const REL_ATTRS = ["id", "embed", "link", "pict", "dm", "lo", "qs", "cs"];

/**
 * Re-point a subtree cloned out of the ORIGINAL document at the result's own
 * relationships.
 *
 * A struck paragraph keeps its content, and that content can carry `r:id="rId7"`
 * — an index into the original package's rels, which the result does not have
 * and where rId7 may well mean something else. Left alone, a deleted hyperlink
 * or image is a reference into nothing, and Word refuses the file.
 *
 * An external target (a hyperlink) is re-added, which is exact. Anything else
 * points at a PART — an image, a chart, an embedded object — and that part is
 * copied across with the parts it references (DocxDocument.adoptRelFrom), so
 * reject-all brings a removed picture back with its image, not as an empty
 * frame. Only a part the original does not actually carry is dropped and
 * reported, rather than left dangling.
 */
export function adoptFromOriginal(root: XmlElement, ctx: Ctx): void {
  const mine = ctx.doc.relationshipsOf(ctx.scope.into);
  const theirs = ctx.original.relationshipsOf(ctx.scope.from);
  const remapped = new Map<string, string | null>();
  const walk = (el: XmlElement): void => {
    // A note reference points into the ORIGINAL's notes; compareNotes has
    // settled where that note lives in the result.
    const ln = localName(el.name);
    if (ln === "footnoteReference" || ln === "endnoteReference") {
      const to = ctx.noteIds[ln === "footnoteReference" ? "footnote" : "endnote"].get(el.attrs["w:id"]);
      if (to !== undefined) el.attrs["w:id"] = to;
    }
    if (ln === "commentReference" || ln === "commentRangeStart" || ln === "commentRangeEnd") {
      el.attrs["w:id"] = adoptComment(el.attrs["w:id"], ctx);
    }
    for (const key of Object.keys(el.attrs)) {
      if (!key.startsWith("r:") || !REL_ATTRS.includes(localName(key))) continue;
      const id = el.attrs[key];
      const source = theirs.get(id);
      if (!source) continue;
      // The same relationship to the same part: nothing to adopt. The same
      // part name with other bytes is another part (a SmartArt diagram's
      // data with a word changed), and reject-all needs the original's.
      const here = mine.get(id);
      if (here && here.target === source.target && here.type === source.type && samePart(source.target, ctx)) continue;
      let replacement = remapped.get(id);
      if (replacement === undefined) {
        replacement = source.external
          ? ctx.doc.addHyperlinkRel(source.target, ctx.scope.into)
          : ctx.doc.adoptRelFrom(ctx.original, id, ctx.scope);
        if (replacement === null) {
          ctx.note({
            code: "unresolved-relationship",
            detail: `removed content referenced ${source.target}, which the result does not carry; the reference was dropped`,
          });
        }
        remapped.set(id, replacement);
      }
      if (replacement === null) delete el.attrs[key];
      else el.attrs[key] = replacement;
    }
    for (const c of el.children) walk(c);
  };
  walk(root);
}

/**
 * True when the result and the original hold the same bytes at `target`. Only
 * an XML part (a SmartArt diagram's data, a chart) is read; a binary part
 * under one name is taken as the same, as it always was.
 */
function samePart(target: string, ctx: Ctx): boolean {
  if (!target.endsWith(".xml")) return true;
  const mine = ctx.doc.pkg.binary(target);
  const theirs = ctx.original.pkg.binary(target);
  return !!mine && !!theirs && mine.length === theirs.length && mine.every((b, i) => b === theirs[i]);
}

/**
 * Whether `el` refers to a part the two documents both hold under one name
 * with different contents: a SmartArt diagram with a word changed leaves its
 * paragraph's XML as it was and changes only the diagram's data part.
 */
export function refersToChangedPart(el: XmlElement, ctx: Ctx): boolean {
  const mine = ctx.doc.relationshipsOf(ctx.scope.into);
  const theirs = ctx.original.relationshipsOf(ctx.scope.from);
  const walk = (e: XmlElement): boolean =>
    Object.keys(e.attrs).some((key) => {
      if (!key.startsWith("r:") || !REL_ATTRS.includes(localName(key))) return false;
      const source = theirs.get(e.attrs[key]);
      const here = mine.get(e.attrs[key]);
      return !!source && !!here && !source.external && source.target === here.target && !samePart(source.target, ctx);
    }) || e.children.some(walk);
  return walk(el);
}

/**
 * The result's id for the original's comment `id`, which struck content
 * refers to: the id itself when the result holds that comment unchanged,
 * else a copy of the comment under a fresh id. Without the copy, reject-all
 * brought a removed comment's reference back pointing at nothing
 * (parity-comments with a comment removed).
 */
function adoptComment(id: string, ctx: Ctx): string {
  const known = ctx.commentIds.get(id);
  if (known !== undefined) return known;
  const find = (root: XmlElement | null): XmlElement | undefined =>
    root?.children.find((c) => localName(c.name) === "comment" && c.attrs["w:id"] === id);
  const theirs = find(ctx.original.commentsTree());
  const mine = find(ctx.doc.commentsTree());
  let to = id;
  if (theirs && !(mine && sameXml(theirs, mine))) {
    const root = ctx.doc.commentsTree(true) as XmlElement;
    let next = 0;
    for (const c of root.children) next = Math.max(next, 1 + (Number(c.attrs["w:id"]) || 0));
    const copy = cloneXml(theirs);
    copy.attrs["w:id"] = String(next);
    root.children.push(copy);
    ctx.doc.markCommentsChanged();
    to = String(next);
  }
  ctx.commentIds.set(id, to);
  return to;
}

export function emitDeleted(el: XmlElement, ctx: Ctx): Entry {
  const clone = cloneXml(el);
  adoptFromOriginal(clone, ctx);
  markWholeBlock(clone, "del", ctx);
  return { el: clone, cls: "deleted" };
}

export function emitInserted(el: XmlElement, ctx: Ctx): Entry {
  markWholeBlock(el, "ins", ctx);
  return { el, cls: "inserted" };
}

/**
 * Mark a whole block removed or added: a paragraph's content (its mark is
 * the mark pass's), a table's rows, and everything a block-level content
 * control or custom XML element holds, its paragraphs' marks included, since
 * the mark pass sees only the control. Left unmarked, a removed control came
 * through accept-all whole (probe2-content-controls with blocks moved).
 */
function markWholeBlock(el: XmlElement, kind: "ins" | "del", ctx: Ctx): void {
  const ln = localName(el.name);
  if (ln === "p") {
    wrapContainerContent(el, kind, ctx);
  } else if (ln === "tbl") {
    for (const row of el.children.filter((c) => localName(c.name) === "tr")) {
      markRowRevision(row, kind, ctx.meta);
      wrapRowContent(row, kind, ctx);
    }
  } else if (ln === "sdt" || ln === "customXml") {
    const inner = ln === "sdt" ? child(el, "sdtContent") : el;
    for (const block of inner?.children ?? []) {
      markWholeBlock(block, kind, ctx);
      if (isParagraph(block)) markParagraphGlyph(block, kind, ctx.meta);
    }
  } else if (!isPassenger(ln) && ln !== "sdtPr" && ln !== "sdtEndPr" && ln !== "customXmlPr") {
    ctx.note({ code: "untracked-block", detail: `a <${el.name}> block was ${kind === "ins" ? "added" : "removed"} and is not expressible as a revision` });
  }
}

function wrapRowContent(row: XmlElement, kind: "ins" | "del", ctx: Ctx): void {
  for (const cell of row.children.filter((c) => localName(c.name) === "tc")) {
    for (const block of cell.children) {
      if (localName(block.name) === "p") wrapContainerContent(block, kind, ctx);
    }
  }
}

/**
 * Wrap every run below `container` in one revision wrapper per contiguous run
 * of them.
 *
 * A revision wrapper may not sit AROUND a `w:hyperlink` — the schema puts
 * w:ins inside it — so a container child that holds runs of its own is
 * recursed into rather than wrapped, which is exactly what Word writes.
 */
export function wrapContainerContent(container: XmlElement, kind: "ins" | "del", ctx: Ctx): void {
  // Run containers the schema lets a revision sit INSIDE (a simple field's
  // result runs among them), as opposed to an inline content control or a
  // run-level customXml, which a revision may wrap whole. Left plain, a removed
  // paragraph's checkbox controls ("☒ ☐") survived accept-all and joined the
  // paragraph before (probe2-form-checkboxes).
  const NESTED = ["hyperlink", "smartTag", "dir", "bdo", "fldSimple"];
  // An equation goes whole: a w:ins or w:del may hold m:oMath and
  // m:oMathPara (EG_MathContent, through EG_RunLevelElements). Left plain,
  // a struck paragraph's equation survived accept-all beside the inserted
  // one, so the result showed it twice (parity-math2).
  const WRAPPABLE = ["r", "sdt", "customXml", "oMath", "oMathPara"];
  const out: XmlElement[] = [];
  let group: XmlElement[] = [];
  const flush = (): void => {
    if (group.length === 0) return;
    if (kind === "del") for (const g of group) markDeletedText(g);
    out.push(revisionWrapper(container, kind, group, ctx.meta));
    group = [];
  };
  for (const c of container.children) {
    const ln = localName(c.name);
    if (WRAPPABLE.includes(ln)) {
      group.push(c);
      continue;
    }
    flush();
    if (NESTED.includes(ln)) wrapContainerContent(c, kind, ctx);
    else if (ln !== "pPr" && ln !== "tcPr" && !isPassenger(ln)) {
      ctx.note({
        code: "untracked-block",
        detail: `<${c.name}> inside a ${kind === "ins" ? "new" : "removed"} paragraph is left as a plain difference`,
      });
    }
    out.push(c);
  }
  flush();
  container.children = out;
}

// ---------------------------------------------------------------------------
// Paragraph pair
// ---------------------------------------------------------------------------

/** Diff one matched paragraph pair, mutating the revised element in place. */
function comparePair(oEl: XmlElement, rEl: XmlElement, ctx: Ctx): void {
  const partChanged = refersToChangedPart(oEl, ctx);
  if (!partChanged && sameXml(oEl, rEl)) return;

  if (ctx.formatting) alignParagraphProps(rEl, child(rEl, "pPr"), child(oEl, "pPr"), ctx);

  const oc = readParagraph(oEl, readOptions(ctx, true));
  const rc = readParagraph(rEl, readOptions(ctx, false));
  if (partChanged || !oc.inlineDiffable || !rc.inlineDiffable) {
    if (partChanged || oc.text !== rc.text || serializeXml(stripProps(oEl)) !== serializeXml(stripProps(rEl))) {
      // Too structured to rebuild safely: strike the old content whole and
      // insert the new content whole, inside this one paragraph.
      const deleted = cloneXml(oEl);
      adoptFromOriginal(deleted, ctx);
      wrapContainerContent(deleted, "del", ctx);
      wrapContainerContent(rEl, "ins", ctx);
      const pPr = child(rEl, "pPr");
      rEl.children = [
        ...(pPr ? [pPr] : []),
        ...deleted.children.filter((c) => localName(c.name) !== "pPr"),
        ...rEl.children.filter((c) => localName(c.name) !== "pPr"),
      ];
      ctx.note({
        code: "coarse-paragraph",
        detail: "a paragraph holding inline math or other content a word diff cannot rebuild changed; struck and reinserted whole",
      });
    }
    return;
  }

  const segments = wordSegments(oc.tokens, rc.tokens);
  const changed = segments.some((s) => !s.equal);
  // Same words in the same order: any remaining difference is formatting.
  // When the runs split differently on the two sides (a bold "Purpose:"
  // against bold "Purpose" + plain ":"), fall through to the rebuild below,
  // which records formatting token by token.
  // The in-place shortcut sees top-level runs only, so a paragraph with run
  // containers (a hyperlink, a content control) always takes the rebuild.
  const flat = oc.tokens.every((t) => t.path.length === 0) && rc.tokens.every((t) => t.path.length === 0);
  if (!changed && (!ctx.formatting || (flat && alignRunPropsInPlace(oEl, rEl, ctx)))) return;

  // Emit flat, each element with its containers and its revision kind, then
  // nest: revisions go INSIDE a hyperlink or content control, never around
  // a hyperlink. Struck words from a container the revised paragraph still
  // has (matched through an unchanged word inside it) go into that
  // container; a container only the original had is carried across.
  const pPr = child(rEl, "pPr");
  const items: Item[] = [];
  const counterpart = new Map<XmlElement, XmlElement>();
  for (const seg of segments) {
    if (seg.equal) {
      for (const it of emitEqualSpan(oc.tokens, rc.tokens, seg, ctx, counterpart)) items.push({ ...it, kind: null });
      continue;
    }
    if (seg.aEnd > seg.aStart) {
      for (const it of sliceItems(oc.tokens, seg.aStart, seg.aEnd, true)) {
        adoptFromOriginal(it.el, ctx);
        markDeletedText(it.el);
        items.push({ ...it, kind: "del" });
      }
    }
    if (seg.bEnd > seg.bStart) {
      for (const it of sliceItems(rc.tokens, seg.bStart, seg.bEnd, false)) items.push({ ...it, kind: "ins" });
    }
  }
  for (const it of items) if (it.kind === "del") it.path = it.path.map((c) => counterpart.get(c) ?? c);
  for (const el of rc.tail) items.push({ el: cloneXml(el), path: [], kind: null });
  // Containers met in the original paragraph: any still in a path after the
  // mapping above has no revised counterpart and is carried across.
  const fromOriginal = new Set(oc.tokens.flatMap((t) => t.path));
  rEl.children = [...(pPr ? [pPr] : []), ...nestItems(items, 0, rEl, ctx, fromOriginal)];
}

type Item = Placed & { kind: "ins" | "del" | null };

/** Containers a revision may wrap whole (run content in the schema). */
const WRAPPABLE_CONTAINERS = new Set(["sdt", "customXml"]);

/**
 * Compare the text boxes of a drawing matched by shape: each w:txbxContent
 * is a small document of its own, aligned and word-diffed like the body. A
 * shape stored twice (DrawingML with a VML fallback) carries the box twice,
 * and the copies pair in order.
 */
function compareTextBoxes(oAtom: XmlElement, rAtom: XmlElement, ctx: Ctx): void {
  const theirs = textBoxesIn(oAtom);
  const mine = textBoxesIn(rAtom);
  for (let i = 0; i < Math.min(theirs.length, mine.length); i++) {
    if (sameXml(theirs[i], mine[i])) continue;
    const entries = alignBlocks(blockList(theirs[i]), blockList(mine[i]), ctx);
    applyParagraphMarks(entries, ctx);
    mine[i].children = entries.map((e) => e.el);
  }
}

/** Rebuild the container tree around flat items, wrapping each revised stretch. */
function nestItems(items: Item[], depth: number, rEl: XmlElement, ctx: Ctx, fromOriginal: Set<XmlElement>): XmlElement[] {
  const out: XmlElement[] = [];
  let i = 0;
  while (i < items.length) {
    const c = items[i].path[depth];
    let j = i + 1;
    if (c) {
      while (j < items.length && items[j].path[depth] === c) j++;
      const group = items.slice(i, j);
      const kind = group[0].kind;
      if (kind && WRAPPABLE_CONTAINERS.has(localName(c.name)) && group.every((it) => it.kind === kind)) {
        // A whole content control added or removed: wrap it, so review takes
        // the control itself away rather than leaving an empty one.
        const plain = group.map((it) => ({ ...it, kind: null }));
        out.push(revisionWrapper(rEl, kind, [containerShell(c, nestItems(plain, depth + 1, rEl, ctx, fromOriginal), fromOriginal.has(c), ctx)], ctx.meta));
      } else {
        out.push(containerShell(c, nestItems(group, depth + 1, rEl, ctx, fromOriginal), fromOriginal.has(c), ctx));
      }
    } else {
      const kind = items[i].kind;
      while (j < items.length && !items[j].path[depth] && items[j].kind === kind) j++;
      const els = items.slice(i, j).map((it) => it.el);
      if (kind) out.push(revisionWrapper(rEl, kind, els, ctx.meta));
      else out.push(...els);
    }
    i = j;
  }
  return out;
}

/** A fresh copy of a run container around new content, keeping its own properties. */
function containerShell(c: XmlElement, content: XmlElement[], fromOriginal: boolean, ctx: Ctx): XmlElement {
  const isSdt = localName(c.name) === "sdt";
  const inner = isSdt ? c.children.find((k) => localName(k.name) === "sdtContent") : undefined;
  const shell: XmlElement = {
    name: c.name,
    attrs: { ...c.attrs },
    // Property children only (sdtPr, smartTagPr, ...): the content is new.
    children: c.children.filter((k) => /Pr$/.test(localName(k.name))).map(cloneXml),
    text: "",
  };
  // A container copied from the original points into ITS package (a
  // hyperlink's r:id); adopt it before the already-adopted content goes in.
  if (fromOriginal) adoptFromOriginal(shell, ctx);
  if (isSdt) shell.children.push({ name: inner?.name ?? c.name.replace(/sdt$/, "sdtContent"), attrs: { ...(inner?.attrs ?? {}) }, children: content, text: "" });
  else shell.children.push(...content);
  return shell;
}

/** The element with its properties child removed, for a shape comparison. */
function stripProps(el: XmlElement): XmlElement {
  return { ...el, children: el.children.filter((c) => localName(c.name) !== "pPr") };
}

/**
 * Emit an equal span from the REVISED side (so accepting yields exactly the
 * revised text) and record a run-format revision wherever the original wore
 * different run properties over the same words.
 */
function emitEqualSpan(
  oTokens: readonly ContentToken[],
  rTokens: readonly ContentToken[],
  seg: Segment,
  ctx: Ctx,
  counterpart: Map<XmlElement, XmlElement>,
): Placed[] {
  // Matched words are emitted from the REVISED runs, split wherever EITHER
  // side changes run, so each stretch compares with the original's
  // properties at exactly that place. A word can now span runs (see
  // readParagraph), and taking the formatting of only its first piece put the
  // colon's bold of "Purpose:" onto the plain space after it.
  type Unit = { token: ContentToken; run: XmlElement; t?: XmlElement; text: string; atom?: XmlElement; oAtom?: XmlElement; was: XmlElement };
  const units: Unit[] = [];
  const pieceAt = (pieces: NonNullable<ContentToken["pieces"]>, offset: number) => {
    let base = 0;
    for (const p of pieces) {
      if (offset < base + p.text.length) return p;
      base += p.text.length;
    }
    return pieces[pieces.length - 1];
  };
  for (let k = 0; k < seg.bEnd - seg.bStart; k++) {
    const o = oTokens[seg.aStart + k];
    const r = rTokens[seg.bStart + k];
    // A word matched inside a container on both sides pairs the containers.
    for (let d = 0; d < Math.min(o.path.length, r.path.length); d++) {
      if (localName(o.path[d].name) === localName(r.path[d].name)) counterpart.set(o.path[d], r.path[d]);
    }
    if (r.atom || !r.pieces || !o.pieces || o.text !== r.text) {
      // An atom, or a word matched by key with different whitespace (the
      // whitespace option off): token by token, as one stretch.
      if (r.atom) units.push({ token: r, run: r.run, text: "", atom: r.atom, oAtom: o.atom, was: o.run });
      else for (const p of r.pieces ?? []) units.push({ token: r, run: p.run, t: p.t, text: p.text, was: o.run });
      continue;
    }
    const cuts = new Set<number>([0, r.text.length]);
    for (const pieces of [o.pieces, r.pieces]) {
      let base = 0;
      for (const p of pieces) cuts.add((base += p.text.length));
    }
    const sorted = [...cuts].sort((x, y) => x - y);
    for (let c = 0; c + 1 < sorted.length; c++) {
      const from = sorted[c];
      const to = sorted[c + 1];
      if (from === to) continue;
      const rp = pieceAt(r.pieces, from);
      units.push({ token: r, run: rp.run, t: rp.t, text: r.text.slice(from, to), was: pieceAt(o.pieces, from).run });
    }
  }

  const out: Placed[] = [];
  let open: { run: XmlElement; was: string; el: XmlElement; prev: XmlElement | undefined } | null = null;
  let text: { t: XmlElement; el: XmlElement } | null = null;
  const close = (): void => {
    if (open && ctx.formatting && open.was !== runPropsKey(open.run)) swapRunProps(open.el, open.prev, ctx);
    open = null;
    text = null;
  };
  let lastToken: ContentToken | null = null;
  for (const u of units) {
    if (u.token !== lastToken && u.token.lead) {
      close();
      out.push(...u.token.lead.map((el) => ({ el: cloneXml(el), path: u.token.path })));
    }
    lastToken = u.token;
    const was = runPropsKey(u.was);
    if (!open || open.run !== u.run || open.was !== was) {
      close();
      const rPr = u.run.children.find((c) => localName(c.name) === "rPr");
      const el: XmlElement = { name: u.run.name, attrs: { ...u.run.attrs }, children: rPr ? [cloneXml(rPr)] : [], text: "" };
      out.push({ el, path: u.token.path });
      open = { run: u.run, was, el, prev: child(u.was, "rPr") };
    }
    if (u.atom) {
      const atom = cloneXml(u.atom);
      if (u.oAtom && !sameXml(u.oAtom, u.atom)) compareTextBoxes(u.oAtom, atom, ctx);
      open.el.children.push(atom);
      text = null;
      continue;
    }
    if (text && text.t === u.t) {
      text.el.text += u.text;
      continue;
    }
    const te: XmlElement = { name: u.t!.name, attrs: { ...u.t!.attrs, "xml:space": "preserve" }, children: [], text: u.text };
    open.el.children.push(te);
    text = { t: u.t!, el: te };
  }
  close();
  return out;
}

/**
 * Formatting-only change with the run partition intact on both sides: record
 * the difference on the revised runs where they stand, rather than rebuilding
 * the paragraph and splitting runs for no reason. False when the partitions
 * differ, so the caller can rebuild instead.
 */
function alignRunPropsInPlace(oEl: XmlElement, rEl: XmlElement, ctx: Ctx): boolean {
  const oRuns = oEl.children.filter((c) => localName(c.name) === "r");
  const rRuns = rEl.children.filter((c) => localName(c.name) === "r");
  if (oRuns.length !== rRuns.length) return false;
  for (let i = 0; i < rRuns.length; i++) {
    if (paragraphText(oRuns[i]) !== paragraphText(rRuns[i])) return false;
  }
  for (let i = 0; i < rRuns.length; i++) swapRunProps(rRuns[i], child(oRuns[i], "rPr"), ctx);
  return true;
}

// ---------------------------------------------------------------------------
// Inner (word-level) diff
// ---------------------------------------------------------------------------

interface Segment {
  equal: boolean;
  aStart: number;
  aEnd: number;
  bStart: number;
  bEnd: number;
}

/**
 * Word-level diff of two token sequences, then the pass that makes the result
 * readable.
 *
 * The raw LCS is confetti on a rewritten sentence: it matches "the", "of" and
 * "a" between the old wording and the new one and returns a dozen interleaved
 * fragments. So any matched run shorter than COMPARE_TUNING.coalesceWindow
 * sitting BETWEEN two changes is absorbed into the change around it, and
 * adjacent changes are then merged into one region — which comes out as one
 * contiguous strike followed by one contiguous insertion, the way a human
 * means "I rewrote that sentence".
 */
function wordSegments(a: readonly ContentToken[], b: readonly ContentToken[]): Segment[] {
  const raw: Segment[] = [];
  let ai = 0;
  let bi = 0;
  for (const part of diffArrays(
    a.map((t) => t.key),
    b.map((t) => t.key),
  )) {
    const n = part.value.length;
    if (part.added) {
      raw.push({ equal: false, aStart: ai, aEnd: ai, bStart: bi, bEnd: bi + n });
      bi += n;
    } else if (part.removed) {
      raw.push({ equal: false, aStart: ai, aEnd: ai + n, bStart: bi, bEnd: bi });
      ai += n;
    } else {
      raw.push({ equal: true, aStart: ai, aEnd: ai + n, bStart: bi, bEnd: bi + n });
      ai += n;
      bi += n;
    }
  }

  // Merge first, so a delete and the insert replacing it are ONE change and
  // the coalescing rule below can weigh a matched run against the whole edit
  // rather than against half of it.
  let segments = mergeChanges(raw);
  for (;;) {
    let absorbed = false;
    for (let i = 1; i + 1 < segments.length; i++) {
      const seg = segments[i];
      if (!seg.equal) continue;
      const len = words(a, seg.aStart, seg.aEnd);
      const beside = Math.min(changeSize(segments[i - 1], a, b), changeSize(segments[i + 1], a, b));
      if (len < COMPARE_TUNING.coalesceWindow || len * COMPARE_TUNING.coalesceRatio < beside) {
        seg.equal = false;
        absorbed = true;
      }
    }
    if (!absorbed) break;
    segments = mergeChanges(segments);
  }
  return segments.filter((s) => s.aEnd > s.aStart || s.bEnd > s.bStart);
}

/** Collapse consecutive changes into one region; segments then alternate. */
function mergeChanges(segments: readonly Segment[]): Segment[] {
  const out: Segment[] = [];
  for (const seg of segments) {
    const prev = out[out.length - 1];
    if (prev && !prev.equal && !seg.equal) {
      prev.aEnd = seg.aEnd;
      prev.bEnd = seg.bEnd;
      continue;
    }
    out.push({ ...seg });
  }
  return out;
}

/** How much a change region touches: the larger of its two sides. */
/**
 * How big a change is, in tokens that are not whitespace: a space between two
 * words is a token of its own, and counting it would make every stretch look
 * twice its size to the coalescing rule calibrated on words.
 */
function changeSize(seg: Segment, a: readonly ContentToken[], b: readonly ContentToken[]): number {
  return Math.max(words(a, seg.aStart, seg.aEnd), words(b, seg.bStart, seg.bEnd));
}

function words(tokens: readonly ContentToken[], from: number, to: number): number {
  let n = 0;
  for (let i = from; i < to; i++) if (tokens[i].atom || !/^\s+$/.test(tokens[i].text)) n++;
  return n;
}

// ---------------------------------------------------------------------------
// Tables
// ---------------------------------------------------------------------------

function compareTable(oTbl: XmlElement, rTbl: XmlElement, ctx: Ctx): void {
  if (sameXml(oTbl, rTbl)) return;
  compareTableProps(oTbl, rTbl, ctx);
  compareRows(oTbl, rTbl, ctx);
}

/** A matched table pair's own properties and grid, before its rows are compared. */
export function compareTableProps(oTbl: XmlElement, rTbl: XmlElement, ctx: Ctx): void {
  if (!ctx.formatting) return;
  swapProps(rTbl, "tblPr", recordTableFormatChange, child(oTbl, "tblPr"), ctx);
  alignTableGrid(oTbl, rTbl, ctx);
}

/** Row-level children: rows, and content controls or custom XML holding rows. */
export const ROW_LEVEL = ["tr", "sdt", "customXml"];

/**
 * Align and diff the rows of a table, or of a row-level content control (a
 * repeating section) inside one. Such a control was once set aside as a
 * non-row child: moved to the top of the table, its rows never compared.
 */
function compareRows(oParent: XmlElement, rParent: XmlElement, ctx: Ctx): void {
  const oRows = oParent.children.filter((c) => ROW_LEVEL.includes(localName(c.name)));
  const rRows = rParent.children.filter((c) => ROW_LEVEL.includes(localName(c.name)));
  const anchors = histogramMatch(oRows.map(rowFingerprint), rRows.map(rowFingerprint));

  const rows: XmlElement[] = [];
  let oAt = 0;
  let rAt = 0;
  for (const anchor of [...anchors, { a: oRows.length, b: rRows.length }]) {
    // Rows inside a gap pair up by similarity, exactly like paragraphs, so a
    // row whose one cell changed is an edited row and not a replaced one.
    const pairings = resolveGap(
      treeSide(oRows, ctx.original, true, ctx),
      treeSide(rRows, ctx.doc, false, ctx),
      oAt,
      anchor.a,
      rAt,
      anchor.b,
    ).filter((p) => p.kind === "pair");
    let oi = oAt;
    let ri = rAt;
    for (const pairing of [...pairings, { kind: "pair" as const, o: anchor.a, r: anchor.b, other: -1 }]) {
      for (let i = oi; i < pairing.o; i++) rows.push(deletedRow(oRows[i], ctx));
      for (let j = ri; j < pairing.r; j++) rows.push(insertedRow(rRows[j], ctx));
      oi = pairing.o;
      ri = pairing.r;
      if (pairing.o >= anchor.a || pairing.r >= anchor.b) break;
      rows.push(...compareRow(oRows[pairing.o], rRows[pairing.r], ctx));
      oi = pairing.o + 1;
      ri = pairing.r + 1;
    }
    if (anchor.a < oRows.length && anchor.b < rRows.length) {
      rows.push(...compareRow(oRows[anchor.a], rRows[anchor.b], ctx));
    }
    oAt = anchor.a + 1;
    rAt = anchor.b + 1;
  }

  rParent.children = [...rParent.children.filter((c) => !ROW_LEVEL.includes(localName(c.name))), ...rows];
}

export function rowFingerprint(tr: XmlElement): string {
  if (localName(tr.name) !== "tr") return `X\u0000${serializeXml(tr)}`;
  const cells = tr.children.filter((c) => localName(c.name) === "tc");
  return `${cells.length}\u0000${normalizeText(paragraphText(tr))}`;
}

/** The rows a row-level element holds: itself, or those inside its container. */
function rowsIn(el: XmlElement): XmlElement[] {
  if (localName(el.name) === "tr") return [el];
  const inner = localName(el.name) === "sdt" ? child(el, "sdtContent") : el;
  return (inner?.children ?? []).filter((c) => ROW_LEVEL.includes(localName(c.name))).flatMap(rowsIn);
}

export function deletedRow(el: XmlElement, ctx: Ctx): XmlElement {
  const clone = cloneXml(el);
  for (const tr of rowsIn(clone)) {
    markRowRevision(tr, "del", ctx.meta);
    wrapRowContent(tr, "del", ctx);
  }
  return clone;
}

export function insertedRow(el: XmlElement, ctx: Ctx): XmlElement {
  for (const tr of rowsIn(el)) {
    markRowRevision(tr, "ins", ctx.meta);
    wrapRowContent(tr, "ins", ctx);
  }
  return el;
}

/** A matched row pair, or — when the cell counts differ — two coarse rows. */
export function compareRow(oTr: XmlElement, rTr: XmlElement, ctx: Ctx): XmlElement[] {
  if (localName(oTr.name) !== "tr") {
    const inner = (el: XmlElement) => (localName(el.name) === "sdt" ? child(el, "sdtContent") : el);
    const o = inner(oTr);
    const r = inner(rTr);
    if (o && r) compareRows(o, r, ctx);
    return [rTr];
  }
  const oCells = oTr.children.filter((c) => localName(c.name) === "tc");
  const rCells = rTr.children.filter((c) => localName(c.name) === "tc");
  if (oCells.length !== rCells.length) {
    ctx.note({
      code: "coarse-row",
      detail: `a row's cell count changed (${oCells.length} → ${rCells.length}); struck and reinserted whole`,
    });
    return [deletedRow(oTr, ctx), insertedRow(rTr, ctx)];
  }
  if (ctx.formatting) swapProps(rTr, "trPr", recordRowFormatChange, child(oTr, "trPr"), ctx, ["ins", "del"]);
  for (let i = 0; i < rCells.length; i++) {
    if (ctx.formatting) swapProps(rCells[i], "tcPr", recordCellFormatChange, child(oCells[i], "tcPr"), ctx);
    const entries = alignBlocks(blockList(oCells[i], "tcPr"), blockList(rCells[i], "tcPr"), ctx);
    applyParagraphMarks(entries, ctx);
    const tcPr = child(rCells[i], "tcPr");
    rCells[i].children = [...(tcPr ? [tcPr] : []), ...entries.map((e) => e.el)];
  }
  return [rTr];
}

function alignTableGrid(oTbl: XmlElement, rTbl: XmlElement, ctx: Ctx): void {
  const oGrid = child(oTbl, "tblGrid");
  const rGrid = child(rTbl, "tblGrid");
  if (!oGrid || !rGrid || sameXml(oGrid, rGrid)) return;
  const live = rGrid.children.map(cloneXml);
  rGrid.children = oGrid.children.map(cloneXml);
  recordTableGridChange(rTbl, ctx.meta);
  const record = rGrid.children[rGrid.children.length - 1];
  rGrid.children = [...live, record];
}

// ---------------------------------------------------------------------------
// Property revisions
// ---------------------------------------------------------------------------

/** What a w:pPrChange cannot carry, so what stays live across a reject. */
const PARA_KEPT = ["rPr", "sectPr"];

/**
 * Make `pEl` wear `live` and remember `previous`, as one w:pPrChange.
 *
 * Accepting keeps the live properties and drops the record; rejecting puts the
 * recorded ones back. That is the whole trick behind the paragraph-mark work
 * below: the paragraph a reviewer ends up with must wear the properties of
 * whichever document they chose, and only a pPrChange can hold both answers.
 */
export function alignParagraphProps(
  pEl: XmlElement,
  live: XmlElement | undefined,
  previous: XmlElement | undefined,
  ctx: Ctx,
): void {
  const liveBase = baseChildren(live, PARA_KEPT);
  const prevBase = baseChildren(previous, PARA_KEPT);
  if (keyOf(liveBase) === keyOf(prevBase)) return;

  // The paragraph MARK's own run properties are outside CT_PPrBase, so a
  // pPrChange cannot carry them: whatever the paragraph wears now, it keeps
  // through both accept and reject. Say so when the two sides disagree
  // instead of quietly picking one.
  const liveMark = keyOf(children(live, "rPr"));
  const prevMark = keyOf(children(previous, "rPr"));
  if (liveMark !== prevMark) {
    ctx.note({
      code: "paragraph-mark-props",
      detail: "the paragraph mark's own run properties differ; w:pPrChange cannot record them, so the result keeps one side's",
    });
  }

  const w = prefixOf(pEl.name);
  let pPr = child(pEl, "pPr");
  if (!pPr) {
    pPr = { name: `${w}pPr`, attrs: {}, children: [], text: "" };
    pEl.children.unshift(pPr);
  }
  const kept = pPr.children.filter((c) => PARA_KEPT.includes(localName(c.name)));
  pPr.children = prevBase.map(cloneXml);
  recordParagraphFormatChange(pEl, ctx.meta);
  const record = pPr.children[pPr.children.length - 1];
  pPr.children = [...liveBase.map(cloneXml), ...kept, record];
}

/**
 * The generic form for w:rPrChange, w:tblPrChange, w:trPrChange and
 * w:tcPrChange. `kept` names the children the record's payload type cannot
 * carry — for a row, its own structural w:ins / w:del — which therefore stay
 * live and are not part of the comparison either.
 */
function swapProps(
  owner: XmlElement,
  propsName: string,
  record: (owner: XmlElement, meta: RevisionMeta) => void,
  previous: XmlElement | undefined,
  ctx: Ctx,
  kept: readonly string[] = [],
): void {
  const changeName = propsName + "Change";
  const skip = (c: XmlElement): boolean => {
    const ln = localName(c.name);
    return ln === changeName || kept.includes(ln);
  };
  const liveEl = child(owner, propsName);
  const live = (liveEl?.children ?? []).filter((c) => !skip(c));
  const prev = (previous?.children ?? []).filter((c) => !skip(c));
  if (keyOf(live) === keyOf(prev)) return;

  const w = prefixOf(owner.name);
  let props = liveEl;
  if (!props) {
    props = { name: `${w}${propsName}`, attrs: {}, children: [], text: "" };
    owner.children.unshift(props);
  }
  const liveClone = live.map(cloneXml);
  const keptEls = props.children.filter((c) => kept.includes(localName(c.name)));
  props.children = prev.map(cloneXml);
  record(owner, ctx.meta);
  const changeEl = props.children[props.children.length - 1];
  props.children = [...liveClone, ...keptEls, changeEl];
}

function swapRunProps(rEl: XmlElement, previous: XmlElement | undefined, ctx: Ctx): void {
  swapProps(rEl, "rPr", recordRunFormatChange, previous, ctx);
}

function baseChildren(props: XmlElement | undefined, kept: string[]): XmlElement[] {
  return (props?.children ?? []).filter((c) => {
    const ln = localName(c.name);
    return ln !== "pPrChange" && !kept.includes(ln);
  });
}

function children(props: XmlElement | undefined, name: string): XmlElement[] {
  return (props?.children ?? []).filter((c) => localName(c.name) === name);
}

function keyOf(els: readonly XmlElement[]): string {
  return els.map((el) => serializeXml(el)).join("");
}

function prefixOf(name: string): string {
  return name.includes(":") ? name.slice(0, name.indexOf(":") + 1) : "";
}

// ---------------------------------------------------------------------------
// Paragraph marks
// ---------------------------------------------------------------------------

/**
 * Record the paragraph MARKS — the pilcrows — that differ between the two
 * documents. Getting this wrong is what turns an otherwise correct compare
 * into one that cannot be accepted back into the revised document.
 *
 * The rule is: **mark the pilcrow of the paragraph BEFORE each added or
 * removed one.**
 *
 * Accepting a deleted mark joins the paragraph with the one after it and the
 * FIRST paragraph's properties survive (Word's rule, and what this engine's
 * own untracked Backspace-merge does). So striking the mark of the paragraph
 * *preceding* a removed one leaves the surviving paragraph wearing the
 * properties of a paragraph both documents agree about — no compensation
 * needed at all. Striking the removed paragraph's OWN mark would leave the
 * survivor wearing the deleted paragraph's properties, which is wrong in
 * exactly the case that matters (a deleted heading before body text).
 *
 * It matches the editor too: Backspace at the start of a paragraph calls
 * markParagraphGlyph on the PREVIOUS paragraph, and a trailing paragraph
 * simply leaves the last mark alone, which is why deleting the last paragraph
 * of a document works.
 *
 * A change run at the very START of a container has no preceding paragraph, so
 * there the marks go on the paragraphs themselves and the head of each
 * same-class chain takes a w:pPrChange to carry the other document's answer.
 */
function applyParagraphMarks(entries: Entry[], ctx: Ctx): void {
  const infos = entries.map((e) => ({ cls: e.cls, paragraph: isParagraph(e.el), endsSection: endsSection(e.el) }));
  paragraphMarkSteps(
    infos.length,
    (i) => infos[i],
    (step) => applyMarkStep(step, (i) => entries[i], (i) => child(entries[i].el, "pPr"), ctx),
  );
}

/** What the mark pass needs to know of one entry. */
export interface MarkInfo {
  cls: BlockClass;
  paragraph: boolean;
  /** A paragraph that ends a section: its mark is the break. */
  endsSection: boolean;
}

/**
 * One step of the mark pass: mark entry `at`'s pilcrow; let the head of a
 * same-class chain wear either document's paragraph properties, the other
 * one being its `follower`'s; or report a final mark that cannot be tracked.
 */
export type MarkStep =
  | { kind: "glyph"; at: number; glyph: "ins" | "del" }
  | { kind: "props"; at: number; follower: number }
  | { kind: "untracked" };

/** The mark pass over `n` entries, step by step, in the order applyParagraphMarks applies them. */
export function paragraphMarkSteps(n: number, info: (i: number) => MarkInfo, step: (s: MarkStep) => void): void {
  let i = 0;
  while (i < n) {
    if (info(i).cls === "kept") {
      i++;
      continue;
    }
    let end = i;
    while (end + 1 < n && info(end + 1).cls !== "kept") end++;
    // Tables carry their own row revisions and no pilcrow, so the paragraphs
    // between them are marked stretch by stretch: a stretch after a table has
    // no paragraph before it to take a mark from, exactly like one opening the
    // container, and marks its own (a paragraph removed right after a removed
    // table used to get no mark at all, and survived accept-all empty).
    let at = i;
    while (at <= end) {
      if (!info(at).paragraph) {
        at++;
        continue;
      }
      let last = at;
      while (last + 1 <= end && info(last + 1).paragraph) last++;
      paragraphRunSteps(n, info, at, last, step);
      at = last + 1;
    }
    i = end + 1;
  }
}

function paragraphRunSteps(n: number, info: (i: number) => MarkInfo, start: number, end: number, step: (s: MarkStep) => void): void {
  // Shifting moves each mark one paragraph up, and a section break lives in
  // its paragraph's mark: shifted, a kept break would be lent to the stretch
  // after it, and an added break would read as already there. So a stretch
  // next to or holding a break marks its own paragraphs.
  let shifted = start > 0 && info(start - 1).paragraph && !info(start - 1).endsSection;
  for (let k = start; shifted && k <= end; k++) if (info(k).endsSection) shifted = false;
  if (shifted) {
    for (let k = start; k <= end; k++) {
      if (!info(k).paragraph || !info(k - 1).paragraph) continue;
      step({ kind: "glyph", at: k - 1, glyph: glyphKind(info(k).cls) });
    }
    return;
  }

  const hasFollower = end + 1 < n;
  for (let k = start; k <= end; k++) {
    if (!info(k).paragraph) continue;
    if (k === end && !hasFollower) {
      step({ kind: "untracked" });
      continue;
    }
    step({ kind: "glyph", at: k, glyph: glyphKind(info(k).cls) });
  }

  // Self-marked paragraphs join with the FOLLOWING one when reviewed, so the
  // head of each same-class chain must be able to wear either document's
  // paragraph properties.
  let at = start;
  while (at <= end) {
    let last = at;
    while (last + 1 <= end && info(last + 1).cls === info(at).cls) last++;
    if (last + 1 < n && info(at).paragraph && info(last + 1).paragraph) step({ kind: "props", at, follower: last + 1 });
    at = last + 1;
  }
}

/**
 * Apply one mark step. `entry` gives an entry by index; `pPrOf` gives an
 * entry's paragraph properties as they stand once it is emitted.
 */
export function applyMarkStep(step: MarkStep, entry: (i: number) => Entry, pPrOf: (i: number) => XmlElement | undefined, ctx: Ctx): void {
  if (step.kind === "untracked") {
    ctx.note({
      code: "no-common-paragraph",
      detail: "the documents share no paragraph, so the final paragraph mark difference is left untracked",
    });
  } else if (step.kind === "glyph") {
    markParagraphGlyph(entry(step.at).el, step.glyph, ctx.meta);
  } else {
    const head = entry(step.at);
    const own = child(head.el, "pPr");
    const other = pPrOf(step.follower);
    if (head.cls === "deleted") alignParagraphProps(head.el, other, own, ctx);
    else alignParagraphProps(head.el, own, other, ctx);
  }
}

function endsSection(el: XmlElement): boolean {
  return isParagraph(el) && !!child(child(el, "pPr"), "sectPr");
}

export function glyphKind(cls: BlockClass): "ins" | "del" {
  return cls === "inserted" ? "ins" : "del";
}

export function isParagraph(el: XmlElement): boolean {
  return localName(el.name) === "p";
}
