import { diffArrays } from "diff";
import { DocxDocument } from "../../docx.js";
import { XmlElement, child, cloneXml, localName, serializeXml } from "../../xml.js";
import {
  RevisionMeta,
  markDeletedText,
  markParagraphGlyph,
  markRowRevision,
  recordCellFormatChange,
  recordParagraphFormatChange,
  recordRowFormatChange,
  recordRunFormatChange,
  recordTableFormatChange,
  recordTableGridChange,
  revisionWrapper,
} from "../suggest.js";
import { ContentToken, isPassenger, paragraphText, readParagraph, runPropsKey, sliceContent } from "./content.js";
import { histogramMatch } from "./histogram.js";
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
 * **Headers, footers, footnotes, endnotes, comments, styles and numbering —
 * NOT compared.** The result carries `revised`'s. `onNote` reports when
 * `original`'s differ, so a reviewer is never told "no change" about something
 * that was never looked at.
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
    | "unresolved-relationship";
  detail: string;
}

export interface CompareOptions {
  /** Who the revisions are attributed to. */
  author: string;
  /** ISO 8601 timestamp for every w:date written. */
  date: string;
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
}

interface Ctx {
  doc: DocxDocument;
  /** The document deleted content is cloned OUT of; see adoptFromOriginal. */
  original: DocxDocument;
  meta: RevisionMeta;
  formatting: boolean;
  whitespace: boolean;
  note: (note: CompareNote) => void;
}

/** Where a block ended up, which is what decides the paragraph-mark revisions. */
type BlockClass = "kept" | "inserted" | "deleted";

interface Entry {
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
  const merged = DocxDocument.load(revised.save());
  const ctx: Ctx = {
    doc: merged,
    original,
    meta: {
      author: options.author,
      date: options.date,
      nextId: () => merged.nextRevisionId(),
    },
    formatting: options.formatting !== false,
    whitespace: options.whitespace !== false,
    note: options.onNote ?? (() => {}),
  };

  const originalBody = child(original.docRoot, "body");
  const mergedBody = child(merged.docRoot, "body");
  if (!originalBody || !mergedBody) throw new Error("document.xml has no w:body");

  reportUncomparedParts(original, revised, ctx);

  const tail = mergedBody.children.filter((c) => localName(c.name) === "sectPr");
  const entries = alignBlocks(blockList(originalBody, "sectPr"), blockList(mergedBody, "sectPr"), ctx);
  applyParagraphMarks(entries, ctx);
  mergedBody.children = [...entries.map((e) => e.el), ...tail];

  merged.refresh();
  return merged;
}

function blockList(container: XmlElement, ...skip: string[]): XmlElement[] {
  return container.children.filter((c) => !skip.includes(localName(c.name)));
}

/** Say so when a part outside the body differs, rather than imply it matched. */
function reportUncomparedParts(original: DocxDocument, revised: DocxDocument, ctx: Ctx): void {
  const parts: [string, XmlElement | null][] = [
    ["styles.xml", original.stylesTree()],
    ["numbering.xml", original.numberingTree()],
    ["footnotes.xml", original.footnotesTree()],
    ["endnotes.xml", original.endnotesTree()],
  ];
  const revisedTrees: Record<string, XmlElement | null> = {
    "styles.xml": revised.stylesTree(),
    "numbering.xml": revised.numberingTree(),
    "footnotes.xml": revised.footnotesTree(),
    "endnotes.xml": revised.endnotesTree(),
  };
  for (const [name, tree] of parts) {
    const other = revisedTrees[name];
    const a = tree ? serializeXml(tree) : "";
    const b = other ? serializeXml(other) : "";
    if (a !== b) {
      ctx.note({ code: "part-not-compared", detail: `${name} differs; the result carries the revised document's` });
    }
  }
  const hf = (doc: DocxDocument): string =>
    [...doc.headerRoots(), ...doc.footerRoots()].map((root) => serializeXml(root)).join("");
  if (hf(original) !== hf(revised)) {
    ctx.note({ code: "part-not-compared", detail: "headers/footers differ; the result carries the revised document's" });
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
  const anchors = histogramMatch(oBlocks.map(blockFingerprint), rBlocks.map(blockFingerprint));
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
function blockFingerprint(el: XmlElement): string {
  const ln = localName(el.name);
  if (ln === "p") {
    const pPr = child(el, "pPr");
    const style = child(pPr, "pStyle")?.attrs["w:val"] ?? "";
    const numPr = child(pPr, "numPr");
    return `P\u0000${normalizeText(paragraphText(el))}\u0000${style}\u0000${numPr ? serializeXml(numPr) : ""}`;
  }
  if (ln === "tbl") return `T\u0000${normalizeText(paragraphText(el))}`;
  return `X\u0000${serializeXml(el)}`;
}

interface Pairing {
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
  const pairings = resolveGap(oBlocks, rBlocks, oLo, oHi, rLo, rHi, ctx, oLo === 0 && rLo === 0);
  let oAt = oLo;
  let rAt = rLo;
  for (const pairing of [...pairings, { kind: "pair" as const, o: oHi, r: rHi, other: -1 }]) {
    for (let i = oAt; i < pairing.o; i++) entries.push(emitDeleted(oBlocks[i], ctx));
    for (let j = rAt; j < pairing.r; j++) entries.push(emitInserted(rBlocks[j], ctx));
    oAt = pairing.o;
    rAt = pairing.r;
    if (pairing.o >= oHi || pairing.r >= rHi) break;
    if (pairing.kind === "split") {
      // One original paragraph became two. Their content already IS the
      // original's, word for word and property for property, so there is no
      // text revision to make — only the new pilcrow between them, which the
      // mark pass records against the first paragraph.
      alignParagraphProps(rBlocks[pairing.r], child(rBlocks[pairing.r], "pPr"), child(oBlocks[pairing.o], "pPr"), ctx);
      entries.push({ el: rBlocks[pairing.r], cls: "kept" });
      entries.push({ el: rBlocks[pairing.other], cls: "inserted" });
      oAt = pairing.o + 1;
      rAt = pairing.other + 1;
    } else if (pairing.kind === "merge") {
      // Two originals became one: both are emitted, and the pilcrow between
      // them is struck. Their content is already what the revised paragraph
      // holds, so there is no text revision to make.
      const first = cloneXml(oBlocks[pairing.o]);
      adoptFromOriginal(first, ctx);
      alignParagraphProps(first, child(rBlocks[pairing.r], "pPr"), child(oBlocks[pairing.o], "pPr"), ctx);
      entries.push({ el: first, cls: "kept" });
      const second = cloneXml(oBlocks[pairing.other]);
      adoptFromOriginal(second, ctx);
      entries.push({ el: second, cls: "deleted" });
      oAt = pairing.other + 1;
      rAt = pairing.r + 1;
    } else {
      entries.push(emitMatched(oBlocks[pairing.o], rBlocks[pairing.r], ctx));
      oAt = pairing.o + 1;
      rAt = pairing.r + 1;
    }
  }
}

function resolveGap(
  oBlocks: XmlElement[],
  rBlocks: XmlElement[],
  oLo: number,
  oHi: number,
  rLo: number,
  rHi: number,
  ctx: Ctx,
  needsAnchor = false,
): Pairing[] {
  const out: Pairing[] = [];
  const oUsed = new Set<number>();
  const rUsed = new Set<number>();

  // Splits and merges first: they are exact (the content keys concatenate),
  // and they are the shape a similarity pass would otherwise turn into
  // delete-one-insert-two, which reads as a rewrite of something nobody
  // rewrote.
  const crosses = (o: number, r: number): boolean => out.some((p) => (o < p.o) !== (r < p.r));

  for (let i = oLo; i < oHi; i++) {
    for (let j = rLo; j + 1 < rHi; j++) {
      if (oUsed.has(i) || rUsed.has(j) || rUsed.has(j + 1) || crosses(i, j)) continue;
      if (concatMatches([oBlocks[i]], [rBlocks[j], rBlocks[j + 1]], ctx)) {
        out.push({ kind: "split", o: i, r: j, other: j + 1 });
        oUsed.add(i);
        rUsed.add(j);
        rUsed.add(j + 1);
      }
    }
  }
  for (let i = oLo; i + 1 < oHi; i++) {
    for (let j = rLo; j < rHi; j++) {
      if (oUsed.has(i) || oUsed.has(i + 1) || rUsed.has(j) || crosses(i, j)) continue;
      if (concatMatches([oBlocks[i], oBlocks[i + 1]], [rBlocks[j]], ctx)) {
        out.push({ kind: "merge", o: i, r: j, other: i + 1 });
        oUsed.add(i);
        oUsed.add(i + 1);
        rUsed.add(j);
      }
    }
  }

  // Then similarity, best score first, over the gap's cross-product only.
  const scored: { o: number; r: number; score: number }[] = [];
  for (let i = oLo; i < oHi; i++) {
    if (oUsed.has(i)) continue;
    for (let j = rLo; j < rHi; j++) {
      if (rUsed.has(j)) continue;
      if (localName(oBlocks[i].name) !== localName(rBlocks[j].name)) continue;
      const a = paragraphText(oBlocks[i]);
      const b = paragraphText(rBlocks[j]);
      const score = similarity(a, b);
      if (score >= similarityFloor(a, b)) scored.push({ o: i, r: j, score });
    }
  }
  scored.sort((x, y) => y.score - x.score || x.o - y.o);
  for (const cand of scored) {
    if (oUsed.has(cand.o) || rUsed.has(cand.r)) continue;
    // Pairings must stay monotone: a crossing pair would put a deletion after
    // the insertion that replaced it and read as two unrelated edits.
    if (crosses(cand.o, cand.r)) continue;
    out.push({ kind: "pair", o: cand.o, r: cand.r, other: -1 });
    oUsed.add(cand.o);
    rUsed.add(cand.r);
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
        if (rUsed.has(j) || localName(oBlocks[i].name) !== localName(rBlocks[j].name)) continue;
        out.push({ kind: "pair", o: i, r: j, other: -1 });
        break;
      }
    }
  }
  return out.sort((a, b) => a.o - b.o);
}

/**
 * True when the two sides carry exactly the same content — same words, same
 * atoms, same run properties — once concatenated. That exactness is what makes
 * a split or a merge emit ONE paragraph-mark revision and no text revision,
 * and what makes rejecting it restore the original byte for byte.
 */
function concatMatches(oEls: XmlElement[], rEls: XmlElement[], ctx: Ctx): boolean {
  const key = (els: XmlElement[]): string | null => {
    const parts: string[] = [];
    for (const el of els) {
      if (localName(el.name) !== "p") return null;
      const content = readParagraph(el, { whitespace: ctx.whitespace });
      if (!content.inlineDiffable) return null;
      for (const token of content.tokens) parts.push(token.key + "\u0002" + runPropsKey(token.run));
    }
    return parts.join("\u0001");
  };
  const a = key(oEls);
  const b = key(rEls);
  return a !== null && b !== null && a === b;
}

// ---------------------------------------------------------------------------
// Emission
// ---------------------------------------------------------------------------

/**
 * A matched pair. Both sides are always the same kind of block: fingerprints
 * are namespaced by kind, and the gap pass refuses to pair across kinds.
 */
function emitMatched(oEl: XmlElement, rEl: XmlElement, ctx: Ctx): Entry {
  const ln = localName(oEl.name);
  if (ln === "p") comparePair(oEl, rEl, ctx);
  else if (ln === "tbl") compareTable(oEl, rEl, ctx);
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
 * points at a PART — an image, a chart, an embedded object — and carrying the
 * part across is out of this increment's scope, so the reference is dropped
 * and reported rather than left dangling: struck text with no picture beats a
 * document that will not open.
 */
function adoptFromOriginal(root: XmlElement, ctx: Ctx): void {
  const mine = ctx.doc.documentRels;
  const theirs = ctx.original.documentRels;
  const remapped = new Map<string, string | null>();
  const walk = (el: XmlElement): void => {
    for (const key of Object.keys(el.attrs)) {
      if (!key.startsWith("r:") || !REL_ATTRS.includes(localName(key))) continue;
      const id = el.attrs[key];
      const source = theirs.get(id);
      if (!source) continue;
      const here = mine.get(id);
      if (here && here.target === source.target && here.type === source.type) continue;
      let replacement = remapped.get(id);
      if (replacement === undefined) {
        replacement = source.external ? ctx.doc.addHyperlinkRel(source.target) : null;
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

function emitDeleted(el: XmlElement, ctx: Ctx): Entry {
  const clone = cloneXml(el);
  adoptFromOriginal(clone, ctx);
  const ln = localName(clone.name);
  if (ln === "p") {
    wrapContainerContent(clone, "del", ctx);
  } else if (ln === "tbl") {
    for (const row of clone.children.filter((c) => localName(c.name) === "tr")) {
      markRowRevision(row, "del", ctx.meta);
      wrapRowContent(row, "del", ctx);
    }
  } else if (!isPassenger(ln)) {
    ctx.note({ code: "untracked-block", detail: `a <${clone.name}> block was removed and is not expressible as a revision` });
  }
  return { el: clone, cls: "deleted" };
}

function emitInserted(el: XmlElement, ctx: Ctx): Entry {
  const ln = localName(el.name);
  if (ln === "p") {
    wrapContainerContent(el, "ins", ctx);
  } else if (ln === "tbl") {
    for (const row of el.children.filter((c) => localName(c.name) === "tr")) {
      markRowRevision(row, "ins", ctx.meta);
      wrapRowContent(row, "ins", ctx);
    }
  } else if (!isPassenger(ln)) {
    ctx.note({ code: "untracked-block", detail: `a <${el.name}> block was added and is not expressible as a revision` });
  }
  return { el, cls: "inserted" };
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
function wrapContainerContent(container: XmlElement, kind: "ins" | "del", ctx: Ctx): void {
  const NESTED = ["hyperlink", "smartTag", "dir", "bdo"];
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
    if (ln === "r") {
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
  if (serializeXml(oEl) === serializeXml(rEl)) return;

  if (ctx.formatting) alignParagraphProps(rEl, child(rEl, "pPr"), child(oEl, "pPr"), ctx);

  const oc = readParagraph(oEl, { whitespace: ctx.whitespace });
  const rc = readParagraph(rEl, { whitespace: ctx.whitespace });
  if (!oc.inlineDiffable || !rc.inlineDiffable) {
    if (oc.text !== rc.text || serializeXml(stripProps(oEl)) !== serializeXml(stripProps(rEl))) {
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
        detail: "a paragraph holding a hyperlink, content control, smart tag or inline math changed; struck and reinserted whole",
      });
    }
    return;
  }

  const segments = wordSegments(oc.tokens, rc.tokens);
  const changed = segments.some((s) => !s.equal);
  if (!changed) {
    // Same words in the same order: any remaining difference is formatting.
    if (ctx.formatting) alignRunPropsInPlace(oEl, rEl, ctx);
    return;
  }

  const pPr = child(rEl, "pPr");
  const out: XmlElement[] = [];
  for (const seg of segments) {
    if (seg.equal) {
      out.push(...emitEqualSpan(oc.tokens, rc.tokens, seg, ctx));
      continue;
    }
    if (seg.aEnd > seg.aStart) {
      const struck = sliceContent(oc.tokens, seg.aStart, seg.aEnd, true);
      for (const run of struck) adoptFromOriginal(run, ctx);
      out.push(revisionWrapper(rEl, "del", struck, ctx.meta));
    }
    if (seg.bEnd > seg.bStart) {
      out.push(revisionWrapper(rEl, "ins", sliceContent(rc.tokens, seg.bStart, seg.bEnd, false), ctx.meta));
    }
  }
  out.push(...rc.tail.map(cloneXml));
  rEl.children = [...(pPr ? [pPr] : []), ...out];
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
): XmlElement[] {
  const out: XmlElement[] = [];
  let at = seg.bStart;
  while (at < seg.bEnd) {
    const run = rTokens[at].run;
    const oKey = runPropsKey(oTokens[seg.aStart + (at - seg.bStart)].run);
    let end = at + 1;
    while (
      end < seg.bEnd &&
      rTokens[end].run === run &&
      runPropsKey(oTokens[seg.aStart + (end - seg.bStart)].run) === oKey
    ) {
      end++;
    }
    const piece = sliceContent(rTokens, at, end, false);
    if (ctx.formatting && oKey !== runPropsKey(run)) {
      for (const el of piece) {
        if (localName(el.name) === "r") {
          swapRunProps(el, child(oTokens[seg.aStart + (at - seg.bStart)].run, "rPr"), ctx);
        }
      }
    }
    out.push(...piece);
    at = end;
  }
  return out;
}

/**
 * Formatting-only change with the run partition intact on both sides: record
 * the difference on the revised runs where they stand, rather than rebuilding
 * the paragraph and splitting runs for no reason.
 */
function alignRunPropsInPlace(oEl: XmlElement, rEl: XmlElement, ctx: Ctx): void {
  const oRuns = oEl.children.filter((c) => localName(c.name) === "r");
  const rRuns = rEl.children.filter((c) => localName(c.name) === "r");
  if (oRuns.length !== rRuns.length) return;
  for (let i = 0; i < rRuns.length; i++) {
    if (paragraphText(oRuns[i]) !== paragraphText(rRuns[i])) return;
  }
  for (let i = 0; i < rRuns.length; i++) swapRunProps(rRuns[i], child(oRuns[i], "rPr"), ctx);
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
      const len = seg.aEnd - seg.aStart;
      const beside = Math.min(changeSize(segments[i - 1]), changeSize(segments[i + 1]));
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
function changeSize(seg: Segment): number {
  return Math.max(seg.aEnd - seg.aStart, seg.bEnd - seg.bStart);
}

// ---------------------------------------------------------------------------
// Tables
// ---------------------------------------------------------------------------

function compareTable(oTbl: XmlElement, rTbl: XmlElement, ctx: Ctx): void {
  if (serializeXml(oTbl) === serializeXml(rTbl)) return;
  if (ctx.formatting) {
    swapProps(rTbl, "tblPr", recordTableFormatChange, child(oTbl, "tblPr"), ctx);
    alignTableGrid(oTbl, rTbl, ctx);
  }

  const oRows = oTbl.children.filter((c) => localName(c.name) === "tr");
  const rRows = rTbl.children.filter((c) => localName(c.name) === "tr");
  const anchors = histogramMatch(oRows.map(rowFingerprint), rRows.map(rowFingerprint));

  const rows: XmlElement[] = [];
  let oAt = 0;
  let rAt = 0;
  for (const anchor of [...anchors, { a: oRows.length, b: rRows.length }]) {
    // Rows inside a gap pair up by similarity, exactly like paragraphs, so a
    // row whose one cell changed is an edited row and not a replaced one.
    const pairings = resolveGap(oRows, rRows, oAt, anchor.a, rAt, anchor.b, ctx).filter((p) => p.kind === "pair");
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

  rTbl.children = [...rTbl.children.filter((c) => localName(c.name) !== "tr"), ...rows];
}

function rowFingerprint(tr: XmlElement): string {
  const cells = tr.children.filter((c) => localName(c.name) === "tc");
  return `${cells.length}\u0000${normalizeText(paragraphText(tr))}`;
}

function deletedRow(tr: XmlElement, ctx: Ctx): XmlElement {
  const clone = cloneXml(tr);
  markRowRevision(clone, "del", ctx.meta);
  wrapRowContent(clone, "del", ctx);
  return clone;
}

function insertedRow(tr: XmlElement, ctx: Ctx): XmlElement {
  markRowRevision(tr, "ins", ctx.meta);
  wrapRowContent(tr, "ins", ctx);
  return tr;
}

/** A matched row pair, or — when the cell counts differ — two coarse rows. */
function compareRow(oTr: XmlElement, rTr: XmlElement, ctx: Ctx): XmlElement[] {
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
  if (!oGrid || !rGrid || serializeXml(oGrid) === serializeXml(rGrid)) return;
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
function alignParagraphProps(
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
  let i = 0;
  while (i < entries.length) {
    if (entries[i].cls === "kept") {
      i++;
      continue;
    }
    let end = i;
    while (end + 1 < entries.length && entries[end + 1].cls !== "kept") end++;
    markRun(entries, i, end, ctx);
    i = end + 1;
  }
}

function markRun(entries: Entry[], start: number, end: number, ctx: Ctx): void {
  const shifted = start > 0 && isParagraph(entries[start - 1].el);
  if (shifted) {
    for (let i = start; i <= end; i++) {
      if (!isParagraph(entries[i].el) || !isParagraph(entries[i - 1].el)) continue;
      markParagraphGlyph(entries[i - 1].el, glyphKind(entries[i].cls), ctx.meta);
    }
    return;
  }

  const hasFollower = end + 1 < entries.length;
  for (let i = start; i <= end; i++) {
    if (!isParagraph(entries[i].el)) continue;
    if (i === end && !hasFollower) {
      ctx.note({
        code: "no-common-paragraph",
        detail: "the documents share no paragraph, so the final paragraph mark difference is left untracked",
      });
      continue;
    }
    markParagraphGlyph(entries[i].el, glyphKind(entries[i].cls), ctx.meta);
  }

  // Self-marked paragraphs join with the FOLLOWING one when reviewed, so the
  // head of each same-class chain must be able to wear either document's
  // paragraph properties.
  let at = start;
  while (at <= end) {
    let last = at;
    while (last + 1 <= end && entries[last + 1].cls === entries[at].cls) last++;
    const follower = entries[last + 1];
    if (follower && isParagraph(entries[at].el) && isParagraph(follower.el)) {
      const own = child(entries[at].el, "pPr");
      const other = child(follower.el, "pPr");
      if (entries[at].cls === "deleted") alignParagraphProps(entries[at].el, other, own, ctx);
      else alignParagraphProps(entries[at].el, own, other, ctx);
    }
    at = last + 1;
  }
}

function glyphKind(cls: BlockClass): "ins" | "del" {
  return cls === "inserted" ? "ins" : "del";
}

function isParagraph(el: XmlElement): boolean {
  return localName(el.name) === "p";
}
