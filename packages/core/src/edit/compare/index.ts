import { diffArrays } from "diff";
import { DocxDocument } from "../../docx.js";
import { XmlElement, child, cloneXml, localName, parseXml, sameXml, serializeXml } from "../../xml.js";
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
  recordTableFormatChange,
  recordTableGridChange,
  revisionWrapper,
} from "../suggest.js";
import { CONTAINERS, ContentToken, Placed, ReadOptions, isPassenger, paragraphText, readParagraph, runPropsKey, sliceItems, textBoxesIn } from "./content.js";
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
    | "unresolved-relationship"
    /** An input held tracked changes; it was compared as if they were accepted. */
    | "existing-revisions-accepted";
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
  /**
   * Write the result into `revised` instead of a copy of it, and return
   * `revised`. A copy is a whole second document in memory (114 MB for a
   * 217-page book), so a caller that does not need `revised` afterwards
   * should pass this. `original` is never changed either way.
   */
  reuseRevised?: boolean;
}

interface Ctx {
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

function compareInto(original: DocxDocument, merged: DocxDocument, options: CompareOptions, lean?: LeanBodies): DocxDocument {
  // Every revision takes an id; with none issued the body is the clone's own,
  // untouched, and its model needs no rebuild (refresh was a sixth of the time
  // on a document compared with itself).
  let issued = 0;
  const ctx: Ctx = {
    doc: merged,
    original,
    meta: {
      author: options.author,
      date: options.date,
      nextId: () => {
        issued++;
        return merged.nextRevisionId();
      },
    },
    formatting: options.formatting !== false,
    whitespace: options.whitespace !== false,
    note: options.onNote ?? (() => {}),
    scope: { from: null, into: null },
    noteIds: { footnote: new Map(), endnote: new Map() },
  };

  const originalBody = child(original.docRoot, "body");
  const mergedBody = child(merged.docRoot, "body");
  if (!originalBody || !mergedBody) throw new Error("document.xml has no w:body");

  // Before the body is touched, merged holds exactly the revised parts.
  reportUncomparedParts(original, merged, ctx);
  const headerFooterPairs = pairHeadersFooters(original, merged, lean);
  // Notes first: struck body text carries the ORIGINAL's note references,
  // and adoptFromOriginal re-points them through the ids this settles.
  compareNotes("footnote", original, merged, ctx, lean);
  compareNotes("endnote", original, merged, ctx, lean);

  if (lean) {
    compareDeferredBodies(lean, ctx);
  } else {
    const tail = mergedBody.children.filter((c) => localName(c.name) === "sectPr");
    const entries = alignBlocks(blockList(originalBody, "sectPr"), blockList(mergedBody, "sectPr"), ctx);
    applyParagraphMarks(entries, ctx);
    mergedBody.children = [...entries.map((e) => e.el), ...tail];
    giveStruckReferencesTheirOwnNotes("footnote", merged, ctx);
    giveStruckReferencesTheirOwnNotes("endnote", merged, ctx);
  }

  for (const [from, into] of headerFooterPairs) {
    const oRoot = original.headerFooterRoot(from);
    const mRoot = merged.headerFooterRoot(into);
    if (oRoot && mRoot) compareContainer(oRoot, mRoot, { ...ctx, scope: { from, into } });
  }

  // A deferred body has no model to rebuild.
  if (issued > 0 && !lean) merged.refresh();
  return merged;
}

/**
 * Compare two .docx files and return the result's bytes: what
 * `compareDocuments(load(original), load(revised)).save()` returns, byte for
 * byte, in a fraction of the memory.
 *
 * Neither body is ever held as one tree. Each loads as its top-level blocks'
 * XML text (DocxDocument's deferBody); each block is read once for its
 * fingerprint and its saved form; and only the blocks the alignment leaves
 * unmatched, or matches with a difference, are parsed again, compared,
 * marked, and written back as text. A document that needs the whole tree
 * (one with tracked changes of its own, percentage-width tables or charts,
 * which save() rewrites from the model) goes through compareDocuments.
 */
export function compareDocx(original: Uint8Array, revised: Uint8Array, options: CompareOptions): Uint8Array {
  const o = DocxDocument.load(original, { deferBody: true });
  const r = DocxDocument.load(revised, { deferBody: true });
  if (needsWholeTree(o) || needsWholeTree(r)) {
    return compareDocuments(DocxDocument.load(original), DocxDocument.load(revised), { ...options, reuseRevised: true }).save();
  }
  const lean: LeanBodies = { original: readDeferredBody(o, false), merged: readDeferredBody(r, true) };
  return compareInto(o, r, options, lean).save();
}

/** What compareDocx cannot do block by block (see compareDocx). */
const WHOLE_TREE =
  /<w:(?:ins|del|moveFrom|moveTo|rPrChange|pPrChange|sectPrChange|tblPrChange|tblGridChange|trPrChange|tcPrChange|numberingChange|cellIns|cellDel|cellMerge)[\s>/]|<w:tblW\b[^>]*w:w="[^"]*%"|ProgID="Package"/;

function needsWholeTree(doc: DocxDocument): boolean {
  if (doc.docRoot.name !== "w:document") return true;
  if (collectRevisions(doc).length > 0) return true;
  if ([...doc.documentRels.values()].some((rel) => rel.type.endsWith("/chart"))) return true;
  return (doc.deferredBody ?? []).some((block) => WHOLE_TREE.test(block));
}

/** A deferred body read once: each block's saved form and fingerprint, and what the rest of compare needs from the body. */
interface LeanBody {
  doc: DocxDocument;
  blocks: string[];
  fingerprints: string[];
  /** Each section-ending paragraph's w:sectPr, in order. */
  sections: XmlElement[];
  /** Note reference ids in document order. */
  noteRefs: Record<"footnote" | "endnote", string[]>;
}

interface LeanBodies {
  original: LeanBody;
  merged: LeanBody;
}

function readDeferredBody(doc: DocxDocument, isResult: boolean): LeanBody {
  const body: LeanBody = { doc, blocks: [], fingerprints: [], sections: [], noteRefs: { footnote: [], endnote: [] } };
  // The result's body as loaded, for save() to tell whether it changed: the
  // same strings as `blocks` except where preparing a block changed it.
  const loaded: string[] = [];
  for (const raw of doc.deferredBody ?? []) {
    const el = parseXml(raw);
    const before = isResult ? serializeXml(el) : "";
    doc.prepareDeferredBlock(el);
    const xml = serializeXml(el);
    body.blocks.push(xml);
    if (isResult) loaded.push(before === xml ? xml : before);
    body.fingerprints.push(blockFingerprint(el, doc));
    if (localName(el.name) === "p") {
      const sectPr = child(child(el, "pPr"), "sectPr");
      if (sectPr) body.sections.push(sectPr);
    }
    const walk = (e: XmlElement): void => {
      const ln = localName(e.name);
      if (ln === "footnoteReference") body.noteRefs.footnote.push(e.attrs["w:id"]);
      else if (ln === "endnoteReference") body.noteRefs.endnote.push(e.attrs["w:id"]);
      for (const c of e.children) walk(c);
    };
    walk(el);
  }
  if (isResult) doc.rememberLoadedBody(loaded);
  doc.setDeferredBody(body.blocks);
  return body;
}

/**
 * The body compare of compareInto, a stretch at a time. Anchored blocks whose
 * text is identical stay text; the blocks of each gap, and anchored blocks
 * that differ, are parsed, emitted exactly as alignBlocks would, and written
 * back to text once their paragraph marks are settled. A stretch's marks
 * touch only the kept paragraph before it and read the one after, so every
 * entry before the latest anchor is final once that anchor is placed.
 */
function compareDeferredBodies(lean: LeanBodies, ctx: Ctx): void {
  const O = lean.original;
  const R = lean.merged;
  const parse = (side: LeanBody, i: number): XmlElement => {
    const el = parseXml(side.blocks[i]);
    side.doc.prepareDeferredBlock(el);
    return el;
  };
  type Slot = { el: XmlElement | null; xml?: string; cls: BlockClass };
  const element = (slot: Slot): Entry => {
    if (!slot.el) {
      slot.el = parseXml(slot.xml!);
      R.doc.prepareDeferredBlock(slot.el);
    }
    return slot as Entry;
  };
  // compareInto settles every paragraph mark after the whole body, so mark
  // revisions take their ids after all the body's other revisions. Here each
  // stretch is settled as soon as it is complete, with a stand-in id that is
  // swapped for the real one, in the same order, once the body is done.
  let stand = 0;
  const markCtx: Ctx = { ...ctx, meta: { ...ctx.meta, nextId: () => `dxw-mark-${stand++}` as unknown as number } };
  const out: string[] = [];
  // Compared blocks that hold note references stay trees until the note pass
  // below: written and read back, an emptied wrapper (<w:ins></w:ins>, whose
  // one child is the editor's placeholder run) would come back self-closing.
  const NOTE_REF = /<w:(?:footnote|endnote)Reference /;
  const trees = new Map<number, XmlElement>();
  let pending: Slot[] = [];
  const settle = (keepLast: boolean): void => {
    if (pending.some((slot) => slot.cls !== "kept")) applyParagraphMarks(pending.map(element), markCtx);
    const last = keepLast ? pending.pop() : undefined;
    for (const slot of pending) {
      const xml = slot.el ? serializeXml(slot.el) : slot.xml!;
      if (slot.el && NOTE_REF.test(xml)) trees.set(out.length, slot.el);
      out.push(xml);
    }
    pending = last ? [last] : [];
  };

  // Sparse: a gap's blocks are parsed into them for that gap only.
  const oEls: XmlElement[] = new Array(O.blocks.length);
  const rEls: XmlElement[] = new Array(R.blocks.length);
  let oAt = 0;
  let rAt = 0;
  for (const anchor of [...histogramMatch(O.fingerprints, R.fingerprints), { a: O.blocks.length, b: R.blocks.length }]) {
    const anchored = anchor.a < O.blocks.length && anchor.b < R.blocks.length;
    if (anchor.a > oAt || anchor.b > rAt) {
      for (let i = oAt; i < anchor.a; i++) oEls[i] = parse(O, i);
      for (let j = rAt; j < anchor.b; j++) rEls[j] = parse(R, j);
      // emitGap looks at the block after the gap.
      if (anchored) rEls[anchor.b] = parse(R, anchor.b);
      if (pending.length > 0) element(pending[pending.length - 1]);
      emitGap(oEls, rEls, oAt, anchor.a, rAt, anchor.b, pending as Entry[], ctx);
      for (let i = oAt; i < anchor.a; i++) delete oEls[i];
      for (let j = rAt; j < anchor.b; j++) delete rEls[j];
    }
    if (anchored) {
      if (O.blocks[anchor.a] === R.blocks[anchor.b]) {
        pending.push({ el: null, xml: R.blocks[anchor.b], cls: "kept" });
      } else {
        pending.push(emitMatched(parse(O, anchor.a), rEls[anchor.b] ?? parse(R, anchor.b), ctx));
      }
      delete rEls[anchor.b];
      settle(true);
    }
    oAt = anchor.a + 1;
    rAt = anchor.b + 1;
  }
  settle(false);
  // The marks' real ids, taken before the note pass's, as compareInto does.
  const ids = Array.from({ length: stand }, () => String(ctx.meta.nextId()));

  // One note, one reference: the blocks that hold references, in order, are
  // all giveStruckReferencesTheirOwnNotes needs to see. A block kept as text
  // was never compared, so reading it again gives it back unchanged.
  const withRefs = out.flatMap((xml, i) => (NOTE_REF.test(xml) ? [i] : []));
  if (withRefs.length > 0) {
    const tree = (i: number): XmlElement => {
      const kept = trees.get(i);
      if (kept) return kept;
      const el = parseXml(out[i]);
      R.doc.prepareDeferredBlock(el);
      return el;
    };
    const body: XmlElement = { name: "w:body", attrs: {}, children: withRefs.map(tree), text: "" };
    giveStruckReferencesTheirOwnNotes("footnote", ctx.doc, ctx, body);
    giveStruckReferencesTheirOwnNotes("endnote", ctx.doc, ctx, body);
    withRefs.forEach((i, k) => (out[i] = serializeXml(body.children[k])));
  }
  if (stand > 0) {
    for (let i = 0; i < out.length; i++) {
      if (out[i].includes('"dxw-mark-')) out[i] = out[i].replace(/"dxw-mark-(\d+)"/g, (_, n: string) => `"${ids[Number(n)]}"`);
    }
  }
  R.doc.setDeferredBody(out);
}

/** Align and diff one container's blocks in place (a header, a note). */
function compareContainer(oRoot: XmlElement, mRoot: XmlElement, ctx: Ctx): void {
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
 * Header and footer parts paired across the documents: section by section,
 * kind by kind (default, first, even), each pair once.
 */
function pairHeadersFooters(original: DocxDocument, merged: DocxDocument, lean?: LeanBodies): [string, string][] {
  const refs = (doc: DocxDocument): Map<string, string> => {
    const out = new Map<string, string>();
    const body = child(doc.docRoot, "body");
    // A deferred body's tree holds only the body's own sectPr; its section
    // breaks were read from the blocks.
    const scanned = lean ? (doc === original ? lean.original : lean.merged).sections : null;
    const sects = [
      ...(scanned ?? []),
      ...(body?.children ?? []).flatMap((c) =>
        localName(c.name) === "sectPr" ? [c] : localName(c.name) === "p" ? [child(child(c, "pPr"), "sectPr")].filter((x): x is XmlElement => !!x) : [],
      ),
    ];
    sects.forEach((sp, i) => {
      for (const ref of sp.children) {
        const ln = localName(ref.name);
        if (ln !== "headerReference" && ln !== "footerReference") continue;
        const target = doc.documentRels.get(ref.attrs["r:id"])?.target;
        if (target) out.set(`${i}:${ln}:${ref.attrs["w:type"] ?? "default"}`, target);
      }
    });
    return out;
  };
  const mine = refs(merged);
  const pairs: [string, string][] = [];
  const done = new Set<string>();
  for (const [key, from] of refs(original)) {
    const into = mine.get(key);
    if (!into || done.has(into)) continue;
    done.add(into);
    pairs.push([from, into]);
  }
  return pairs;
}

/**
 * Footnotes or endnotes, paired in the order their references appear and
 * compared inside. A note only the revised document has gets its content
 * marked inserted; one only the original had is copied into the result
 * under a fresh id with its content struck, and struck body text pointing
 * at it is re-pointed (adoptFromOriginal, through ctx.noteIds).
 */
function compareNotes(kind: "footnote" | "endnote", original: DocxDocument, merged: DocxDocument, ctx: Ctx, lean?: LeanBodies): void {
  const tree = (doc: DocxDocument, create = false): XmlElement | null =>
    kind === "footnote" ? doc.footnotesTree(create) : doc.endnotesTree(create);
  const oRoot = tree(original);
  const mRoot = tree(merged);
  if (!oRoot && !mRoot) return;
  const refName = `${kind}Reference`;
  const inOrder = (doc: DocxDocument, root: XmlElement | null): XmlElement[] => {
    if (!root) return [];
    const byId = new Map(
      root.children.filter((c) => localName(c.name) === kind && !c.attrs["w:type"]).map((c) => [c.attrs["w:id"], c] as const),
    );
    const out: XmlElement[] = [];
    const seen = new Set<string>();
    const take = (id: string): void => {
      const note = byId.get(id);
      if (note && !seen.has(id)) {
        seen.add(id);
        out.push(note);
      }
    };
    const walk = (e: XmlElement): void => {
      if (localName(e.name) === refName) take(e.attrs["w:id"]);
      for (const c of e.children) walk(c);
    };
    // A deferred body's references were read from its blocks, in order.
    if (lean) for (const id of (doc === original ? lean.original : lean.merged).noteRefs[kind]) take(id);
    walk(doc.docRoot);
    return out;
  };
  const oNotes = inOrder(original, oRoot);
  const rNotes = inOrder(merged, mRoot);
  const text = (n: XmlElement): string => normalizeText(paragraphText(n));
  // A note with the same text on both sides is the same note wherever its
  // reference now sits: a note's place is its reference's, so moving the
  // paragraph that holds it must not unpair it. Equal texts pair first, in
  // reference order.
  const pairs: [number, number][] = [];
  const waiting = new Map<string, number[]>();
  rNotes.forEach((n, j) => {
    const list = waiting.get(text(n));
    if (list) list.push(j);
    else waiting.set(text(n), [j]);
  });
  oNotes.forEach((n, i) => {
    const j = waiting.get(text(n))?.shift();
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
  const rest = histogramMatch(oRest.map((i) => text(oNotes[i])), rRest.map((j) => text(rNotes[j])));
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
    ids.set(oNotes[i].attrs["w:id"], rNotes[j].attrs["w:id"]);
    const before = ctx.meta.nextId;
    let touched = false;
    compareContainer(oNotes[i], rNotes[j], { ...ctx, meta: { ...ctx.meta, nextId: () => ((touched = true), before()) } });
    changed ||= touched;
  }
  rNotes.forEach((note, j) => {
    if (pairedR.has(j)) return;
    for (const block of blockList(note)) if (localName(block.name) === "p") wrapContainerContent(block, "ins", ctx);
    changed = true;
  });
  const removed = oNotes.filter((_, i) => !pairedO.has(i));
  if (removed.length > 0) {
    const root = tree(merged, true);
    if (root) {
      let next = 1 + Math.max(0, ...root.children.map((c) => Number(c.attrs["w:id"]) || 0));
      for (const o of removed) {
        const copy = cloneXml(o);
        copy.attrs["w:id"] = String(next);
        ids.set(o.attrs["w:id"], String(next++));
        for (const block of blockList(copy)) {
          if (localName(block.name) !== "p") continue;
          adoptFromOriginal(block, { ...ctx, scope: { from: null, into: null } });
          wrapContainerContent(block, "del", ctx);
        }
        root.children.push(copy);
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
function giveStruckReferencesTheirOwnNotes(kind: "footnote" | "endnote", merged: DocxDocument, ctx: Ctx, body: XmlElement = merged.docRoot): void {
  const root = kind === "footnote" ? merged.footnotesTree() : merged.endnotesTree();
  if (!root) return;
  const refName = `${kind}Reference`;
  const refs = new Map<string, { ref: XmlElement; struck: boolean }[]>();
  const walk = (e: XmlElement, struck: boolean): void => {
    if (localName(e.name) === refName) {
      const list = refs.get(e.attrs["w:id"]) ?? [];
      list.push({ ref: e, struck });
      refs.set(e.attrs["w:id"], list);
    }
    for (const c of e.children) walk(c, struck || localName(c.name) === "del" || rowDeleted(c));
  };
  // A deleted row records it in its trPr, and its cells (a nested table's
  // too) may hold references with no w:del around them.
  const rowDeleted = (tr: XmlElement): boolean =>
    localName(tr.name) === "tr" && !!child(child(tr, "trPr"), "del");
  walk(body, false);
  let next = 1 + Math.max(0, ...root.children.map((c) => Number(c.attrs["w:id"]) || 0));
  let changed = false;
  for (const [id, list] of refs) {
    if (list.length < 2) continue;
    const note = root.children.find((c) => localName(c.name) === kind && c.attrs["w:id"] === id);
    if (!note) continue;
    // The live reference keeps the note; every struck one beyond it gets a copy.
    const keep = list.find((r) => !r.struck) ?? list[0];
    for (const r of list) {
      if (r === keep || !r.struck) continue;
      const copy = cloneXml(note);
      copy.attrs["w:id"] = String(next);
      r.ref.attrs["w:id"] = String(next++);
      for (const block of blockList(copy)) if (localName(block.name) === "p") wrapContainerContent(block, "del", ctx);
      root.children.push(copy);
      changed = true;
    }
  }
  if (changed) merged.markNotePartsChanged([kind === "footnote" ? "footnotes" : "endnotes"]);
}

function blockList(container: XmlElement, ...skip: string[]): XmlElement[] {
  return container.children.filter((c) => !skip.includes(localName(c.name)));
}

/** Say so when a part outside the body differs, rather than imply it matched. */
function reportUncomparedParts(original: DocxDocument, revised: DocxDocument, ctx: Ctx): void {
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
function blockFingerprint(el: XmlElement, doc: DocxDocument): string {
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
  // A gap with no paragraph on either side has no mark to stand on: an added
  // or removed paragraph there can only join a neighbour that is a table, and
  // review cannot do that (a heading rewritten between two tables came back
  // from reject-all with an empty paragraph after it). Such a gap, like one
  // opening the container, must pair something and diff it in place.
  const before = entries[entries.length - 1];
  const unanchored = (!before || !isParagraph(before.el)) && (rHi >= rBlocks.length || !isParagraph(rBlocks[rHi]));
  const pairings = resolveGap(oBlocks, rBlocks, oLo, oHi, rLo, rHi, ctx, (oLo === 0 && rLo === 0) || unanchored);
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

/** Gap size (original x revised blocks) past which pairing turns windowed. */
const GAP_PAIR_CAP = 200 * 200;
/** How far either side of its own relative position a block looks, windowed. */
const GAP_WINDOW = 50;

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

  const keys: KeyCache = new Map();
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
      if (concatMatches([oBlocks[i]], [rBlocks[j], rBlocks[j + 1]], ctx, keys)) {
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
      if (concatMatches([oBlocks[i], oBlocks[i + 1]], [rBlocks[j]], ctx, keys)) {
        accept({ kind: "merge", o: i, r: j, other: i + 1 });
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
    const [jLo, jHi] = span(i);
    for (let j = jLo; j < jHi; j++) {
      if (rUsed.has(j)) continue;
      if (!pairable(oBlocks[i], rBlocks[j], ctx)) continue;
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
    accept({ kind: "pair", o: cand.o, r: cand.r, other: -1 });
    oUsed.add(cand.o);
    rUsed.add(cand.r);
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
    if (os.length === 1 && rs.length === 1 && sameShape(oBlocks[os[0]], rBlocks[rs[0]], ctx)) {
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
        if (rUsed.has(j) || !pairable(oBlocks[i], rBlocks[j], ctx)) continue;
        accept({ kind: "pair", o: i, r: j, other: -1 });
        break;
      }
    }
  }
  return out.sort((a, b) => a.o - b.o);
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
function pairable(a: XmlElement, b: XmlElement, ctx: Ctx): boolean {
  if (localName(a.name) !== localName(b.name)) return false;
  if (!isParagraph(a)) return true;
  return sectionKey(a, ctx.original) === sectionKey(b, ctx.doc);
}

/**
 * The section a paragraph ends, as its w:sectPr with each relationship id
 * replaced by the part it points to, so the two documents' ids compare; "" for
 * a paragraph that ends no section.
 */
function sectionKey(el: XmlElement, doc: DocxDocument): string {
  const sectPr = child(child(el, "pPr"), "sectPr");
  if (!sectPr) return "";
  return serializeXml(sectPr).replace(/r:id="([^"]*)"/g, (_, id: string) => `r:id="${doc.documentRels.get(id)?.target ?? id}"`);
}

/**
 * Two paragraphs, two tables, two block-level content controls (or custom XML
 * elements), or two table rows with the same number of cells.
 */
function sameShape(a: XmlElement, b: XmlElement, ctx: Ctx): boolean {
  if (isParagraph(a) && isParagraph(b)) return pairable(a, b, ctx);
  if (localName(a.name) === localName(b.name) && ["tbl", "sdt", "customXml"].includes(localName(a.name))) return true;
  const cells = (tr: XmlElement) => tr.children.filter((c) => localName(c.name) === "tc").length;
  return localName(a.name) === "tr" && localName(b.name) === "tr" && cells(a) === cells(b);
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
function readOptions(ctx: Ctx, original: boolean): ReadOptions {
  const noteKey = (ref: XmlElement): string => {
    const kind = localName(ref.name) === "footnoteReference" ? "footnote" : "endnote";
    const id = ref.attrs["w:id"];
    return original ? `${ctx.noteIds[kind].get(id) ?? `original ${id}`}` : id;
  };
  return { whitespace: ctx.whitespace, noteKey };
}

function concatMatches(oEls: XmlElement[], rEls: XmlElement[], ctx: Ctx, keys: KeyCache): boolean {
  const a = concatKey(oEls, true, ctx, keys);
  if (!a) return false;
  const b = concatKey(rEls, false, ctx, keys);
  return !!b && a.length === b.length && a.every((part, i) => part === b[i]);
}

/** Split/merge key parts per paragraph, kept for the one gap that tries them. */
type KeyCache = Map<XmlElement, string[] | null>;

function concatKey(els: XmlElement[], original: boolean, ctx: Ctx, keys: KeyCache): string[] | null {
  const out: string[] = [];
  for (const el of els) {
    const parts = paragraphKeyParts(el, original, ctx, keys);
    if (!parts) return null;
    for (const part of parts) out.push(part);
  }
  return out;
}

/**
 * One paragraph's tokens with their run properties and containers, or null
 * when it cannot take part in a split or merge. Read once per gap: the
 * split and merge passes try every nearby pair, and re-reading both
 * paragraphs for each try was half the time of a long rewritten document
 * (the Yiddish novel, 90% of paragraphs edited: 3.4 of 6.9 s). Kept for one
 * gap only, so a long document's keys are not all alive at once.
 */
function paragraphKeyParts(el: XmlElement, original: boolean, ctx: Ctx, keys: KeyCache): string[] | null {
  const cached = keys.get(el);
  if (cached !== undefined) return cached;
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
  keys.set(el, parts);
  return parts;
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
function adoptFromOriginal(root: XmlElement, ctx: Ctx): void {
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
    for (const key of Object.keys(el.attrs)) {
      if (!key.startsWith("r:") || !REL_ATTRS.includes(localName(key))) continue;
      const id = el.attrs[key];
      const source = theirs.get(id);
      if (!source) continue;
      const here = mine.get(id);
      if (here && here.target === source.target && here.type === source.type) continue;
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

function emitDeleted(el: XmlElement, ctx: Ctx): Entry {
  const clone = cloneXml(el);
  adoptFromOriginal(clone, ctx);
  markWholeBlock(clone, "del", ctx);
  return { el: clone, cls: "deleted" };
}

function emitInserted(el: XmlElement, ctx: Ctx): Entry {
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
function wrapContainerContent(container: XmlElement, kind: "ins" | "del", ctx: Ctx): void {
  // Run containers the schema lets a revision sit INSIDE (a simple field's
  // result runs among them), as opposed to an inline content control or a
  // run-level customXml, which a revision may wrap whole. Left plain, a removed
  // paragraph's checkbox controls ("☒ ☐") survived accept-all and joined the
  // paragraph before (probe2-form-checkboxes).
  const NESTED = ["hyperlink", "smartTag", "dir", "bdo", "fldSimple"];
  const WRAPPABLE = ["r", "sdt", "customXml"];
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
  if (sameXml(oEl, rEl)) return;

  if (ctx.formatting) alignParagraphProps(rEl, child(rEl, "pPr"), child(oEl, "pPr"), ctx);

  const oc = readParagraph(oEl, readOptions(ctx, true));
  const rc = readParagraph(rEl, readOptions(ctx, false));
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
  if (ctx.formatting) {
    swapProps(rTbl, "tblPr", recordTableFormatChange, child(oTbl, "tblPr"), ctx);
    alignTableGrid(oTbl, rTbl, ctx);
  }

  compareRows(oTbl, rTbl, ctx);
}

/** Row-level children: rows, and content controls or custom XML holding rows. */
const ROW_LEVEL = ["tr", "sdt", "customXml"];

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

  rParent.children = [...rParent.children.filter((c) => !ROW_LEVEL.includes(localName(c.name))), ...rows];
}

function rowFingerprint(tr: XmlElement): string {
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

function deletedRow(el: XmlElement, ctx: Ctx): XmlElement {
  const clone = cloneXml(el);
  for (const tr of rowsIn(clone)) {
    markRowRevision(tr, "del", ctx.meta);
    wrapRowContent(tr, "del", ctx);
  }
  return clone;
}

function insertedRow(el: XmlElement, ctx: Ctx): XmlElement {
  for (const tr of rowsIn(el)) {
    markRowRevision(tr, "ins", ctx.meta);
    wrapRowContent(tr, "ins", ctx);
  }
  return el;
}

/** A matched row pair, or — when the cell counts differ — two coarse rows. */
function compareRow(oTr: XmlElement, rTr: XmlElement, ctx: Ctx): XmlElement[] {
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
  // Tables carry their own row revisions and no pilcrow, so the paragraphs
  // between them are marked stretch by stretch: a stretch after a table has
  // no paragraph before it to take a mark from, exactly like one opening the
  // container, and marks its own (a paragraph removed right after a removed
  // table used to get no mark at all, and survived accept-all empty).
  let at = start;
  while (at <= end) {
    if (!isParagraph(entries[at].el)) {
      at++;
      continue;
    }
    let last = at;
    while (last + 1 <= end && isParagraph(entries[last + 1].el)) last++;
    markParagraphRun(entries, at, last, ctx);
    at = last + 1;
  }
}

function markParagraphRun(entries: Entry[], start: number, end: number, ctx: Ctx): void {
  // Shifting moves each mark one paragraph up, and a section break lives in
  // its paragraph's mark: shifted, a kept break would be lent to the stretch
  // after it, and an added break would read as already there. So a stretch
  // next to or holding a break marks its own paragraphs.
  const endsSection = (el: XmlElement): boolean => isParagraph(el) && !!child(child(el, "pPr"), "sectPr");
  const before = start > 0 ? entries[start - 1].el : undefined;
  const shifted =
    !!before && isParagraph(before) && !endsSection(before) && !entries.slice(start, end + 1).some((e) => endsSection(e.el));
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
