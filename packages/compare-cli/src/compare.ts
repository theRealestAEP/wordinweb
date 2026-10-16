import { closeSync, createWriteStream, mkdtempSync, openSync, rmSync, statSync, statfsSync, writeSync } from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join, resolve } from "node:path";
import { XmlElement, child, localName, parseXml, sameXml, serializeXml } from "../../core/src/xml.js";
import { histogramMatch } from "../../core/src/edit/compare/histogram.js";
import { paragraphText } from "../../core/src/edit/compare/content.js";
import {
  BlockClass,
  CompareNote,
  CompareOptions,
  Ctx,
  Entry,
  GapSide,
  NoteRef,
  applyMarkStep,
  compareContext,
  compareHeadersFooters,
  compareSettings,
  compareLastSection,
  compareNotes,
  compareRow,
  compareTableProps,
  copyNotesForStruckReferences,
  settleLeftoverNotes,
  deletedRow,
  emitDeleted,
  emitGapOp,
  emitInserted,
  emitMatched,
  gapOps,
  insertedRow,
  noteRefsIn,
  originalProps,
  pairHeadersFooters,
  paragraphKeyParts,
  refersToChangedPart,
  paragraphMarkSteps,
  reportUncomparedParts,
  resolveGap,
} from "../../core/src/edit/compare/index.js";
import { MARKERS, RevisionKind, RevisionRef, collectRevisionsIn } from "../../core/src/edit/suggest.js";
import { CompareError, isDiskFull } from "./errors.js";
import { PdfRenderError } from "../../core/src/pdf/errors.js";
import { renderDocxToPdf } from "../../core/src/pdf/render.js";
import { Input, ROW_SLOT, Side, readInput, stubSource } from "./input.js";
import { barParagraphs, flatten, markLastParagraph } from "./printable.js";
import { DiskNotes } from "./notes.js";
import { BlockStore } from "./store.js";
import { SummarySpec, summaryXml } from "./summary.js";
import { ZipEntry, ZipReader, ZipWriter, mb } from "./zip.js";

export { CompareError, type CompareErrorCode } from "./errors.js";

export interface CompareFilesOptions extends Omit<CompareOptions, "reuseRevised"> {
  /** Where the working files go. Default: the system's temporary directory. */
  tempDir?: string;
  /**
   * Also write a printable copy to `path`: the same document with its tracked
   * changes as plain formatting (printable.ts), from the same pass.
   * `beforeSectPr` is XML put at the end of its body, ahead of the body's
   * final section properties: a page break and a summary, say. As a
   * function it is given the result's revision counts, known by then.
   */
  printable?: { path: string; beforeSectPr?: string | ((revisions: Record<RevisionKind, number>) => string) };
  /**
   * End the printable copy (and so the PDF) with a summary page the compare
   * writes itself (summary.ts): the date, the two documents' names, the
   * revision counts, their total, and each distinct note. It follows
   * `printable.beforeSectPr` when both are given. It is never in the
   * tracked result.
   */
  summary?: SummarySpec;
  /**
   * Also write the printable copy as a PDF to `path`, rendered without a
   * browser (core's pdf/render.ts) from the printable .docx (a working file
   * when `printable` is not asked for), with its review comments beside the
   * pages when `comments` is set.
   */
  pdf?: { path: string; comments?: boolean };
  /**
   * Called (and awaited) once the result and the printable copy are written,
   * before the PDF is rendered: a caller that keeps the redline somewhere
   * safe does it here, so it survives a render that fails or runs out of
   * memory.
   */
  onResultWritten?: () => void | Promise<void>;
}

export interface CompareFilesResult {
  /** The result's revisions by kind, as collectRevisions counts them. */
  revisions: Record<RevisionKind, number>;
  /** The PDF's page count, when `pdf` was asked for. */
  pdfPages?: number;
  /** With `pdf`: font families no installed font holds (a substitute drew
   * them), and characters no installed font has (each an empty box). */
  missingFonts?: string[];
  missingCharacters?: string[];
  /** With `pdf`: picture parts drawn as empty frames (a format the renderer
   * does not read, or an image over its pixel budget). */
  unsupportedImages?: string[];
}

/**
 * Compare two .docx files on disk and write the result to `outPath`: the
 * revised document with every difference from the original as a tracked
 * change, as compareDocuments gives it.
 *
 * Memory does not grow with the size of the documents' bodies. Each body is
 * written to disk block by block with a small index (hashes, kinds, section
 * breaks), a table's rows as blocks of their own; the alignment runs on the
 * index, and the blocks are read back one pair at a time to be compared and
 * written out. The notes go to disk the same way. Everything else outside
 * the body (styles, numbering, headers, footers) is held in memory, at most
 * MAX_HELD_BYTES together, checked against the zip's directory before any
 * of it is inflated (too_large), so a compare fits in a 256 MB heap; each binary
 * part is copied from the input zip unread. Buffers are fixed in size, or
 * as large as the largest single block, row or note, or those limits.
 *
 * Every error is a CompareError, with a code and the input at fault.
 */
export async function compareFiles(originalPath: string, revisedPath: string, outPath: string, options: CompareFilesOptions): Promise<CompareFilesResult> {
  for (const [which, path] of [["original", originalPath], ["revised", revisedPath]] as const) {
    const stat = statSync(path, { throwIfNoEntry: false });
    if (!stat) throw new CompareError("missing", `${path}: the ${which} file does not exist`, which);
    if (!stat.isFile()) throw new CompareError("missing", `${path}: the ${which} file is not a file`, which);
  }
  for (const out of [outPath, options.printable?.path, options.pdf?.path]) {
    if (out === undefined) continue;
    // The result is written while the inputs are still read.
    if ([originalPath, revisedPath].some((path) => resolve(path) === resolve(out))) {
      throw new CompareError("result_is_input", `${out}: the result must not overwrite an input`);
    }
    if (!statSync(dirname(resolve(out)), { throwIfNoEntry: false })?.isDirectory()) {
      throw new CompareError("result_folder_missing", `${out}: the result's folder does not exist`);
    }
  }
  const zips: ZipReader[] = [];
  const stores: { close(): void }[] = [];
  let dir: string | null = null;
  try {
    dir = mkdtempSync(join(options.tempDir ?? tmpdir(), "docx-compare-"));
    const oZip = openInput(originalPath, "original");
    zips.push(oZip);
    const rZip = openInput(revisedPath, "revised");
    zips.push(rZip);
    checkDiskSpace(dir, [oZip, rZip]);

    const O = await readNamed(oZip, "original", dir);
    stores.push(O.blocks, O.texts, O.rows, ...notesOf(O));
    const R = await readNamed(rZip, "revised", dir);
    stores.push(R.blocks, R.texts, R.rows, ...notesOf(R));

    // Every note once more for the summary page, which lists them.
    const noted: string[] = [];
    const onNote = options.onNote;
    const note = (n: CompareNote): void => {
      noted.push(n.detail);
      onNote?.(n);
    };
    options = { ...options, onNote: note };
    for (const input of [O, R]) {
      if (input.hadRevisions) {
        note({ code: "existing-revisions-accepted", detail: `the ${input.side} document held tracked changes; it was compared as if they were accepted` });
      }
    }
    const ctx = compareContext(O.shell, R.shell, options);
    reportUncomparedParts(O.shell, R.shell, ctx);
    const headerFooterPairs = pairHeadersFooters(O.shell, R.shell, { original: O.sectPrs, merged: R.sectPrs });
    // Notes first: struck body text carries the ORIGINAL's note references,
    // and adoptFromOriginal re-points them through the ids this settles.
    for (const kind of ["footnote", "endnote"] as const) {
      compareNotes(kind, O.shell, R.shell, ctx, { original: O.noteRefs[kind], merged: R.noteRefs[kind] }, { original: O.notes[kind], merged: R.notes[kind] });
    }

    const plan = planBody(O, R, ctx);
    const body = new BlockStore(join(dir, "result-blocks"));
    stores.push(body);
    // The PDF is rendered from the printable copy: a working file when the
    // caller did not ask for the copy itself.
    const printableSpec = options.printable ?? (options.pdf ? { path: join(dir, "printable.docx") } : undefined);
    const printable = printableSpec ? new BlockStore(join(dir, "printable-blocks")) : null;
    if (printable) stores.push(printable);
    // A save drops a chart part whose relationship the body it sees does not use.
    const charts = [O, R].some((input) => [...input.shell.documentRels.values()].some((rel) => rel.type.endsWith("/chart")));
    const sink = new BodySink(body, printable, charts ? new Set() : null);
    emitBody(plan, O, R, ctx, sink);
    const noteIds = giveStruckReferencesTheirOwnNotes(body, sink.withNoteRefs, O, R, ctx);

    compareHeadersFooters(O.shell, R.shell, headerFooterPairs, ctx);
    compareLastSection(O.shell, R.shell, ctx);
    compareSettings(O.shell, R.shell, ctx);
    sink.finish(!!child(child(child(R.shell.docRoot, "body"), "sectPr"), "sectPrChange"));

    // The body's revisions were counted as it was written; the rest are in
    // the shell (headers, footers) and in the notes the compare changed.
    const revisions = sink.revisions;
    for (const root of R.shell.revisionRoots()) countRevisions(root, revisions);
    for (const notes of notesOf(R)) notes.forEachChange((note) => countRevisions(note, revisions));

    const custom = printableSpec?.beforeSectPr;
    const beforeSectPr =
      (typeof custom === "function" ? custom(revisions) : custom ?? "") +
      (options.summary ? summaryXml(options.summary, options.date ?? new Date().toISOString(), revisions, noted, child(child(R.shell.docRoot, "body"), "sectPr")) : "");
    await writeResult(outPath, dir, O, R, body, noteIds, sink.relIds, printable && printableSpec ? { path: printableSpec.path, beforeSectPr, store: printable } : null);
    await options.onResultWritten?.();
    if (options.pdf && printableSpec) {
      const rendered = await renderDocxToPdf(printableSpec.path, createWriteStream(options.pdf.path), { comments: options.pdf.comments });
      return { revisions, pdfPages: rendered.pages, missingFonts: rendered.missingFonts, missingCharacters: rendered.missingCharacters, unsupportedImages: rendered.unsupportedImages };
    }
    return { revisions };
  } catch (error) {
    if (error instanceof CompareError) throw error;
    const message = error instanceof Error ? error.message : String(error);
    if (error instanceof PdfRenderError && error.code === "unsupported_script") throw new CompareError("unsupported_script", message, undefined, { cause: error });
    throw new CompareError(isDiskFull(error) ? "disk_full" : "failed", message, undefined, { cause: error });
  } finally {
    for (const store of stores) store.close();
    for (const zip of zips) zip.close();
    if (dir) rmSync(dir, { recursive: true, force: true });
  }
}

function notesOf(input: Input): DiskNotes[] {
  return [input.notes.footnote, input.notes.endnote].filter((notes): notes is DiskNotes => notes !== null);
}

function openInput(path: string, side: Side): ZipReader {
  try {
    return ZipReader.open(path);
  } catch (error) {
    throw inputError(error, side, path);
  }
}

async function readNamed(zip: ZipReader, side: Side, dir: string): Promise<Input> {
  try {
    return await readInput(zip, side, dir);
  } catch (error) {
    throw inputError(error, side, zip.path);
  }
}

/**
 * An error reading an input, as a CompareError naming the input: its own
 * code when it has one, disk_full when the disk filled up, else damaged (a
 * part that does not parse).
 */
function inputError(error: unknown, side: Side, path: string): CompareError {
  const message = error instanceof Error ? error.message : String(error);
  const named = message.startsWith(path) ? message : `${path}: ${message}`;
  if (isDiskFull(error)) return new CompareError("disk_full", named, undefined, { cause: error });
  return new CompareError(error instanceof CompareError ? error.code : "damaged", named, side, { cause: error });
}

/**
 * The working files take about this many times the two bodies' XML: each
 * body extracted, its blocks and their text, and the result's body before
 * and after it is deflated.
 */
const DISK_PER_BODY_BYTE = 5;

function checkDiskSpace(dir: string, zips: ZipReader[]): void {
  let xml = 0;
  for (const zip of zips) for (const entry of zip.entries.values()) if (/document\d*\.xml$/.test(entry.name)) xml += entry.size;
  const fs = statfsSync(dir);
  const free = fs.bavail * fs.bsize;
  const need = DISK_PER_BODY_BYTE * xml;
  if (need > free) throw new CompareError("disk_full", `not enough disk space in ${dir}: about ${mb(need)} needed, ${mb(free)} free`);
}

// ---------------------------------------------------------------------------
// The plan: the result's body as a list of entries, each naming its source
// ---------------------------------------------------------------------------

/** An anchored block whose XML is the same on both sides: copied, not compared. */
const COPY = 0;
/** A matched pair: emitMatched. */
const MATCH = 1;
const DELETED = 2;
const INSERTED = 3;
/** A split's kept paragraph; the next entry is its SPLIT_SECOND. */
const SPLIT = 4;
const SPLIT_SECOND = 5;
/** A merge's kept paragraph; the next entry is its MERGE_SECOND. */
const MERGE = 6;
const MERGE_SECOND = 7;

/** The entries of the result's body, in order, as three columns. */
class Plan {
  op = new Uint8Array(1024);
  o = new Int32Array(1024);
  r = new Int32Array(1024);
  n = 0;

  push(op: number, o: number, r: number): void {
    if (this.n === this.op.length) {
      const op2 = new Uint8Array(this.n * 2);
      op2.set(this.op);
      this.op = op2;
      const o2 = new Int32Array(this.n * 2);
      o2.set(this.o);
      this.o = o2;
      const r2 = new Int32Array(this.n * 2);
      r2.set(this.r);
      this.r = r2;
    }
    this.op[this.n] = op;
    this.o[this.n] = o;
    this.r[this.n] = r;
    this.n++;
  }

  /** Whether entry k's element is cloned from the original. */
  fromOriginal(k: number): boolean {
    return this.op[k] === DELETED || this.op[k] === MERGE || this.op[k] === MERGE_SECOND;
  }

  cls(k: number): BlockClass {
    const op = this.op[k];
    if (op === DELETED || op === MERGE_SECOND) return "deleted";
    if (op === INSERTED || op === SPLIT_SECOND) return "inserted";
    return "kept";
  }
}

/** The body's alignment, as compareInto's alignBlocks does it, read from the index. */
function planBody(O: Input, R: Input, ctx: Ctx): Plan {
  const plan = new Plan();
  const oN = O.blocks.count;
  const rN = R.blocks.count;
  const oSide = diskSide(O, true, ctx);
  const rSide = diskSide(R, false, ctx);
  const kindOf = (k: number): string => (plan.fromOriginal(k) ? O.kinds[plan.o[k]] : R.kinds[plan.r[k]]);
  let oAt = 0;
  let rAt = 0;
  for (const anchor of [...histogramMatch(O.fingerprints, R.fingerprints), { a: oN, b: rN }]) {
    if (anchor.a > oAt || anchor.b > rAt) {
      // As emitGap: a gap with no paragraph on either side must pair something.
      const unanchored = (plan.n === 0 || kindOf(plan.n - 1) !== "p") && (anchor.b >= rN || R.kinds[anchor.b] !== "p");
      const pairings = resolveGap(oSide, rSide, oAt, anchor.a, rAt, anchor.b, (oAt === 0 && rAt === 0) || unanchored);
      for (const op of gapOps(pairings, oAt, anchor.a, rAt, anchor.b)) {
        if (op.kind === "deleted") plan.push(DELETED, op.o, -1);
        else if (op.kind === "inserted") plan.push(INSERTED, -1, op.r);
        else if (op.kind === "split") {
          plan.push(SPLIT, op.o, op.r);
          plan.push(SPLIT_SECOND, -1, op.other);
        } else if (op.kind === "merge") {
          plan.push(MERGE, op.o, op.r);
          plan.push(MERGE_SECOND, op.other, -1);
        } else plan.push(MATCH, op.o, op.r);
      }
    }
    if (anchor.a < oN && anchor.b < rN) plan.push(O.hashes[anchor.a] === R.hashes[anchor.b] ? COPY : MATCH, anchor.a, anchor.b);
    oAt = anchor.a + 1;
    rAt = anchor.b + 1;
  }
  return plan;
}

/**
 * One side of every gap, answered from the index and the block store. A
 * windowed gap reads the blocks within GAP_WINDOW (50) of each position, so
 * the caches hold a few windows; a paragraph's split/merge keys run to tens
 * of kilobytes.
 */
function diskSide(input: Input, original: boolean, ctx: Ctx): GapSide {
  const texts = new Lru<string>(1024);
  const keys = new Lru<string[] | null>(256);
  return {
    kind: (i) => input.kinds[i],
    text: (i) => texts.get(i, () => (input.kinds[i] === "tbl" ? tableText(input, i) : input.texts.read(i))),
    section: (i) => input.sectionKeys.get(i) ?? "",
    // Body blocks are never table rows.
    cells: () => 0,
    keyParts: (i) =>
      keys.get(i, () => (input.kinds[i] === "p" && !input.endsSection[i] ? paragraphKeyParts(parseXml(input.blocks.read(i)), original, ctx) : null)),
  };
}

/**
 * How much of a table's text a gap scores it by. Its start tells one table
 * from another as well as the whole does, and the whole runs to megabytes.
 */
const TABLE_TEXT_CAP = 64 * 1024;

/** A table's paragraphText, from its rows, to about TABLE_TEXT_CAP. Its other children hold no text. */
function tableText(input: Input, i: number): string {
  const [from, to] = input.tables.get(i) as [number, number];
  let text = "";
  for (let k = from; k < to && text.length < TABLE_TEXT_CAP; k++) text += paragraphText(parseXml(input.rows.read(k)));
  return text;
}

/** The most recently used values, by index. */
class Lru<V> {
  private readonly map = new Map<number, V>();

  constructor(private readonly capacity: number) {}

  get(key: number, make: () => V): V {
    if (this.map.has(key)) {
      const value = this.map.get(key) as V;
      this.map.delete(key);
      this.map.set(key, value);
      return value;
    }
    const value = make();
    this.map.set(key, value);
    if (this.map.size > this.capacity) this.map.delete(this.map.keys().next().value as number);
    return value;
  }
}

// ---------------------------------------------------------------------------
// Emission: one entry at a time, its paragraph marks applied before it is written
// ---------------------------------------------------------------------------

const NOTE_REF = /<w:(?:footnote|endnote)Reference /;
const GLYPH = { ins: 1, del: 2 } as const;

/** Revision markup: XML without it has no revision to count or flatten. */
const REVISION = /<w:(?:ins|del|moveFrom|moveTo|rPrChange|pPrChange|tblPrChange|trPrChange|tcPrChange)[\s>/]/;

/** Revision counts by kind, all zero. */
function noRevisions(): Record<RevisionKind, number> {
  return {
    insertion: 0,
    deletion: 0,
    markInsertion: 0,
    markDeletion: 0,
    rowInsertion: 0,
    rowDeletion: 0,
    runFormat: 0,
    paragraphFormat: 0,
    tableFormat: 0,
    rowFormat: 0,
    cellFormat: 0,
    sectionFormat: 0,
  };
}

/** Add `el`'s revisions, as collectRevisions finds them, to `into`. */
function countRevisions(el: XmlElement, into: Record<RevisionKind, number>): void {
  const refs: RevisionRef[] = [];
  collectRevisionsIn(el, refs);
  for (const ref of refs) into[ref.kind]++;
}

/**
 * Where the result's body goes, a piece at a time: the result store and,
 * for a printable copy, its store, entry for entry. Revisions are counted
 * on the way, from each piece's tree.
 */
class BodySink {
  /** The result entries that hold note references. */
  readonly withNoteRefs: number[] = [];
  readonly revisions = noRevisions();
  /** The printable copy's last piece, held back for the mark the body's end decides. */
  private held: string | null = null;
  /** Inside a table whose own format changed: each row's paragraphs get the change bar. */
  tableBarred = false;

  constructor(
    readonly out: BlockStore,
    readonly printable: BlockStore | null,
    /** Relationship ids the body uses, when the save would otherwise drop parts it cannot see used. */
    readonly relIds: Set<string> | null,
  ) {}

  /** XML that no revision touched (a copied block, a table's tags), or one whose revisions are counted apart. */
  raw(xml: string, printableXml = xml): void {
    this.put(xml, printableXml);
  }

  /** An element as the compare left it. */
  element(el: XmlElement): void {
    const xml = serializeXml(el);
    const changed = REVISION.test(xml);
    if (changed) countRevisions(el, this.revisions);
    if (!this.printable || (!changed && !this.tableBarred)) return this.put(xml, xml);
    flatten(el);
    if (this.tableBarred) barParagraphs(el);
    this.put(xml, serializeXml(el));
  }

  /** The body is written: the final section's format change bars its last paragraph (a piece that is one). */
  finish(sectionChanged: boolean): void {
    if (!this.printable || this.held === null) return;
    let last = this.held;
    if (sectionChanged && /^<w:p[\s>]/.test(last)) {
      const body = parseXml(`<w:body>${last}</w:body>`);
      markLastParagraph(body);
      last = serializeXml(body.children[0]);
    }
    this.printable.add(last);
    this.held = null;
  }

  private put(xml: string, printableXml: string): void {
    const k = this.out.add(xml);
    if (this.printable) {
      if (this.held !== null) this.printable.add(this.held);
      this.held = printableXml;
    }
    if (NOTE_REF.test(xml)) this.withNoteRefs.push(k);
    if (this.relIds) for (const m of xml.matchAll(/="(rId\d+)"/g)) this.relIds.add(m[1]);
  }
}

function emitBody(plan: Plan, O: Input, R: Input, ctx: Ctx, sink: BodySink): void {
  const n = plan.n;
  // The mark pass is planned over the whole body first. Each entry takes at
  // most one pilcrow mark and one property swap, so the steps fit two columns.
  const glyph = new Uint8Array(n);
  const props = new Int32Array(n).fill(-1);
  let untracked = 0;
  const info = (k: number) => {
    const fromOriginal = plan.fromOriginal(k);
    const kind = fromOriginal ? O.kinds[plan.o[k]] : R.kinds[plan.r[k]];
    const ends = fromOriginal ? O.endsSection[plan.o[k]] : R.endsSection[plan.r[k]];
    return { cls: plan.cls(k), paragraph: kind === "p", endsSection: ends, marker: MARKERS.includes(kind) };
  };
  paragraphMarkSteps(n, info, (step) => {
    if (step.kind === "untracked") untracked++;
    else if (step.kind === "glyph") {
      if (glyph[step.at]) throw new Error("compare: two pilcrow marks for one paragraph");
      glyph[step.at] = GLYPH[step.glyph];
    } else {
      if (props[step.at] !== -1) throw new Error("compare: two property swaps for one paragraph");
      props[step.at] = step.follower;
    }
  });
  const noEntry = (): Entry => {
    throw new Error("compare: an untracked mark names no entry");
  };
  for (let i = 0; i < untracked; i++) applyMarkStep({ kind: "untracked" }, noEntry, () => undefined, ctx);

  const parseO = (i: number): XmlElement => parseXml(O.blocks.read(i));
  const parseR = (j: number): XmlElement => parseXml(R.blocks.read(j));
  // The second entry of a split or merge, emitted with the first.
  let second: Entry | null = null;

  /**
   * An entry's paragraph properties once it is emitted, for a property swap
   * that reads a later entry. Emitting keeps the live side's base properties
   * (alignParagraphProps), so the source paragraph answers.
   */
  const pPrOf = (f: number, side: "original" | "revised"): XmlElement | undefined => {
    if (second && f === emittedUpTo + 1) return side === "revised" ? child(second.el, "pPr") : originalProps(child(second.el, "pPr"));
    const op = plan.op[f];
    if (op === DELETED || op === MERGE_SECOND) return child(parseO(plan.o[f]), "pPr");
    // A matched, split or merged paragraph wears the original's properties
    // in a w:pPrChange once compared, when formatting is tracked.
    if (side === "original" && ctx.formatting && (op === MATCH || op === COPY || op === SPLIT || op === MERGE)) return child(parseO(plan.o[f]), "pPr");
    return child(parseR(plan.r[f]), "pPr");
  };
  let emittedUpTo = -1;

  for (let k = 0; k < n; k++) {
    const op = plan.op[k];
    const o = plan.o[k];
    const r = plan.r[k];
    emittedUpTo = k;
    // A table is on disk as its shell and rows, and takes no mark.
    if ((plan.fromOriginal(k) ? O.kinds[o] : R.kinds[r]) === "tbl") {
      if (op === DELETED) emitTableThrough(O, o, (tbl) => emitDeleted(tbl, ctx).el, sink);
      else if (op === INSERTED) emitTableThrough(R, r, (tbl) => emitInserted(tbl, ctx).el, sink);
      else emitMatchedTables(O, o, R, r, ctx, sink);
      continue;
    }
    let entry: Entry | null = null;
    let xml: string | null = null;
    if (op === COPY) {
      const revisedXml = R.blocks.read(r);
      // Same XML can still refer to a part that changed (a SmartArt diagram).
      const same = revisedXml === O.blocks.read(o) && !(/ r:\w+="/.test(revisedXml) && refersToChangedPart(parseXml(revisedXml), ctx));
      if (same && !glyph[k] && props[k] === -1) xml = revisedXml;
      else if (same) entry = { el: parseXml(revisedXml), cls: "kept" };
      else entry = emitMatched(parseO(o), parseR(r), ctx);
    } else if (op === MATCH) {
      entry = emitMatched(parseO(o), parseR(r), ctx);
    } else if (op === DELETED) {
      entry = emitDeleted(parseO(o), ctx);
    } else if (op === INSERTED) {
      entry = emitInserted(parseR(r), ctx);
    } else if (op === SPLIT || op === MERGE) {
      const other = op === SPLIT ? plan.r[k + 1] : plan.o[k + 1];
      [entry, second] = emitGapOp({ kind: op === SPLIT ? "split" : "merge", o, r, other }, parseO, parseR, ctx);
    } else {
      entry = second;
      second = null;
    }
    if (entry) {
      const self = entry;
      const at = (): Entry => self;
      if (glyph[k]) applyMarkStep({ kind: "glyph", at: k, glyph: glyph[k] === GLYPH.ins ? "ins" : "del" }, at, pPrOf, ctx);
      if (props[k] !== -1) applyMarkStep({ kind: "props", at: k, follower: props[k] }, at, pPrOf, ctx);
      sink.element(entry.el);
    } else {
      sink.raw(xml as string);
    }
  }
}

// ---------------------------------------------------------------------------
// Tables: the shell in memory, the rows read one at a time
// ---------------------------------------------------------------------------

/** A deleted or inserted table's children go through `emitWhole` about this much XML at a time. */
const TABLE_CHUNK_BYTES = 1 << 20;

/** A row's slot in a serialized shell. */
const ROW_SLOT_XML = serializeXml(ROW_SLOT);

/** `el`'s start tag and end tag, as serializeXml writes them around its children. */
function tagsOf(el: XmlElement): [string, string] {
  return serializeXml({ ...el, children: [ROW_SLOT] }).split(ROW_SLOT_XML) as [string, string];
}

/**
 * A table whose every child is emitted alike (emitDeleted, emitInserted):
 * the children go through `emitWhole` in a stand-in table, a chunk at a
 * time, and come out in document order. One chunk is what compareDocuments
 * does; a table over TABLE_CHUNK_BYTES whose rows share a relationship the
 * result lacks (a struck image) gets it adopted once per chunk, not once.
 */
function emitTableThrough(input: Input, i: number, emitWhole: (tbl: XmlElement) => XmlElement, sink: BodySink): void {
  const shell = parseXml(input.blocks.read(i));
  if (shell.children.length === 0) {
    sink.element(emitWhole(shell));
    return;
  }
  let [next] = input.tables.get(i) as [number, number];
  let chunk: XmlElement[] = [];
  let bytes = 0;
  let tail = "";
  const flush = (): void => {
    const done = emitWhole({ ...shell, children: chunk });
    if (!tail) {
      const tags = tagsOf(done);
      sink.raw(tags[0]);
      tail = tags[1];
    }
    for (const c of done.children) sink.element(c);
    chunk = [];
    bytes = 0;
  };
  for (const c of shell.children) {
    if (c.name === ROW_SLOT.name) {
      const xml = input.rows.read(next++);
      chunk.push(parseXml(xml));
      bytes += xml.length;
    } else chunk.push(c);
    if (bytes >= TABLE_CHUNK_BYTES) flush();
  }
  flush();
  sink.raw(tail);
}

/** Whether two tables write the same XML, as compareTable asks of a matched pair. */
function sameTables(O: Input, o: number, R: Input, r: number): boolean {
  const [oFrom, oTo] = O.tables.get(o) as [number, number];
  const [rFrom, rTo] = R.tables.get(r) as [number, number];
  if (oTo - oFrom !== rTo - rFrom) return false;
  const same = (a: string, b: string): boolean => a === b || sameXml(parseXml(a), parseXml(b));
  if (!same(O.blocks.read(o), R.blocks.read(r))) return false;
  for (let i = 0; i < oTo - oFrom; i++) if (!same(O.rows.read(oFrom + i), R.rows.read(rFrom + i))) return false;
  return true;
}

/** A matched table pair, as emitMatched compares it: the same table unchanged, else compareTable. */
function emitMatchedTables(O: Input, o: number, R: Input, r: number, ctx: Ctx, sink: BodySink): void {
  const shell = parseXml(R.blocks.read(r));
  const [rFrom, rTo] = R.tables.get(r) as [number, number];
  if (sameTables(O, o, R, r)) {
    // Unchanged: the revised table as it was, its rows back in their slots,
    // each row an entry of its own.
    const parts = serializeXml(shell).split(ROW_SLOT_XML);
    for (let i = 0; i < rTo - rFrom; i++) {
      if (parts[i]) sink.raw(parts[i]);
      sink.raw(R.rows.read(rFrom + i));
    }
    sink.raw(parts[rTo - rFrom]);
    return;
  }
  compareTableProps(parseXml(O.blocks.read(o)), shell, ctx);
  // compareRows puts the rows after the table's other children.
  shell.children = [...shell.children.filter((c) => c.name !== ROW_SLOT.name), ROW_SLOT];
  const [head, tail] = serializeXml(shell).split(ROW_SLOT_XML);
  // The table's own format change, here; its rows' as they are written.
  countRevisions(shell, sink.revisions);
  // The table's own format change bars every row's paragraphs (flatten).
  sink.tableBarred = !!child(child(shell, "tblPr"), "tblPrChange") || !!child(child(shell, "tblGrid"), "tblGridChange");
  sink.raw(head, sink.printable ? serializeXml(flatten(shell)).split(ROW_SLOT_XML)[0] : head);
  const [oFrom, oTo] = O.tables.get(o) as [number, number];
  emitComparedRows(O, oFrom, oTo, R, rFrom, rTo, ctx, sink);
  sink.tableBarred = false;
  sink.raw(tail);
}

/** compareRows over two tables' rows on disk: the result's rows, in order. */
function emitComparedRows(O: Input, oFrom: number, oTo: number, R: Input, rFrom: number, rTo: number, ctx: Ctx, sink: BodySink): void {
  const oN = oTo - oFrom;
  const rN = rTo - rFrom;
  const anchors = histogramMatch(O.rowFingerprints.slice(oFrom, oTo), R.rowFingerprints.slice(rFrom, rTo));
  const oSide = rowSide(O, oFrom);
  const rSide = rowSide(R, rFrom);
  const parseO = (i: number): XmlElement => parseXml(O.rows.read(oFrom + i));
  const parseR = (j: number): XmlElement => parseXml(R.rows.read(rFrom + j));
  const emitRows = (rows: XmlElement[]): void => {
    for (const row of rows) sink.element(row);
  };
  let oAt = 0;
  let rAt = 0;
  for (const anchor of [...anchors, { a: oN, b: rN }]) {
    const pairings = resolveGap(oSide, rSide, oAt, anchor.a, rAt, anchor.b).filter((p) => p.kind === "pair");
    let oi = oAt;
    let ri = rAt;
    for (const pairing of [...pairings, { kind: "pair" as const, o: anchor.a, r: anchor.b, other: -1 }]) {
      for (let i = oi; i < pairing.o; i++) emitRows([deletedRow(parseO(i), ctx)]);
      for (let j = ri; j < pairing.r; j++) emitRows([insertedRow(parseR(j), ctx)]);
      oi = pairing.o;
      ri = pairing.r;
      if (pairing.o >= anchor.a || pairing.r >= anchor.b) break;
      emitRows(compareRow(parseO(pairing.o), parseR(pairing.r), ctx));
      oi = pairing.o + 1;
      ri = pairing.r + 1;
    }
    if (anchor.a < oN && anchor.b < rN) emitRows(compareRow(parseO(anchor.a), parseR(anchor.b), ctx));
    oAt = anchor.a + 1;
    rAt = anchor.b + 1;
  }
}

/** One table's rows as a gap side, indexed from `base`: what treeSide answers for rows. */
function rowSide(input: Input, base: number): GapSide {
  const texts = new Lru<string>(1024);
  const row = (i: number): XmlElement => parseXml(input.rows.read(base + i));
  return {
    kind: (i) => input.rowKinds[base + i],
    text: (i) => texts.get(i, () => paragraphText(row(i))),
    // A row ends no section, and is no paragraph to split or merge.
    section: () => "",
    cells: (i) => row(i).children.filter((c) => localName(c.name) === "tc").length,
    keyParts: () => null,
  };
}

/** Per result piece: note kind -> the reference's place among that kind's references in the piece -> its new id. */
type NoteIdChanges = Map<number, Record<"footnote" | "endnote", Map<number, string>>>;

/**
 * giveStruckReferencesTheirOwnNotes for a body on disk: the references are
 * read from the pieces that hold them, the notes copied, and the new ids
 * returned for the writer to put in.
 */
function giveStruckReferencesTheirOwnNotes(body: BlockStore, withNoteRefs: number[], O: Input, R: Input, ctx: Ctx): NoteIdChanges {
  const changes: NoteIdChanges = new Map();
  for (const kind of ["footnote", "endnote"] as const) {
    const refs: { el: { attrs: Record<string, string> }; struck: boolean; inserted: boolean; block: number; place: number }[] = [];
    for (const k of withNoteRefs) {
      const found: NoteRef[] = [];
      // Walked as a child, so a piece that is a deleted row counts as struck.
      noteRefsIn({ name: "w:dxw-piece", attrs: {}, children: [parseXml(body.read(k))], text: "" }, kind, false, false, found);
      found.forEach((ref, place) => refs.push({ el: { attrs: { "w:id": ref.el.attrs["w:id"] } }, struck: ref.struck, inserted: ref.inserted, block: k, place }));
    }
    const repoint = (i: number, id: string): void => {
      refs[i].el.attrs["w:id"] = id;
      let block = changes.get(refs[i].block);
      if (!block) changes.set(refs[i].block, (block = { footnote: new Map(), endnote: new Map() }));
      block[kind].set(refs[i].place, id);
    };
    for (const [i, id] of settleLeftoverNotes(kind, O.shell, R.shell, ctx, refs, { original: O.notes[kind], merged: R.notes[kind] })) repoint(i, id);
    for (const [i, id] of copyNotesForStruckReferences(kind, R.shell, ctx, refs, R.notes[kind] ?? undefined)) repoint(i, id);
  }
  return changes;
}

// ---------------------------------------------------------------------------
// Output
// ---------------------------------------------------------------------------

async function writeResult(
  outPath: string,
  dir: string,
  O: Input,
  R: Input,
  body: BlockStore,
  noteIds: NoteIdChanges,
  relIds: Set<string> | null,
  printable: { path: string; beforeSectPr?: string; store: BlockStore } | null,
): Promise<void> {
  // document.xml: the revised shell's document, with the result's blocks
  // ahead of the body's own section properties.
  const shellBody = child(R.shell.docRoot, "body");
  if (!shellBody) throw new Error("document.xml has no w:body");
  const marker: XmlElement = { name: "w:dxw-body", attrs: {}, children: [], text: "" };
  shellBody.children.unshift(marker);
  const skeleton = serializeXml(R.shell.docRoot, true);
  shellBody.children.shift();
  const at = skeleton.indexOf("<w:dxw-body/>");
  const head = skeleton.slice(0, at);
  const tail = skeleton.slice(at + "<w:dxw-body/>".length);
  const docFile = join(dir, "result-document.xml");
  writeDocument(docFile, head, body, noteIds, tail);

  // Every other part comes from the shell's save, a part at a time. A save
  // leaves out chart parts the body it sees does not use, so the body's ids
  // stand in for it.
  const saveShell = (): Record<string, Uint8Array> => {
    if (relIds) shellBody.children.unshift({ name: "w:dxw-rels", attrs: Object.fromEntries([...relIds].map((id, i) => [`r:dxw${i}`, id])), children: [], text: "" });
    const files = R.shell.saveFiles();
    if (relIds) shellBody.children.shift();
    return files;
  };
  const writeNotes = (prefix: string, transform?: (note: XmlElement) => XmlElement): Map<string, string> => {
    const notes = new Map<string, string>();
    for (const part of notesOf(R)) {
      const file = join(dir, `${prefix}-${part.kind}s.xml`);
      part.write(file, transform);
      notes.set(part.part, file);
    }
    return notes;
  };
  await writePackage(outPath, docFile, saveShell(), writeNotes("result"), O, R);

  if (printable) {
    const printableFile = join(dir, "printable-document.xml");
    // The last section's page setup change, if any, as plain formatting.
    const printableTail = tail.replace(/<w:sectPrChange\b[\s\S]*?<\/w:sectPrChange>/, "");
    writeDocument(printableFile, head, printable.store, noteIds, (printable.beforeSectPr ?? "") + printableTail);
    // The headers, footers and notes carry revisions of their own: flattened
    // in the shell, now that the result is saved, and in the notes on disk.
    for (const root of R.shell.revisionRoots()) flatten(root);
    await writePackage(printable.path, printableFile, saveShell(), writeNotes("printable", flatten), O, R);
  }
}

/** A document.xml: `head`, the body's entries from `store` with their note ids settled, then `tail`. */
function writeDocument(path: string, head: string, store: BlockStore, noteIds: NoteIdChanges, tail: string): void {
  const fd = openSync(path, "w");
  try {
    const write = (text: string | Buffer): void => {
      const buf = typeof text === "string" ? Buffer.from(text, "utf8") : text;
      let done = 0;
      while (done < buf.length) done += writeSync(fd, buf, done, buf.length - done);
    };
    write(head);
    for (let k = 0; k < store.count; k++) {
      const change = noteIds.get(k);
      if (!change) {
        store.withBytes(k, write);
        continue;
      }
      const el = parseXml(store.read(k));
      for (const kind of ["footnote", "endnote"] as const) {
        const found: NoteRef[] = [];
        noteRefsIn(el, kind, false, false, found);
        for (const [place, id] of change[kind]) found[place].el.attrs["w:id"] = id;
      }
      write(serializeXml(el));
    }
    write(tail);
  } finally {
    closeSync(fd);
  }
}

/** The result's zip at `path`: `docFile` as document.xml, the notes parts from disk, the rest from the shell's save and the inputs. */
async function writePackage(path: string, docFile: string, files: Record<string, Uint8Array>, notes: Map<string, string>, O: Input, R: Input): Promise<void> {
  const writer = new ZipWriter(path);
  const names = Object.keys(files).sort((a, b) => Number(b === "[Content_Types].xml") - Number(a === "[Content_Types].xml"));
  for (const name of names) {
    const file = name === R.docPart ? docFile : notes.get(name);
    if (file) {
      await writer.addFile(name, file);
      continue;
    }
    const stub = stubSource(files[name]);
    if (stub) {
      const zip = stub.side === "original" ? O.zip : R.zip;
      const entry = zip.entries.get(stub.name);
      if (!entry) throw new Error(`compare: ${stub.name} is missing from the ${stub.side} document`);
      zip.copyInto(writer, entry, name);
    } else {
      writer.addBuffer(name, files[name]);
    }
  }
  writer.close();
}
