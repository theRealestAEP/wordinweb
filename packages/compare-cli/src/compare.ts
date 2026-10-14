import { closeSync, mkdtempSync, openSync, rmSync, statSync, statfsSync, writeSync } from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join, resolve } from "node:path";
import { unzipSync } from "fflate";
import { XmlElement, child, parseXml, serializeXml } from "../../core/src/xml.js";
import { histogramMatch } from "../../core/src/edit/compare/histogram.js";
import {
  BlockClass,
  CompareOptions,
  Ctx,
  Entry,
  GapSide,
  NoteRef,
  applyMarkStep,
  compareContainer,
  compareContext,
  compareNotes,
  copyNotesForStruckReferences,
  emitDeleted,
  emitGapOp,
  emitInserted,
  emitMatched,
  gapOps,
  noteRefsIn,
  pairHeadersFooters,
  paragraphKeyParts,
  paragraphMarkSteps,
  reportUncomparedParts,
  resolveGap,
} from "../../core/src/edit/compare/index.js";
import { Input, Side, readInput, stubSource } from "./input.js";
import { BlockStore } from "./store.js";
import { ZipReader, ZipWriter } from "./zip.js";

export interface CompareFilesOptions extends Omit<CompareOptions, "reuseRevised"> {
  /** Where the working files go. Default: the system's temporary directory. */
  tempDir?: string;
}

/**
 * Compare two .docx files on disk and write the result to `outPath`: the
 * revised document with every difference from the original as a tracked
 * change, as compareDocuments gives it.
 *
 * Memory does not grow with the size of the documents' bodies. Each body is
 * written to disk block by block with a small index (hashes, kinds, section
 * breaks); the alignment runs on the index, and the blocks are read back one
 * pair at a time to be compared and written out. Everything outside the
 * body (styles, numbering, headers, footers, notes, comments) is held in
 * memory, and each binary part is copied from the input zip unread.
 */
export async function compareFiles(originalPath: string, revisedPath: string, outPath: string, options: CompareFilesOptions): Promise<void> {
  for (const [which, path] of [["original", originalPath], ["revised", revisedPath]]) {
    const stat = statSync(path, { throwIfNoEntry: false });
    if (!stat) throw new Error(`${path}: the ${which} file does not exist`);
    if (!stat.isFile()) throw new Error(`${path}: the ${which} file is not a file`);
  }
  // The result is written while the inputs are still read.
  if ([originalPath, revisedPath].some((path) => resolve(path) === resolve(outPath))) throw new Error(`${outPath}: the result must not overwrite an input`);
  if (!statSync(dirname(resolve(outPath)), { throwIfNoEntry: false })?.isDirectory()) throw new Error(`${outPath}: the result's folder does not exist`);
  const dir = mkdtempSync(join(options.tempDir ?? tmpdir(), "docx-compare-"));
  const zips: ZipReader[] = [];
  const stores: BlockStore[] = [];
  try {
    const oZip = ZipReader.open(originalPath);
    zips.push(oZip);
    const rZip = ZipReader.open(revisedPath);
    zips.push(rZip);
    checkDiskSpace(dir, [oZip, rZip]);

    const O = await readNamed(oZip, "original", dir);
    stores.push(O.blocks, O.texts);
    const R = await readNamed(rZip, "revised", dir);
    stores.push(R.blocks, R.texts);

    const note = options.onNote ?? (() => {});
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
      compareNotes(kind, O.shell, R.shell, ctx, { original: O.noteRefs[kind], merged: R.noteRefs[kind] });
    }

    const plan = planBody(O, R, ctx);
    const body = new BlockStore(join(dir, "result-blocks"));
    stores.push(body);
    const emitted = emitBody(plan, O, R, ctx, body);
    const noteIds = giveStruckReferencesTheirOwnNotes(body, emitted.withNoteRefs, R, ctx);

    for (const [from, into] of headerFooterPairs) {
      const oRoot = O.shell.headerFooterRoot(from);
      const mRoot = R.shell.headerFooterRoot(into);
      if (oRoot && mRoot) compareContainer(oRoot, mRoot, { ...ctx, scope: { from, into } });
    }

    await writeResult(outPath, dir, O, R, body, noteIds, emitted.relIds);
  } finally {
    for (const store of stores) store.close();
    for (const zip of zips) zip.close();
    rmSync(dir, { recursive: true, force: true });
  }
}

/** readInput, with the input's path on any error that does not already name it. */
async function readNamed(zip: ZipReader, side: Side, dir: string): Promise<Input> {
  try {
    return await readInput(zip, side, dir);
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    throw new Error(message.startsWith(zip.path) ? message : `${zip.path}: ${message}`);
  }
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
  if (need > free) throw new Error(`not enough disk space in ${dir}: about ${mb(need)} needed, ${mb(free)} free`);
}

function mb(bytes: number): string {
  return `${Math.ceil(bytes / (1024 * 1024))} MB`;
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
    text: (i) => texts.get(i, () => input.texts.read(i)),
    section: (i) => input.sectionKeys.get(i) ?? "",
    // Body blocks are never table rows.
    cells: () => 0,
    keyParts: (i) =>
      keys.get(i, () => (input.kinds[i] === "p" && !input.endsSection[i] ? paragraphKeyParts(parseXml(input.blocks.read(i)), original, ctx) : null)),
  };
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

interface Emitted {
  /** The result blocks that hold note references. */
  withNoteRefs: number[];
  /** Relationship ids the result body uses, when the save would otherwise drop parts it cannot see used. */
  relIds: Set<string> | null;
}

function emitBody(plan: Plan, O: Input, R: Input, ctx: Ctx, out: BlockStore): Emitted {
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
    return { cls: plan.cls(k), paragraph: kind === "p", endsSection: ends };
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
  // A save drops a chart part whose relationship the body it sees does not use.
  const relIds = [O, R].some((input) => [...input.shell.documentRels.values()].some((rel) => rel.type.endsWith("/chart"))) ? new Set<string>() : null;
  const withNoteRefs: number[] = [];
  // The second entry of a split or merge, emitted with the first.
  let second: Entry | null = null;

  /**
   * An entry's paragraph properties once it is emitted, for a property swap
   * that reads a later entry. Emitting keeps the live side's base properties
   * (alignParagraphProps), so the source paragraph answers.
   */
  const pPrOf = (f: number): XmlElement | undefined => {
    if (second && f === emittedUpTo + 1) return child(second.el, "pPr");
    const op = plan.op[f];
    if (op === DELETED || op === MERGE_SECOND) return child(parseO(plan.o[f]), "pPr");
    return child(parseR(plan.r[f]), "pPr");
  };
  let emittedUpTo = -1;

  for (let k = 0; k < n; k++) {
    const op = plan.op[k];
    const o = plan.o[k];
    const r = plan.r[k];
    let entry: Entry | null = null;
    let xml: string | null = null;
    if (op === COPY) {
      const revisedXml = R.blocks.read(r);
      const same = revisedXml === O.blocks.read(o);
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
    emittedUpTo = k;
    if (entry) {
      const self = entry;
      const at = (): Entry => self;
      if (glyph[k]) applyMarkStep({ kind: "glyph", at: k, glyph: glyph[k] === GLYPH.ins ? "ins" : "del" }, at, pPrOf, ctx);
      if (props[k] !== -1) applyMarkStep({ kind: "props", at: k, follower: props[k] }, at, pPrOf, ctx);
      xml = serializeXml(entry.el);
    }
    const text = xml as string;
    out.add(text);
    if (NOTE_REF.test(text)) withNoteRefs.push(k);
    if (relIds) for (const m of text.matchAll(/="(rId\d+)"/g)) relIds.add(m[1]);
  }
  return { withNoteRefs, relIds };
}

/** Per result block: note kind -> the reference's place among that kind's references in the block -> its new id. */
type NoteIdChanges = Map<number, Record<"footnote" | "endnote", Map<number, string>>>;

/**
 * giveStruckReferencesTheirOwnNotes for a body on disk: the references are
 * read from the blocks that hold them, the notes copied, and the new ids
 * returned for the writer to put in.
 */
function giveStruckReferencesTheirOwnNotes(body: BlockStore, withNoteRefs: number[], R: Input, ctx: Ctx): NoteIdChanges {
  const changes: NoteIdChanges = new Map();
  for (const kind of ["footnote", "endnote"] as const) {
    const refs: { el: { attrs: Record<string, string> }; struck: boolean; block: number; place: number }[] = [];
    for (const k of withNoteRefs) {
      const found: NoteRef[] = [];
      noteRefsIn(parseXml(body.read(k)), kind, false, found);
      found.forEach((ref, place) => refs.push({ el: { attrs: { "w:id": ref.el.attrs["w:id"] } }, struck: ref.struck, block: k, place }));
    }
    for (const [i, id] of copyNotesForStruckReferences(kind, R.shell, ctx, refs)) {
      let block = changes.get(refs[i].block);
      if (!block) changes.set(refs[i].block, (block = { footnote: new Map(), endnote: new Map() }));
      block[kind].set(refs[i].place, id);
    }
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
  const docFile = join(dir, "result-document.xml");
  const fd = openSync(docFile, "w");
  try {
    const write = (text: string | Buffer): void => {
      const buf = typeof text === "string" ? Buffer.from(text, "utf8") : text;
      let done = 0;
      while (done < buf.length) done += writeSync(fd, buf, done, buf.length - done);
    };
    write(skeleton.slice(0, at));
    for (let k = 0; k < body.count; k++) {
      const change = noteIds.get(k);
      if (!change) {
        body.withBytes(k, write);
        continue;
      }
      const el = parseXml(body.read(k));
      for (const kind of ["footnote", "endnote"] as const) {
        const found: NoteRef[] = [];
        noteRefsIn(el, kind, false, found);
        for (const [place, id] of change[kind]) found[place].el.attrs["w:id"] = id;
      }
      write(serializeXml(el));
    }
    write(skeleton.slice(at + "<w:dxw-body/>".length));
  } finally {
    closeSync(fd);
  }

  // Every other part comes from the shell's save. A save leaves out chart
  // parts the body it sees does not use, so the body's ids stand in for it.
  if (relIds) shellBody.children.unshift({ name: "w:dxw-rels", attrs: Object.fromEntries([...relIds].map((id, i) => [`r:dxw${i}`, id])), children: [], text: "" });
  const files = unzipSync(R.shell.save());
  if (relIds) shellBody.children.shift();

  const writer = new ZipWriter(outPath);
  const names = Object.keys(files).sort((a, b) => Number(b === "[Content_Types].xml") - Number(a === "[Content_Types].xml"));
  for (const name of names) {
    if (name === R.docPart) {
      await writer.addFile(name, docFile);
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
