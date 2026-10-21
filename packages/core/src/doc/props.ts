// Direct formatting ([MS-DOC] 2.4.6.1, 2.4.6.2): the character runs (CHPX)
// and paragraph properties (PAPX) kept in 512-byte formatted-disk pages
// (FKPs) of the WordDocument stream, found through the bin tables
// (PlcBteChpx, PlcBtePapx) by the byte offset (FC) of the text. The piece
// table maps a CP to its FC.
import { u16, u32 } from "./bytes.js";
import { FC, Fib, tableBlob } from "./fib.js";
import { PieceTable } from "./text.js";
import { word6Grpprl } from "./word6.js";

export interface CharRun {
  cp: number;
  end: number;
  /** The run's Chpx sprms, then those of its piece's Prm. */
  grpprl: Uint8Array;
  /** The run's text is 8-bit in the file (a symbol font's bytes map to U+F0xx). */
  compressed: boolean;
}

export interface ParaProps {
  istd: number;
  /** The paragraph's sprms (sprmPHugePapx resolved), then those of its piece's Prm. */
  grpprl: Uint8Array;
}

interface ChpxPage {
  rgfc: number[];
  runs: (Uint8Array | null)[];
}

interface PapxPage {
  rgfc: number[];
  entries: ({ istd: number; grpprl: Uint8Array } | null)[];
}

interface BinTable {
  fcs: number[];
  pns: number[];
}

const EMPTY = new Uint8Array(0);

/** The sprm, operand and total length of each Prl in `grpprl`, in order. */
export function* sprms(grpprl: Uint8Array): Generator<{ sprm: number; operand: Uint8Array }> {
  let i = 0;
  while (i + 2 <= grpprl.length) {
    const sprm = u16(grpprl, i);
    i += 2;
    const n = operandLength(sprm, grpprl, i);
    if (i + n > grpprl.length) return;
    yield { sprm, operand: grpprl.subarray(i, i + n) };
    i += n;
  }
}

/** The operand length of `sprm` whose operand starts at `at` ([MS-DOC] 2.2.5.1). */
export function operandLength(sprm: number, grpprl: Uint8Array, at: number): number {
  switch (sprm >>> 13) {
    case 0:
    case 1:
      return 1;
    case 2:
    case 4:
    case 5:
      return 2;
    case 3:
      return 4;
    case 7:
      return 3;
    default: {
      if (sprm === 0xd608) return u16(grpprl, at) + 1;
      const cb = grpprl[at];
      if (sprm === 0xc615 && cb === 0xff) {
        // A sprmPChgTabs too long for one byte: its parts give the length.
        const cDel = grpprl[at + 1];
        const cAdd = grpprl[at + 2 + cDel * 4];
        return 1 + 1 + cDel * 4 + 1 + cAdd * 3;
      }
      return cb + 1;
    }
  }
}

function concat(a: Uint8Array, b: Uint8Array | null): Uint8Array {
  if (!b || b.length === 0) return a;
  if (a.length === 0) return b;
  const out = new Uint8Array(a.length + b.length);
  out.set(a);
  out.set(b, a.length);
  return out;
}

export class PropertyTables {
  private readonly chpxBins: BinTable;
  private readonly papxBins: BinTable;
  private readonly chpxPages = new Map<number, ChpxPage>();
  private readonly papxPages = new Map<number, PapxPage>();

  constructor(
    private readonly word: Uint8Array,
    table: Uint8Array,
    private readonly data: Uint8Array | null,
    fib: Fib,
    private readonly pieces: PieceTable,
  ) {
    this.word6 = !!fib.word6;
    this.chpxBins = binTable(tableBlob(fib, table, FC.plcfBteChpx), this.word6);
    this.papxBins = binTable(tableBlob(fib, table, FC.plcfBtePapx), this.word6);
    if (fib.word6) {
      completeBins(this.chpxBins, word, fib.word6.cpnBteChp);
      completeBins(this.papxBins, word, fib.word6.cpnBtePap);
    }
  }

  /** A Word 6.0/95 file: 2-byte bin table entries, 7-byte BXs, and Word 6 sprms (word6.ts). */
  private readonly word6: boolean;

  /** The character runs covering [start, end), clipped to it, in order. */
  charRuns(start: number, end: number): CharRun[] {
    const out: CharRun[] = [];
    for (let i = this.pieces.pieceIndex(start); i < this.pieces.pieces.length; i++) {
      const piece = this.pieces.pieces[i];
      if (piece.cp >= end) break;
      const lo = Math.max(piece.cp, start);
      const hi = Math.min(piece.end, end);
      if (lo >= hi) continue;
      const width = piece.compressed ? 1 : 2;
      const pieceFc = piece.compressed ? piece.fc / 2 : piece.fc;
      const prm = this.pieces.prmGrpprl(piece.prm);
      let fc = pieceFc + (lo - piece.cp) * width;
      const fcEnd = pieceFc + (hi - piece.cp) * width;
      while (fc < fcEnd) {
        const found = this.chpxAt(fc);
        // Text no page's run covers keeps the default properties up to the next page's range.
        const runEnd = Math.min(found ? found.end : nextBinFc(this.chpxBins, fc), fcEnd);
        const cpA = piece.cp + Math.floor((fc - pieceFc) / width);
        const cpB = piece.cp + Math.ceil((runEnd - pieceFc) / width);
        const grpprl = concat(found?.grpprl ?? EMPTY, prm);
        const last = out[out.length - 1];
        // Runs never span pieces: a piece's text may be 8-bit where the next is 16-bit.
        if (last && last.end === cpA && last.compressed === piece.compressed && sameBytes(last.grpprl, grpprl)) last.end = cpB;
        else out.push({ cp: cpA, end: cpB, grpprl, compressed: piece.compressed });
        if (runEnd <= fc) break;
        fc = runEnd;
      }
    }
    return out;
  }

  /** The direct paragraph properties of the paragraph whose mark is at `cpMark`. */
  paragraphProps(cpMark: number): ParaProps {
    const piece = this.pieces.pieces[this.pieces.pieceIndex(cpMark)];
    const fc = this.pieces.fcOf(cpMark);
    const prm = this.pieces.prmGrpprl(piece.prm);
    const entry = this.papxAt(fc);
    let grpprl = entry?.grpprl ?? EMPTY;
    // A paragraph with more properties than a page holds keeps them in the Data stream.
    for (const { sprm, operand } of sprms(grpprl)) {
      if (sprm === 0x6646 && this.data) {
        const at = u32(operand, 0);
        if (at + 2 <= this.data.length) {
          const cb = u16(this.data, at);
          grpprl = this.data.subarray(at + 2, at + 2 + cb);
        }
        break;
      }
    }
    return { istd: entry?.istd ?? 0, grpprl: concat(grpprl, prm) };
  }

  private chpxAt(fc: number): { end: number; grpprl: Uint8Array | null } | null {
    const page = this.chpxPage(fc);
    if (!page) return null;
    const j = runIndex(page.rgfc, fc);
    if (j < 0) return null;
    return { end: page.rgfc[j + 1], grpprl: page.runs[j] };
  }

  private papxAt(fc: number): { istd: number; grpprl: Uint8Array } | null {
    const page = this.papxPage(fc);
    if (!page) return null;
    const j = runIndex(page.rgfc, fc);
    return j < 0 ? null : page.entries[j];
  }

  private chpxPage(fc: number): ChpxPage | null {
    const pn = binLookup(this.chpxBins, fc);
    if (pn === undefined) return null;
    let page = this.chpxPages.get(pn);
    if (!page) {
      page = readChpxPage(this.word, pn);
      if (this.word6) page.runs = page.runs.map((grpprl) => grpprl && word6Grpprl(grpprl));
      this.chpxPages.set(pn, page);
    }
    return page;
  }

  private papxPage(fc: number): PapxPage | null {
    const pn = binLookup(this.papxBins, fc);
    if (pn === undefined) return null;
    let page = this.papxPages.get(pn);
    if (!page) {
      page = readPapxPage(this.word, pn, this.word6);
      if (this.word6) page.entries = page.entries.map((entry) => entry && { istd: entry.istd, grpprl: word6Grpprl(entry.grpprl) });
      this.papxPages.set(pn, page);
    }
    return page;
  }
}

function sameBytes(a: Uint8Array, b: Uint8Array): boolean {
  if (a === b) return true;
  if (a.length !== b.length) return false;
  for (let i = 0; i < a.length; i++) if (a[i] !== b[i]) return false;
  return true;
}

/** A PlcBteChpx or PlcBtePapx; a Word 6 one has 2-byte page numbers. */
function binTable(plc: Uint8Array, word6: boolean): BinTable {
  const cbBte = word6 ? 2 : 4;
  const n = Math.floor((plc.length - 4) / (4 + cbBte));
  const fcs: number[] = [];
  const pns: number[] = [];
  if (n < 0) return { fcs, pns };
  for (let i = 0; i <= n; i++) fcs.push(u32(plc, i * 4));
  for (let i = 0; i < n; i++) pns.push(word6 ? u16(plc, (n + 1) * 4 + i * 2) : u32(plc, (n + 1) * 4 + i * 4) & 0x3fffff);
  return { fcs, pns };
}

/**
 * A Word 6 bin table can list only the first of the file's `count` pages
 * (the Word 6 format's "incomplete plcfbte"): the others follow the last
 * one listed, and each starts at its own first FC.
 */
function completeBins(bins: BinTable, word: Uint8Array, count: number): void {
  if (bins.pns.length === 0) return;
  for (let pn = bins.pns[bins.pns.length - 1] + 1; bins.pns.length < count && (pn + 1) * 512 <= word.length; pn++) {
    bins.fcs[bins.fcs.length - 1] = u32(word, pn * 512);
    bins.pns.push(pn);
    // The page's last FC ends its range.
    bins.fcs.push(u32(word, pn * 512 + word[pn * 512 + 511] * 4));
  }
}

/** The page number for `fc`: that of the largest aFC at or below it. A fc past the last range gets the last page. */
function binLookup(bins: BinTable, fc: number): number | undefined {
  if (bins.pns.length === 0) return undefined;
  let lo = 0;
  let hi = bins.pns.length - 1;
  while (lo < hi) {
    const mid = (lo + hi + 1) >> 1;
    if (bins.fcs[mid] <= fc) lo = mid;
    else hi = mid - 1;
  }
  return bins.pns[lo];
}

/** The first aFC above `fc`, or Infinity past the last. */
function nextBinFc(bins: BinTable, fc: number): number {
  for (const bound of bins.fcs) if (bound > fc) return bound;
  return Infinity;
}

/** The largest j with rgfc[j] <= fc < rgfc[j+1], or -1 when fc is outside the page's runs. */
function runIndex(rgfc: number[], fc: number): number {
  const n = rgfc.length - 1;
  if (n <= 0 || fc < rgfc[0] || fc >= rgfc[n]) return -1;
  let lo = 0;
  let hi = n - 1;
  while (lo < hi) {
    const mid = (lo + hi + 1) >> 1;
    if (rgfc[mid] <= fc) lo = mid;
    else hi = mid - 1;
  }
  return lo;
}

function readChpxPage(word: Uint8Array, pn: number): ChpxPage {
  const page = word.subarray(pn * 512, pn * 512 + 512);
  if (page.length < 512) return { rgfc: [], runs: [] };
  const crun = page[511];
  const rgfc: number[] = [];
  for (let i = 0; i <= crun; i++) rgfc.push(u32(page, i * 4));
  const runs: (Uint8Array | null)[] = [];
  for (let i = 0; i < crun; i++) {
    const off = page[(crun + 1) * 4 + i] * 2;
    if (off === 0) {
      runs.push(null);
      continue;
    }
    const cb = page[off];
    runs.push(page.subarray(off + 1, off + 1 + cb));
  }
  return { rgfc, runs };
}

/** A PAPX FKP. A Word 6 one has 7-byte BXs, and its PAPX's count of words covers all of the bytes after it. */
function readPapxPage(word: Uint8Array, pn: number, word6: boolean): PapxPage {
  const page = word.subarray(pn * 512, pn * 512 + 512);
  if (page.length < 512) return { rgfc: [], entries: [] };
  const cpara = page[511];
  const rgfc: number[] = [];
  for (let i = 0; i <= cpara; i++) rgfc.push(u32(page, i * 4));
  const entries: ({ istd: number; grpprl: Uint8Array } | null)[] = [];
  for (let i = 0; i < cpara; i++) {
    const off = page[(cpara + 1) * 4 + i * (word6 ? 7 : 13)] * 2;
    if (off === 0) {
      entries.push(null);
      continue;
    }
    let cb = page[off];
    let at = off + 1;
    let size: number;
    if (word6) {
      size = cb * 2;
    } else if (cb === 0) {
      cb = page[at];
      at++;
      size = cb * 2;
    } else {
      size = cb * 2 - 1;
    }
    if (size < 2) {
      entries.push(null);
      continue;
    }
    entries.push({ istd: u16(page, at), grpprl: page.subarray(at + 2, at + size) });
  }
  return { rgfc, entries };
}
