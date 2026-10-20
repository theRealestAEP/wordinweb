/**
 * A minimal reader of the Word 97-2003 (.doc, [MS-DOC]) structures the
 * review module needs: the FIB's fc/lcb pairs, the piece table, the text,
 * string tables, and the CHPX/PAPX FKPs mapped to character positions.
 *
 * Small on purpose: the foundation .doc reader has its own FIB, piece table
 * and FKP code; this file is what review.ts and its tests use until the two
 * are merged. The structure follows [MS-DOC] 2.5 (FIB), 2.8.35 (PlcPcd),
 * 2.9.33 (ChpxFkp), 2.9.174 (PapxFkp) and 2.2.4 (STTB), with the FKP walk
 * laid out as in Apache POI's HWPF (CHPBinTable/PAPBinTable, Apache 2.0).
 */

// ---------- bytes ----------

const u16 = (b: Uint8Array, o: number): number => b[o] | (b[o + 1] << 8);
const i16 = (b: Uint8Array, o: number): number => (u16(b, o) << 16) >> 16;
const u32 = (b: Uint8Array, o: number): number => (b[o] | (b[o + 1] << 8) | (b[o + 2] << 16) | (b[o + 3] << 24)) >>> 0;
const i32 = (b: Uint8Array, o: number): number => u32(b, o) | 0;
export { u16, i16, u32, i32 };

const cp1252 = new TextDecoder("windows-1252");
const utf16 = new TextDecoder("utf-16le");

// ---------- FIB ----------

/** Offsets of the FibRgFcLcb97 (and later) fc/lcb pairs the review reads. */
export const FcLcb = {
  plcfandRef: 4,
  plcfandTxt: 5,
  plcfBteChpx: 12,
  plcfBtePapx: 13,
  sttbfFfn: 15,
  clx: 33,
  grpXstAtnOwners: 36,
  sttbfAtnBkmk: 37,
  plcfAtnBkf: 42,
  plcfAtnBkl: 43,
  sttbfRMark: 51,
  atrdExtra: 112,
} as const;

export interface DocFib {
  nFib: number;
  encrypted: boolean;
  /** Which table stream holds the tables: "1Table" when fWhichTblStm is set. */
  tableStream: "0Table" | "1Table";
  /** FibRgLw97 character counts: main text, footnotes, headers, macros,
   * comments, endnotes, text boxes, header text boxes. */
  ccpText: number;
  ccpFtn: number;
  ccpHdd: number;
  ccpMcr: number;
  ccpAtn: number;
  ccpEdn: number;
  ccpTxbx: number;
  ccpHdrTxbx: number;
  /** The fc/lcb pair at an index of FcLcb, or [0, 0] when the FIB is too short. */
  fcLcb(index: number): [number, number];
}

const FIB_RG_FC_LCB = 0x9a; // base (32) + csw (2) + fibRgW (28) + cslw (2) + fibRgLw (88) + cbRgFcLcb (2)

export function readFib(wordDocument: Uint8Array): DocFib {
  if (wordDocument.length < 0x9a || u16(wordDocument, 0) !== 0xa5ec) throw new Error("not a Word binary document");
  const flags = u16(wordDocument, 0x0a);
  const cbRgFcLcb = u16(wordDocument, 0x98);
  const lw = (k: number) => u32(wordDocument, 0x4c + 4 * k);
  return {
    nFib: u16(wordDocument, 2),
    encrypted: (flags & 0x0100) !== 0,
    tableStream: flags & 0x0200 ? "1Table" : "0Table",
    ccpText: lw(0),
    ccpFtn: lw(1),
    ccpHdd: lw(2),
    ccpMcr: lw(3),
    ccpAtn: lw(4),
    ccpEdn: lw(5),
    ccpTxbx: lw(6),
    ccpHdrTxbx: lw(7),
    fcLcb(index) {
      if (index >= cbRgFcLcb) return [0, 0];
      const o = FIB_RG_FC_LCB + index * 8;
      if (o + 8 > wordDocument.length) return [0, 0];
      return [u32(wordDocument, o), u32(wordDocument, o + 4)];
    },
  };
}

/** The bytes of a table-stream structure named by an fc/lcb pair, clipped to the stream. */
export function tableBytes(table: Uint8Array, fib: DocFib, index: number): Uint8Array {
  const [fc, lcb] = fib.fcLcb(index);
  if (!lcb || fc >= table.length) return new Uint8Array(0);
  return table.subarray(fc, Math.min(table.length, fc + lcb));
}

// ---------- piece table ----------

export interface DocPiece {
  cpStart: number;
  cpEnd: number;
  /** Byte offset of the text in the WordDocument stream. */
  fc: number;
  /** 8-bit (Windows-1252) text; else UTF-16LE. */
  compressed: boolean;
  /** The Pcd's Prm: a single sprm (Prm0) or an index into the Clx's grpprls (Prm1). */
  prm: number;
}

export interface DocPieceTable {
  pieces: DocPiece[];
  /** Clx.RgPrc grpprls, for Prm1. */
  prcs: Uint8Array[];
}

export function readPieceTable(table: Uint8Array, fib: DocFib): DocPieceTable {
  const clx = tableBytes(table, fib, FcLcb.clx);
  const prcs: Uint8Array[] = [];
  const pieces: DocPiece[] = [];
  let i = 0;
  while (i < clx.length) {
    const clxt = clx[i];
    if (clxt === 1) {
      const cb = u16(clx, i + 1);
      prcs.push(clx.subarray(i + 3, i + 3 + cb));
      i += 3 + cb;
    } else if (clxt === 2) {
      const lcb = u32(clx, i + 1);
      const n = Math.floor((lcb - 4) / 12);
      const base = i + 5;
      for (let k = 0; k < n; k++) {
        const o = base + (n + 1) * 4 + k * 8;
        const fcc = u32(clx, o + 2);
        const compressed = (fcc & 0x40000000) !== 0;
        const fc = fcc & 0x3fffffff;
        pieces.push({
          cpStart: u32(clx, base + k * 4),
          cpEnd: u32(clx, base + (k + 1) * 4),
          fc: compressed ? fc >>> 1 : fc,
          compressed,
          prm: u16(clx, o + 6),
        });
      }
      break;
    } else break;
  }
  return { pieces, prcs };
}

/** The text at [start, end) in CP space, as the stream holds it (control characters kept). */
export function textOf(wordDocument: Uint8Array, pt: DocPieceTable, start: number, end: number): string {
  let out = "";
  for (const p of pt.pieces) {
    const lo = Math.max(p.cpStart, start);
    const hi = Math.min(p.cpEnd, end);
    if (lo >= hi) continue;
    const a = p.fc + (lo - p.cpStart) * (p.compressed ? 1 : 2);
    const b = p.fc + (hi - p.cpStart) * (p.compressed ? 1 : 2);
    const bytes = wordDocument.subarray(a, Math.min(b, wordDocument.length));
    out += p.compressed ? cp1252.decode(bytes) : utf16.decode(bytes);
  }
  return out;
}

// ---------- STTB, Xst ----------

/** An STTB's strings ([MS-DOC] 2.2.4 STTB) with each string's extra data. */
export function readSttb(b: Uint8Array): { strings: string[]; extra: Uint8Array[] } {
  const strings: string[] = [];
  const extra: Uint8Array[] = [];
  if (b.length < 4) return { strings, extra };
  const wide = u16(b, 0) === 0xffff;
  let i = wide ? 2 : 0;
  const n = u16(b, i);
  const cbExtra = u16(b, i + 2);
  i += 4;
  for (let k = 0; k < n && i + (wide ? 2 : 1) <= b.length; k++) {
    const cch = wide ? u16(b, i) : b[i];
    i += wide ? 2 : 1;
    const cb = wide ? cch * 2 : cch;
    const raw = b.subarray(i, i + cb);
    strings.push(wide ? utf16.decode(raw) : cp1252.decode(raw));
    i += cb;
    extra.push(b.subarray(i, i + cbExtra));
    i += cbExtra;
  }
  return { strings, extra };
}

/** A sequence of Xst (cch, then UTF-16 characters), as GrpXstAtnOwners holds. */
export function readXstGroup(b: Uint8Array): string[] {
  const out: string[] = [];
  let i = 0;
  while (i + 2 <= b.length) {
    const cch = u16(b, i);
    i += 2;
    out.push(utf16.decode(b.subarray(i, i + cch * 2)));
    i += cch * 2;
  }
  return out;
}

/** The CPs of a PLC whose data elements are cbData bytes each. */
export function plcCps(b: Uint8Array, cbData: number): number[] {
  const n = cbData === 0 ? Math.floor(b.length / 4) : Math.floor((b.length - 4) / (4 + cbData)) + 1;
  const cps: number[] = [];
  for (let k = 0; k < n; k++) cps.push(u32(b, k * 4));
  return cps;
}

// ---------- sprms ----------

export interface Prl {
  sprm: number;
  operand: Uint8Array;
}

/** The size of a sprm's operand ([MS-DOC] 2.2.5.1 Sprm.spra; the two
 * variable-size exceptions are sprmTDefTable and sprmPChgTabs). */
export function operandSize(sprm: number, grpprl: Uint8Array, at: number): number {
  switch (sprm >> 13) {
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
  }
  if (sprm === 0xd608) return u16(grpprl, at) + 1; // sprmTDefTable: cb counts the rest plus one
  if (sprm === 0xc615) {
    // sprmPChgTabs: cb, or 255 then the two counted arrays
    const cb = grpprl[at];
    if (cb !== 255) return cb + 1;
    const cTabsDel = grpprl[at + 1];
    const cTabsAdd = grpprl[at + 2 + cTabsDel * 4];
    return 1 + 1 + cTabsDel * 4 + 1 + cTabsAdd * 3;
  }
  return grpprl[at] + 1;
}

export function readPrls(grpprl: Uint8Array): Prl[] {
  const out: Prl[] = [];
  let i = 0;
  while (i + 2 <= grpprl.length) {
    const sprm = u16(grpprl, i);
    i += 2;
    const n = operandSize(sprm, grpprl, i);
    out.push({ sprm, operand: grpprl.subarray(i, i + n) });
    i += n;
  }
  return out;
}

// ---------- FKPs ----------

/** A run of text with one CHPX, or a paragraph with one PAPX, in FC space. */
export interface FkpRun {
  fcStart: number;
  fcEnd: number;
  grpprl: Uint8Array;
  /** PAPX only: the paragraph's style index from GrpPrlAndIstd. */
  istd?: number;
}

/** Every CHPX (papx=false) or PAPX (papx=true) run the bin table names, in FC order. */
export function readFkpRuns(wordDocument: Uint8Array, table: Uint8Array, fib: DocFib, papx: boolean): FkpRun[] {
  const plc = tableBytes(table, fib, papx ? FcLcb.plcfBtePapx : FcLcb.plcfBteChpx);
  const n = Math.floor((plc.length - 4) / 8);
  const runs: FkpRun[] = [];
  for (let k = 0; k < n; k++) {
    const pn = u32(plc, (n + 1) * 4 + k * 4) & 0x3fffff;
    const page = wordDocument.subarray(pn * 512, (pn + 1) * 512);
    if (page.length < 512) continue;
    const crun = page[511];
    for (let r = 0; r < crun; r++) {
      const fcStart = u32(page, r * 4);
      const fcEnd = u32(page, r * 4 + 4);
      if (papx) {
        const off = page[(crun + 1) * 4 + r * 13] * 2;
        if (off === 0) {
          runs.push({ fcStart, fcEnd, grpprl: new Uint8Array(0), istd: 0 });
          continue;
        }
        let cb = page[off];
        let grp: Uint8Array;
        if (cb === 0) {
          cb = page[off + 1] * 2;
          grp = page.subarray(off + 2, off + 2 + cb);
        } else {
          grp = page.subarray(off + 1, off + 1 + cb * 2 - 1);
        }
        runs.push({ fcStart, fcEnd, grpprl: grp.subarray(2), istd: grp.length >= 2 ? u16(grp, 0) : 0 });
      } else {
        const off = page[(crun + 1) * 4 + r] * 2;
        const grp = off === 0 ? new Uint8Array(0) : page.subarray(off + 1, off + 1 + page[off]);
        runs.push({ fcStart, fcEnd, grpprl: grp });
      }
    }
  }
  return runs;
}

/** An FKP run mapped onto a piece: the CPs it covers there. */
export interface CpRun {
  start: number;
  end: number;
  run: FkpRun;
  piece: DocPiece;
  /** This fragment holds the run's last character (a run split over pieces has one such). */
  tail: boolean;
}

/** The FKP runs projected through the piece table onto CP space, in CP order.
 * A run that spans two pieces appears once per piece. */
export function runsByCp(runs: FkpRun[], pt: DocPieceTable): CpRun[] {
  const out: CpRun[] = [];
  for (const p of pt.pieces) {
    const w = p.compressed ? 1 : 2;
    const fcA = p.fc;
    const fcB = p.fc + (p.cpEnd - p.cpStart) * w;
    // first run that ends after the piece starts (runs are sorted by fc)
    let lo = 0;
    let hi = runs.length;
    while (lo < hi) {
      const mid = (lo + hi) >> 1;
      if (runs[mid].fcEnd <= fcA) lo = mid + 1;
      else hi = mid;
    }
    for (let k = lo; k < runs.length && runs[k].fcStart < fcB; k++) {
      const run = runs[k];
      const a = Math.max(run.fcStart, fcA);
      const b = Math.min(run.fcEnd, fcB);
      if (a >= b) continue;
      out.push({
        start: p.cpStart + Math.floor((a - fcA) / w),
        end: p.cpStart + Math.ceil((b - fcA) / w),
        run,
        piece: p,
        tail: run.fcEnd <= fcB,
      });
    }
  }
  return out;
}

/** The sprms a piece's Prm applies on top of the FKP properties
 * ([MS-DOC] 2.9.214 Prm; Prm0's isprm table is reduced to the revision sprms). */
export function prmPrls(prm: number, prcs: Uint8Array[]): Prl[] {
  if (prm === 0) return [];
  if (prm & 1) {
    const grp = prcs[prm >> 1];
    return grp ? readPrls(grp) : [];
  }
  const isprm = (prm >> 1) & 0x7f;
  const val = (prm >> 8) & 0xff;
  const sprm = PRM0_SPRMS[isprm];
  return sprm === undefined ? [] : [{ sprm, operand: new Uint8Array([val]) }];
}

const PRM0_SPRMS: Record<number, number> = {
  0x41: 0x0800, // sprmCFRMarkDel
  0x42: 0x0801, // sprmCFRMarkIns
};

// ---------- DTTM ----------

/** A DTTM ([MS-DOC] 2.9.65) as an ISO 8601 timestamp the way Word writes
 * it into WordprocessingML (minute precision, "Z"); undefined when the day
 * or month is zero, which the spec says to ignore. */
export function dttmToIso(dttm: number): string | undefined {
  const mint = dttm & 0x3f;
  const hr = (dttm >>> 6) & 0x1f;
  const dom = (dttm >>> 11) & 0x1f;
  const mon = (dttm >>> 16) & 0x0f;
  const yr = ((dttm >>> 20) & 0x1ff) + 1900;
  if (dom === 0 || mon === 0) return undefined;
  const p2 = (n: number) => String(n).padStart(2, "0");
  return `${yr}-${p2(mon)}-${p2(dom)}T${p2(hr)}:${p2(mint)}:00Z`;
}
