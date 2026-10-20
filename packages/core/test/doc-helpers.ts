/**
 * Just enough of a .doc reader for the drawings tests: the CFB streams, the
 * FIB fields that locate OfficeArt and shape anchors, and the CHPX scan that
 * finds every sprmCPicLocation. The real reader lives in src/doc; this stays
 * private to the tests.
 */
import * as CFB from "cfb";
import { parsePlcfSpa, type DocFspa } from "../src/doc/drawings.js";

export interface DocStreams {
  wordDocument: Uint8Array;
  table: Uint8Array;
  data: Uint8Array;
}

export function openDoc(bytes: Uint8Array): DocStreams {
  const cfb = CFB.read(bytes, { type: "array" });
  const stream = (name: string): Uint8Array => {
    const entry = CFB.find(cfb, name);
    return entry?.content ? new Uint8Array(entry.content as ArrayLike<number>) : new Uint8Array(0);
  };
  const wordDocument = stream("WordDocument");
  const flags = new DataView(wordDocument.buffer, wordDocument.byteOffset).getUint16(0x0a, true);
  const table = stream((flags & 0x0200) !== 0 ? "1Table" : "0Table");
  return { wordDocument, table, data: stream("Data") };
}

/** The (fc, lcb) pair at index `i` of FibRgFcLcb97. */
export function fibFcLcb(wordDocument: Uint8Array, i: number): { fc: number; lcb: number } {
  const v = new DataView(wordDocument.buffer, wordDocument.byteOffset);
  const at = 154 + i * 8;
  return { fc: v.getUint32(at, true), lcb: v.getUint32(at + 4, true) };
}

export const FIB = { plcfBteChpx: 12, plcSpaMom: 40, plcSpaHdr: 41, dggInfo: 50, plcfTxbxTxt: 56 } as const;

export function officeArtRange(streams: DocStreams): { fc: number; lcb: number } {
  return fibFcLcb(streams.wordDocument, FIB.dggInfo);
}

export function shapeAnchors(streams: DocStreams, which: "main" | "header" = "main"): { cp: number; fspa: DocFspa }[] {
  const { fc, lcb } = fibFcLcb(streams.wordDocument, which === "main" ? FIB.plcSpaMom : FIB.plcSpaHdr);
  return lcb ? parsePlcfSpa(streams.table, fc, lcb) : [];
}

/** Operand size of a sprm from its spra bits; -1 for the variable-length kind. */
function sprmOperandSize(sprm: number): number {
  switch (sprm >> 13) {
    case 0: case 1: return 1;
    case 2: case 4: case 5: return 2;
    case 3: return 4;
    case 7: return 3;
    default: return -1;
  }
}

/**
 * Every sprmCPicLocation (0x6A03) operand in the document's CHPX pages whose
 * run is a special character (sprmCFSpec), each offset once in file order.
 * Word leaves sprmCPicLocation in other runs too; those carry no picture.
 */
export function pictureLocations(streams: DocStreams): { sprm: number; offset: number; fcStart: number; fcEnd: number; ole: boolean }[] {
  const { wordDocument, table } = streams;
  const { fc, lcb } = fibFcLcb(wordDocument, FIB.plcfBteChpx);
  const out: { sprm: number; offset: number; fcStart: number; fcEnd: number; ole: boolean }[] = [];
  if (!lcb) return out;
  const offsets = new Set<number>();
  const tv = new DataView(table.buffer, table.byteOffset);
  const count = Math.floor((lcb - 4) / 8);
  const seen = new Set<number>();
  for (let i = 0; i < count; i++) {
    const pn = tv.getUint32(fc + (count + 1) * 4 + i * 4, true) & 0x3fffff;
    if (seen.has(pn)) continue;
    seen.add(pn);
    const page = pn * 512;
    if (page + 512 > wordDocument.length) continue;
    const pv = new DataView(wordDocument.buffer, wordDocument.byteOffset + page, 512);
    const crun = wordDocument[page + 511];
    for (let r = 0; r < crun; r++) {
      const fcStart = pv.getUint32(r * 4, true);
      const fcEnd = pv.getUint32((r + 1) * 4, true);
      const rgb = wordDocument[page + (crun + 1) * 4 + r];
      if (!rgb) continue;
      const chpx = page + rgb * 2;
      const cb = wordDocument[chpx];
      let at = chpx + 1;
      const end = chpx + 1 + cb;
      let picture: number | undefined;
      let spec = false;
      let ole = false;
      let data = false;
      while (at + 2 <= end) {
        const sprm = pv.getUint16(at - page, true);
        at += 2;
        let size = sprmOperandSize(sprm);
        if (size === -1) {
          size = sprm === 0xc615 ? wordDocument[at] + 1 : sprm === 0xd608 ? pv.getUint16(at - page, true) + 2 : wordDocument[at] + 1;
        }
        if (sprm === 0x6a03) picture = pv.getInt32(at - page, true);
        if (sprm === 0x0855) spec = wordDocument[at] !== 0;
        if (sprm === 0x080a) ole = wordDocument[at] !== 0;
        if (sprm === 0x0806) data = wordDocument[at] !== 0; // sprmCFData: the offset names field data, not a PICF
        at += size;
      }
      if (picture !== undefined && spec && !data && !offsets.has(picture)) {
        offsets.add(picture);
        out.push({ sprm: 0x6a03, offset: picture, fcStart, fcEnd, ole });
      }
    }
  }
  return out.sort((a, b) => a.fcStart - b.fcStart);
}
