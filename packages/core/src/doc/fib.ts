// The File Information Block at the start of the WordDocument stream
// ([MS-DOC] 2.5): the format version, the flags, the length of each
// document part in characters, and the offset and size in the Table stream
// of every other structure (the fc/lcb pairs of FibRgFcLcb97 and later).
import { u16, u32 } from "./bytes.js";

/** Indices into the fc/lcb pairs ([MS-DOC] 2.5.6-2.5.10), by name. */
export const FC = {
  stshf: 1,
  plcffndRef: 2,
  plcffndTxt: 3,
  plcfandRef: 4,
  plcfandTxt: 5,
  plcfSed: 6,
  plcfHdd: 11,
  plcfBteChpx: 12,
  plcfBtePapx: 13,
  sttbfFfn: 15,
  plcfFldMom: 16,
  sttbfBkmk: 21,
  plcfBkf: 22,
  plcfBkl: 23,
  dop: 31,
  clx: 33,
  grpXstAtnOwners: 36,
  sttbfAtnBkmk: 37,
  plcSpaMom: 40,
  plcSpaHdr: 41,
  plcfAtnBkf: 42,
  plcfAtnBkl: 43,
  plcfendRef: 46,
  plcfendTxt: 47,
  dggInfo: 50,
  sttbfRMark: 51,
  plcftxbxTxt: 56,
  plcfHdrtxbxTxt: 58,
  plfLst: 73,
  plfLfo: 74,
  plcfTxbxBkd: 75,
  plcfTxbxHdrBkd: 76,
  atrdExtra: 112,
  sttbfBkmkMoveFrom: 165,
  plcfBkfMoveFrom: 166,
  plcfBklMoveFrom: 167,
  sttbfBkmkMoveTo: 168,
  plcfBkfMoveTo: 169,
  plcfBklMoveTo: 170,
} as const;

export interface Fib {
  /** The format version: 0x00C1 Word 97, 0x00D9 Word 2000, 0x0101 Word 2002, 0x010C Word 2003, 0x0112 Word 2007. */
  nFib: number;
  encrypted: boolean;
  obfuscated: boolean;
  /** The document parts' lengths in characters, in their order in the text ([MS-DOC] 2.3). */
  ccpText: number;
  ccpFtn: number;
  ccpHdd: number;
  ccpAtn: number;
  ccpEdn: number;
  ccpTxbx: number;
  ccpHdrTxbx: number;
  /** The count of fc/lcb pairs the file holds. */
  cbRgFcLcb: number;
  /** The offset and size in the Table stream of the structure at fc/lcb pair `index`; zeros past the end. */
  fcLcb(index: number): { fc: number; lcb: number };
}

/** Where each document part starts, in the document's one CP space. */
export interface PartStarts {
  main: number;
  footnotes: number;
  headers: number;
  comments: number;
  endnotes: number;
  textboxes: number;
  headerTextboxes: number;
  /** One past the last part. */
  end: number;
}

export function partStarts(fib: Fib): PartStarts {
  const main = 0;
  const footnotes = main + fib.ccpText;
  const headers = footnotes + fib.ccpFtn;
  const comments = headers + fib.ccpHdd;
  const endnotes = comments + fib.ccpAtn;
  const textboxes = endnotes + fib.ccpEdn;
  const headerTextboxes = textboxes + fib.ccpTxbx;
  return { main, footnotes, headers, comments, endnotes, textboxes, headerTextboxes, end: headerTextboxes + fib.ccpHdrTxbx };
}

/**
 * Parse the FIB. Throws "unsupported format: ..." for a Word 6.0/95 or
 * older file, and "unsupported format: a password-protected document ..."
 * for an encrypted or obfuscated one.
 */
export function parseFib(word: Uint8Array): Fib {
  const wIdent = u16(word, 0);
  if (wIdent === 0xa5dc) throw new Error("unsupported format: a Word 6.0/95 document (only Word 97-2003 .doc files are read)");
  if (wIdent !== 0xa5ec) throw new Error("not a Word document: the WordDocument stream does not start with a FIB");
  let nFib = u16(word, 2);
  const flags = u16(word, 0x0a);
  const encrypted = (flags & 0x0100) !== 0;
  const obfuscated = (flags & 0x8000) !== 0;
  const csw = u16(word, 0x20);
  const atLw = 0x22 + csw * 2;
  const cslw = u16(word, atLw);
  const lw = atLw + 2;
  const atFcLcb = lw + cslw * 4;
  const cbRgFcLcb = u16(word, atFcLcb);
  const fcLcbBase = atFcLcb + 2;
  const atCswNew = fcLcbBase + cbRgFcLcb * 8;
  if (atCswNew + 2 <= word.length) {
    const cswNew = u16(word, atCswNew);
    if (cswNew > 0 && atCswNew + 4 <= word.length) nFib = u16(word, atCswNew + 2);
  }
  if (nFib < 0x00c1) throw new Error("unsupported format: a Word 6.0/95 document (only Word 97-2003 .doc files are read)");
  if (encrypted) throw new Error(`unsupported format: a password-protected document (${obfuscated ? "XOR obfuscated" : "RC4 encrypted"})`);
  const ccp = (i: number): number => (lw + 12 + i * 4 + 4 <= word.length ? u32(word, lw + 12 + i * 4) : 0);
  return {
    nFib,
    encrypted,
    obfuscated,
    ccpText: ccp(0),
    ccpFtn: ccp(1),
    ccpHdd: ccp(2),
    ccpAtn: ccp(4),
    ccpEdn: ccp(5),
    ccpTxbx: ccp(6),
    ccpHdrTxbx: ccp(7),
    cbRgFcLcb,
    fcLcb(index) {
      if (index >= cbRgFcLcb) return { fc: 0, lcb: 0 };
      const at = fcLcbBase + index * 8;
      if (at + 8 > word.length) return { fc: 0, lcb: 0 };
      return { fc: u32(word, at), lcb: u32(word, at + 4) };
    },
  };
}

/** The bytes of the Table stream structure at fc/lcb pair `index`; empty when absent or out of range. */
export function tableBlob(fib: Fib, table: Uint8Array, index: number): Uint8Array {
  const { fc, lcb } = fib.fcLcb(index);
  if (lcb === 0 || fc + lcb > table.length) return new Uint8Array(0);
  return table.subarray(fc, fc + lcb);
}
