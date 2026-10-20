// What a Word 6.0 or Word 95 .doc (nFib 0x65-0x68) keeps differently from a
// Word 97 one, as small translations onto the Word 97 structures the rest
// of the reader already handles. The source is Microsoft's "Word for
// Windows 6.0 Binary File Format" (1994); where it is wrong (some operand
// sizes), the files themselves decided. Apache POI's HWPFOldDocument
// (Apache License 2.0) gave the layout of the font table.
//
// - The FIB (fib.ts parseFib6) has its own fixed layout, and every
//   structure is in the WordDocument stream.
// - Text is 8-bit, in the character set of its font (or the Macintosh
//   character set for a file Word for the Macintosh saved).
// - Sprms are a one-byte opcode with their own numbering, and some
//   operands are smaller (a BRC is 2 bytes, a TC 10).
// - String tables and the font table hold 8-bit strings after a 2-byte
//   byte count.
// - Lists are autonumbering (ANLD) on each paragraph; numbering.ts reads
//   them through the two sprms below.
import { u16 } from "./bytes.js";
import type { Font } from "./styles.js";

/** The Word 97 shapes the Word 6 autonumbering sprms are given, which no Word 97 sprm carries: sprmPAnld (an ANLD) and sprmPNLvlAnm (a level). */
export const SPRM_ANLD = 0xc63e;
export const SPRM_NLVLANM = 0x243e;

/** A variable operand: a count byte, then that many bytes. */
const V = -1;
/** sprmTDefTable(10): a 2-byte count of the operand's bytes plus one. */
const LONG = -2;
/** sprmPChgTabs: a count byte that is 255 when the operand is longer. */
const TABS = -3;

/** The operand size of each Word 6 sprm, by opcode. */
const OPERAND: Record<number, number> = {
  2: 2, 3: V, 4: 1, 5: 1, 6: 1, 7: 1, 8: 1, 9: 1, 10: 1, 11: 1, 12: V, 13: 1, 14: 1, 15: V, 16: 2, 17: 2, 18: 2, 19: 2, 20: 4, 21: 2, 22: 2, 23: TABS,
  24: 1, 25: 1, 26: 2, 27: 2, 28: 2, 29: 1, 30: 2, 31: 2, 32: 2, 33: 2, 34: 2, 35: 2, 36: 2, 37: 1, 38: 2, 39: 2, 40: 2, 41: 2, 42: 2, 43: 2,
  44: 1, 45: 2, 46: 2, 47: 2, 48: 2, 49: 2, 50: 1, 51: 1, 52: V, 53: 1, 54: 1, 55: 1, 56: 1, 57: 1, 58: 1,
  65: 1, 66: 1, 67: 1, 68: V, 69: 2, 70: 4, 71: 1, 72: 2, 73: 3, 74: V, 75: 1, 80: 2, 81: V, 82: V, 83: 0, 85: 1, 86: 1, 87: 1, 88: 1, 89: 1,
  90: 1, 91: 1, 92: 1, 93: 2, 94: 1, 95: 3, 96: 2, 97: 2, 98: 1, 99: 2, 100: 1, 101: 2, 102: 1, 103: V, 104: 1, 105: V, 106: V, 107: 2, 108: V,
  109: 2, 110: 2, 117: 1, 118: 1, 119: 1, 120: V, 121: 2, 122: 2, 123: 2, 124: 2,
  131: 1, 132: 1, 133: V, 136: 3, 137: 3, 138: 1, 139: 1, 140: 2, 141: 2, 142: 1, 143: 1, 144: 2, 145: 2, 146: 1, 147: 1, 148: 2, 149: 2, 150: 1,
  151: 1, 152: 1, 153: 1, 154: 2, 155: 2, 156: 2, 157: 2, 158: 1, 159: 1, 160: 2, 161: 2, 162: 1, 163: 0, 164: 2, 165: 2, 166: 2, 167: 2, 168: 2,
  169: 2, 170: 2, 171: 2,
  182: 2, 183: 2, 184: 2, 185: 1, 186: 1, 187: 12, 188: LONG, 189: 2, 190: LONG, 191: V, 192: 4, 193: 5, 194: 4, 195: 2, 196: 4, 197: 2, 198: 2,
  199: 5, 200: 4,
};

/** The Word 6 sprms whose Word 97 sprm takes the same operand, by Word 6 opcode. */
const SAME: Record<number, number> = {
  // Paragraph: the Word 97 ispmd is the Word 6 opcode minus 2.
  2: 0x4600, 4: 0x2602, 5: 0x2403, 6: 0x2404, 7: 0x2405, 8: 0x2406, 9: 0x2407, 10: 0x2408, 11: 0x2409, 12: SPRM_ANLD, 13: SPRM_NLVLANM,
  14: 0x240c, 15: 0xc60d, 16: 0x840e, 17: 0x840f, 18: 0x4610, 19: 0x8411, 20: 0x6412, 21: 0xa413, 22: 0xa414, 23: 0xc615, 24: 0x2416,
  25: 0x2417, 26: 0x8418, 27: 0x8419, 28: 0x841a, 29: 0x261b, 37: 0x2423, 44: 0x242a, 45: 0x442b, 46: 0x442c, 47: 0x442d, 48: 0x842e,
  49: 0x842f, 50: 0x2430, 51: 0x2431, 53: 0x2433, 54: 0x2434, 55: 0x2435, 56: 0x2436, 57: 0x2437, 58: 0x2438,
  // Character.
  65: 0x0800, 66: 0x0801, 67: 0x0802, 71: 0x0806, 75: 0x080a, 80: 0x4a30, 85: 0x0835, 86: 0x0836, 87: 0x0837, 88: 0x0838, 89: 0x0839,
  90: 0x083a, 91: 0x083b, 92: 0x083c, 94: 0x2a3e, 96: 0x8840, 97: 0x486d, 98: 0x2a42, 99: 0x4a43, 101: 0x4845, 104: 0x2a48, 107: 0x484b,
  117: 0x0855, 118: 0x0856,
  // Section: the Word 97 ispmd is the Word 6 opcode minus 131, then minus 133.
  131: 0x3000, 132: 0x3001, 133: 0xd202, 136: 0xf203, 137: 0xf204, 138: 0x3005, 139: 0x3006, 140: 0x5007, 141: 0x5008, 142: 0x3009,
  143: 0x300a, 144: 0x500b, 145: 0x900c, 146: 0x300d, 147: 0x300e, 148: 0xb00f, 149: 0xb010, 150: 0x3011, 151: 0x3012, 152: 0x3013,
  153: 0x3014, 154: 0x5015, 155: 0x9016, 156: 0xb017, 157: 0xb018, 158: 0x3019, 159: 0x301a, 160: 0x501b, 161: 0x501c, 162: 0x301d,
  164: 0xb01f, 165: 0xb020, 166: 0xb021, 167: 0xb022, 168: 0x9023, 169: 0x9024, 170: 0xb025, 171: 0x5026,
  // Table.
  182: 0x5400, 183: 0x9601, 184: 0x9602, 185: 0x3403, 186: 0x3404, 189: 0x9407, 191: 0xd609, 192: 0x740a, 194: 0x7621, 195: 0x5622,
  196: 0x7623, 197: 0x5624, 198: 0x5625,
};

/** The paragraph border sprms (a 2-byte BRC), as Word 97's (a 4-byte Brc80). */
const BORDERS: Record<number, number> = { 38: 0x6424, 39: 0x6425, 40: 0x6426, 41: 0x6427, 42: 0x6428, 43: 0x6629 };

const translated = new WeakMap<Uint8Array, Uint8Array>();

/**
 * A Word 6 grpprl as the Word 97 grpprl that sets the same properties.
 * Sprms with no Word 97 counterpart the reader uses (revision reasons,
 * pasted-text character sets, size steps) are left out. A zero opcode
 * (padding) ends the grpprl.
 */
export function word6Grpprl(grpprl: Uint8Array): Uint8Array {
  const cached = translated.get(grpprl);
  if (cached) return cached;
  const out: number[] = [];
  const prl = (sprm: number, ...operand: number[]): void => {
    out.push(sprm & 0xff, sprm >>> 8, ...operand);
  };
  let i = 0;
  while (i < grpprl.length) {
    const op = grpprl[i];
    const size = OPERAND[op];
    if (size === undefined) break;
    let n: number;
    if (size === V) n = 1 + grpprl[i + 1];
    else if (size === LONG) n = 1 + u16(grpprl, i + 1);
    else if (size === TABS) {
      n = 1 + grpprl[i + 1];
      if (grpprl[i + 1] === 255) {
        const del = grpprl[i + 2];
        n = 2 + del * 4 + 1 + grpprl[i + 3 + del * 4] * 3;
      }
    } else n = size;
    if (i + 1 + n > grpprl.length) break;
    const o = grpprl.subarray(i + 1, i + 1 + n);
    i += 1 + n;
    const same = SAME[op];
    if (same !== undefined) {
      prl(same, ...o);
      continue;
    }
    const border = BORDERS[op];
    if (border !== undefined) {
      prl(border, ...brc80(u16(o, 0)));
      continue;
    }
    switch (op) {
      case 68: // sprmCPicLocation: a count byte, then the picture's offset. It makes the character special.
        prl(0x6a03, o[1], o[2], o[3], o[4]);
        prl(0x0855, 1);
        break;
      case 69: // sprmCIbstRMark: the one revision author, for an insertion and a deletion alike.
        prl(0x4804, o[0], o[1]);
        prl(0x4863, o[0], o[1]);
        break;
      case 70: // sprmCDttmRMark
        prl(0x6805, o[0], o[1], o[2], o[3]);
        prl(0x6864, o[0], o[1], o[2], o[3]);
        break;
      case 74: // sprmCSymbol: a count byte, the font, and an 8-bit character in it.
        prl(0x6a09, o[1], o[2], o[3], 0xf0);
        prl(0x0855, 1);
        break;
      case 82: // sprmCDefault: the toggles, the underline and the color off.
        for (const toggle of [0x0835, 0x0836, 0x0837, 0x0838, 0x0839, 0x083a, 0x083b, 0x083c]) prl(toggle, 0);
        prl(0x2a3e, 0);
        prl(0x2a42, 0);
        break;
      case 83: // sprmCPlain
        prl(0x2a33, 0);
        break;
      case 93: // sprmCFtc: one font for all text.
        prl(0x4a4f, o[0], o[1]);
        prl(0x4a51, o[0], o[1]);
        break;
      case 95: // sprmCSizePos: a size when not zero, and a raised position when not 128.
        if (o[0] !== 0) prl(0x4a43, o[0], 0);
        if (o[2] !== 128) prl(0x4845, ...int8To16(o[2]));
        break;
      case 102: // sprmCHpsPosAdj: the raised position (the size step it implies is left out).
        prl(0x4845, ...int8To16(o[0]));
        break;
      case 105: // sprmCHpsNew50: a count byte, then the size.
        prl(0x4a43, o[1], o[2]);
        break;
      case 187: // sprmTTableBorders: six BRCs.
        prl(0xd605, 24, ...[0, 2, 4, 6, 8, 10].flatMap((at) => brc80(u16(o, at))));
        break;
      case 188: // sprmTDefTable10: its TCs' borders are BRC10s, left out.
      case 190: // sprmTDefTable
        prl(0xd608, ...defTable(o, op === 190));
        break;
      case 193: // sprmTSetBrc
        prl(0xd620, 7, o[0], o[1], o[2], ...brc80(u16(o, 3)));
        break;
      default:
        break;
    }
  }
  const result = new Uint8Array(out);
  translated.set(grpprl, result);
  return result;
}

function int8To16(b: number): number[] {
  return b & 0x80 ? [b, 0xff] : [b, 0];
}

/**
 * A Word 6 BRC (2 bytes: a line width in 0.75 points, or 6 dotted and 7
 * dashed; a type: none, single, thick, double; a shadow; an Ico; the space
 * in points) as a Brc80 (4 bytes: a width in eighths of a point, a
 * BrcType, an Ico, the space and the shadow).
 */
export function brc80(v: number): number[] {
  const lineWidth = v & 7;
  const type = (v >>> 3) & 3;
  if (type === 0 || lineWidth === 0) return [0, 0, 0, 0];
  const shadow = (v >>> 5) & 1;
  const ico = (v >>> 6) & 0x1f;
  const space = (v >>> 11) & 0x1f;
  const brcType = lineWidth === 6 ? 6 : lineWidth === 7 ? 7 : type;
  const width = lineWidth >= 6 ? 6 : lineWidth * 6;
  return [width, brcType, ico, space | (shadow << 5)];
}

/**
 * sprmTDefTable's operand (a 2-byte count, the cell count, the cell edges,
 * then a 10-byte TC per cell, trailing ones left out when empty) as Word
 * 97's (the same, with 20-byte TC80s).
 */
function defTable(o: Uint8Array, borders: boolean): number[] {
  const count = o[2];
  const edgesEnd = 3 + (count + 1) * 2;
  const body: number[] = [count, ...o.subarray(3, edgesEnd)];
  for (let c = 0; c < count; c++) {
    const at = edgesEnd + c * 10;
    const tc = at + 10 <= o.length ? o.subarray(at, at + 10) : new Uint8Array(10);
    // rgf: fFirstMerged and fMerged are the TC80's first two bits too; the width is left to the edges.
    body.push(tc[0] & 3, 0, 0, 0);
    for (let side = 0; side < 4; side++) body.push(...(borders ? brc80(u16(tc, 2 + side * 2)) : [0, 0, 0, 0]));
  }
  const cb = body.length + 1;
  return [cb & 0xff, cb >>> 8, ...body];
}

/** A Word 6 string table: a 2-byte byte count, then each string's length byte, its 8-bit characters and `cbExtra` bytes. */
export function parseSttb6(b: Uint8Array, cbExtra: number, decode: (bytes: Uint8Array) => string = latin1): { strings: string[]; extra: Uint8Array[] } {
  const strings: string[] = [];
  const extra: Uint8Array[] = [];
  const end = b.length >= 2 ? Math.min(b.length, u16(b, 0)) : 0;
  let at = 2;
  while (at < end) {
    const cch = b[at];
    strings.push(decode(b.subarray(at + 1, at + 1 + cch)));
    at += 1 + cch;
    extra.push(b.subarray(at, at + cbExtra));
    at += cbExtra;
  }
  return { strings, extra };
}

/** The Word 6 font table (SttbfFfn): a 2-byte byte count, then each FFN with its 8-bit name. */
export function parseFonts6(blob: Uint8Array): Font[] {
  const fonts: Font[] = [];
  const end = blob.length >= 2 ? Math.min(blob.length, u16(blob, 0)) : 0;
  let at = 2;
  while (at < end) {
    const cbFfnM1 = blob[at];
    const ffn = blob.subarray(at + 1, at + 1 + cbFfnM1);
    at += 1 + cbFfnM1;
    // prg and ff, wWeight, chs, ibszAlt, then the name.
    let name = "";
    for (let k = 5; k < ffn.length && ffn[k] !== 0; k++) name += String.fromCharCode(ffn[k]);
    fonts.push({ name, charset: ffn[3] ?? 0 });
  }
  return fonts;
}

function latin1(bytes: Uint8Array): string {
  let s = "";
  for (let i = 0; i < bytes.length; i++) s += String.fromCharCode(bytes[i]);
  return s;
}

/** The code page of each Windows character set an FFN names. */
const CHARSET_ENCODINGS: Record<number, string> = {
  77: "macintosh", 161: "windows-1253", 162: "windows-1254", 163: "windows-1258", 177: "windows-1255", 178: "windows-1256", 186: "windows-1257",
  204: "windows-1251", 222: "windows-874", 238: "windows-1250",
};

const highHalves = new Map<string, string>();

/**
 * The characters bytes 0x80-0xFF stand for in a font with the given
 * character set: the document's own character set (Windows or Macintosh)
 * for an ANSI or default font, the font's for the others.
 */
export function highHalf(charset: number, mac: boolean): string {
  const encoding = CHARSET_ENCODINGS[charset] ?? (mac ? "macintosh" : "windows-1252");
  let half = highHalves.get(encoding);
  if (half === undefined) {
    const bytes = new Uint8Array(128).map((_, i) => 0x80 + i);
    half = new TextDecoder(encoding).decode(bytes);
    highHalves.set(encoding, half);
  }
  return half;
}
