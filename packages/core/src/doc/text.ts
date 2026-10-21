// The piece table ([MS-DOC] 2.4.1): where each run of the document's text
// is in the WordDocument stream, and whether it is 8-bit (Windows-1252) or
// 16-bit. Every document part (main text, footnotes, headers, comments,
// endnotes, text boxes) is in the one CP space the pieces cover.
import { u16, u32 } from "./bytes.js";
import type { Word6Fib } from "./fib.js";
import { word6Grpprl } from "./word6.js";
import { windows1252 as cp1252 } from "../windows1252.js";

export interface Piece {
  /** The first CP of the piece and one past its last. */
  cp: number;
  end: number;
  /** The byte offset in the WordDocument stream of the piece's first character. */
  fc: number;
  /** 8-bit characters at `fc`; otherwise 16-bit ones. */
  compressed: boolean;
  /** The Pcd's Prm: extra properties of the piece's text (prmGrpprl resolves it). */
  prm: number;
}

/** The Prm0 sprms by isprm ([MS-DOC] 2.9.215). */
const PRM0_SPRMS: Record<number, number> = {
  0x00: 0x2879, 0x04: 0x2602, 0x05: 0x2461, 0x07: 0x2405, 0x08: 0x2406, 0x09: 0x2407, 0x0c: 0x260a, 0x0d: 0x2470, 0x0e: 0x240c,
  0x0f: 0x2471, 0x18: 0x2416, 0x19: 0x2417, 0x1d: 0x261b, 0x25: 0x2423, 0x2c: 0x242a, 0x32: 0x2430, 0x33: 0x2431, 0x35: 0x2433,
  0x36: 0x2434, 0x37: 0x2435, 0x38: 0x2436, 0x39: 0x2437, 0x3a: 0x2438, 0x41: 0x0800, 0x42: 0x0801, 0x43: 0x0802, 0x47: 0x0806,
  0x4b: 0x080a, 0x4d: 0x2a0c, 0x4e: 0x0858, 0x4f: 0x2859, 0x50: 0x0811, 0x51: 0x0818, 0x53: 0x2a33, 0x55: 0x0835, 0x56: 0x0836,
  0x57: 0x0837, 0x58: 0x0838, 0x59: 0x0839, 0x5a: 0x083a, 0x5b: 0x083b, 0x5c: 0x083c, 0x5e: 0x2a3e, 0x62: 0x2a42, 0x68: 0x2a48,
  0x73: 0x2a53, 0x74: 0x0854, 0x75: 0x0855, 0x76: 0x0856, 0x78: 0x2640, 0x7b: 0x2a90, 0x7c: 0x2a86, 0x7e: 0x2443,
};

/** Each byte as the character of the same code (the Encoding Standard's "latin1" is Windows-1252). */
const bytesDecoder = {
  decode(bytes: Uint8Array): string {
    const parts: string[] = [];
    for (let at = 0; at < bytes.length; at += 8192) parts.push(String.fromCharCode(...bytes.subarray(at, at + 8192)));
    return parts.join("");
  },
};

export class PieceTable {
  readonly pieces: Piece[] = [];
  /** The Clx's Prc grpprls, which a Prm1 indexes. */
  private readonly prcs: Uint8Array[] = [];
  private readonly prmCache = new Map<number, Uint8Array | null>();
  /** 8-bit text: Windows-1252, or for a Word 6 file the bytes themselves (its characters depend on their font). */
  private readonly eightBit: { decode(bytes: Uint8Array): string };

  /**
   * A Word 6.0/95 file (`word6`) has 8-bit text only: at the offset its
   * Pcd gives, with no compression flag, or, when it was not fast-saved,
   * in one run from fcMin to fcMac with no piece table at all.
   */
  constructor(clx: Uint8Array, private readonly word: Uint8Array, private readonly word6?: Word6Fib) {
    this.eightBit = word6 ? bytesDecoder : cp1252;
    if (word6 && !word6.complex) {
      this.pieces.push({ cp: 0, end: Math.max(0, word6.fcMac - word6.fcMin), fc: word6.fcMin * 2, compressed: true, prm: 0 });
      return;
    }
    let at = 0;
    while (at < clx.length) {
      const clxt = clx[at];
      if (clxt === 0x01) {
        const cb = u16(clx, at + 1);
        const prc = clx.subarray(at + 3, at + 3 + cb);
        this.prcs.push(word6 ? word6Grpprl(prc) : prc);
        at += 3 + cb;
      } else if (clxt === 0x02) {
        const lcb = u32(clx, at + 1);
        const plc = clx.subarray(at + 5, at + 5 + lcb);
        const n = Math.floor((lcb - 4) / 12);
        for (let i = 0; i < n; i++) {
          const pcd = (n + 1) * 4 + i * 8;
          const fcc = u32(plc, pcd + 2);
          const cp = u32(plc, i * 4);
          const end = u32(plc, (i + 1) * 4);
          // A Word 6 Pcd's fc is the plain byte offset of 8-bit text (stored doubled, as Word 97's compressed fc is).
          if (word6) this.pieces.push({ cp, end, fc: fcc * 2, compressed: true, prm: u16(plc, pcd + 6) });
          else this.pieces.push({ cp, end, fc: fcc & 0x3fffffff, compressed: (fcc & 0x40000000) !== 0, prm: u16(plc, pcd + 6) });
        }
        break;
      } else {
        throw new Error("damaged: the piece table (Clx) holds an unknown entry");
      }
    }
    if (this.pieces.length === 0) throw new Error("damaged: the document has no piece table");
  }

  /** One past the last CP any piece covers. */
  get end(): number {
    return this.pieces[this.pieces.length - 1].end;
  }

  /** The index of the piece holding `cp`, or the last piece for a CP past the end. */
  pieceIndex(cp: number): number {
    let lo = 0;
    let hi = this.pieces.length - 1;
    while (lo < hi) {
      const mid = (lo + hi + 1) >> 1;
      if (this.pieces[mid].cp <= cp) lo = mid;
      else hi = mid - 1;
    }
    return lo;
  }

  /** The byte offset in the WordDocument stream of the character at `cp`. */
  fcOf(cp: number): number {
    const piece = this.pieces[this.pieceIndex(cp)];
    return piece.compressed ? piece.fc / 2 + (cp - piece.cp) : piece.fc + 2 * (cp - piece.cp);
  }

  /** The text at [start, end): control characters and all, as the file holds them. */
  text(start: number, end: number): string {
    if (end <= start) return "";
    const parts: string[] = [];
    for (let i = this.pieceIndex(start); i < this.pieces.length; i++) {
      const piece = this.pieces[i];
      if (piece.cp >= end) break;
      const lo = Math.max(piece.cp, start);
      const hi = Math.min(piece.end, end);
      if (lo >= hi) continue;
      if (piece.compressed) {
        const from = piece.fc / 2 + (lo - piece.cp);
        parts.push(this.eightBit.decode(this.word.subarray(from, from + (hi - lo))));
      } else {
        const from = piece.fc + 2 * (lo - piece.cp);
        const bytes = this.word.subarray(from, from + 2 * (hi - lo));
        // String.fromCharCode in bounded slices: one call per 8K characters.
        for (let at = 0; at < bytes.length; at += 16384) {
          const chunk = bytes.subarray(at, Math.min(at + 16384, bytes.length));
          const codes = new Array<number>(chunk.length >> 1);
          for (let k = 0; k < codes.length; k++) codes[k] = chunk[2 * k] | (chunk[2 * k + 1] << 8);
          parts.push(String.fromCharCode(...codes));
        }
      }
    }
    return parts.join("");
  }

  /** The grpprl a piece's Prm adds to its text's direct properties, or null for none. */
  prmGrpprl(prm: number): Uint8Array | null {
    if (prm === 0) return null;
    const cached = this.prmCache.get(prm);
    if (cached !== undefined) return cached;
    let out: Uint8Array | null = null;
    if (prm & 0x0001) {
      const index = prm >>> 1;
      const prc = this.prcs[index];
      if (prc) out = prc;
    } else if (this.word6) {
      // A Word 6 Prm holds one sprm whose operand is one byte or none.
      const grpprl = word6Grpprl(new Uint8Array([(prm >>> 1) & 0x7f, (prm >>> 8) & 0xff]));
      if (grpprl.length) out = grpprl;
    } else {
      const isprm = (prm >>> 1) & 0x7f;
      const val = (prm >>> 8) & 0xff;
      const sprm = PRM0_SPRMS[isprm];
      if (sprm !== undefined && !(isprm === 0 && val === 0)) {
        // Every Prm0 sprm takes a one-byte operand.
        out = new Uint8Array([sprm & 0xff, sprm >>> 8, val]);
      }
    }
    this.prmCache.set(prm, out);
    return out;
  }
}
