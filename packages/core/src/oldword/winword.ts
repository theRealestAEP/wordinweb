// Word for Windows 1.x (signature 9B A5) and 2.0 (DB A5): word2.ts's format,
// little-endian, with the text in the Windows ANSI code page (code page
// 1252), the fonts in a table of names, a character's properties in Word
// for Windows 2.0's CHP layout, fields, and pictures as Windows metafiles
// or bitmaps.
//
// Written from Microsoft's description of the Word for Windows 2.0 binary
// format as LibreOffice's Word reader (sw/source/filter/ww8, ww::eWW2)
// applies it, which was a guide to the structures only.
import { dibToBmp } from "../doc/officeart.js";
import { newSectionProps } from "../doc/sprm.js";
import { symbolText } from "../pdf/dib.js";
import { Inline, OldDocument } from "./model.js";
import { Char, DocProps, Entry, Fib, Sprm, Word2Reader, popcount } from "./word2.js";

const CP1252 = new TextDecoder("windows-1252");
const CP1252_CHARS: string[] = Array.from({ length: 256 }, (_, i) => CP1252.decode(Uint8Array.of(i)));

/** The operand sizes of Word for Windows 2.0's sprms. */
const SPRM_SIZES: Record<number, number> = {
  2: 1, 3: -1, 4: 1, 5: 1, 6: 1, 7: 1, 8: 1, 9: 1, 10: 1, 11: 1, 12: 1, 13: 1, 14: 1, 15: -1, 16: 2, 17: 2, 18: 2, 19: 2, 20: 2, 21: 2, 22: 2,
  23: -1, 24: 1, 25: 1, 26: 2, 27: 2, 28: 2, 29: 1, 30: 2, 31: 2, 32: 2, 33: 2, 34: 2, 35: 2, 36: 2, 37: 1, 38: 2, 39: 2, 40: 2, 41: 2,
  42: 2, 43: 2, 44: 1, 45: 2, 46: 2, 47: 2, 48: 2, 49: 2, 50: 1, 51: 1, 52: 0, 53: 1, 54: 1, 55: 1, 57: -1, 58: 0, 60: 1, 61: 1, 62: 1,
  63: 1, 64: 1, 65: 1, 66: 1, 67: 1, 68: 2, 69: 1, 70: 3, 71: 2, 72: 2, 73: 1, 74: 1, 75: 1, 76: 1, 77: 1, 78: -1, 80: 1, 81: 1, 82: 2,
  83: 2, 84: 1, 85: 1, 86: 1, 87: 1, 94: 1, 95: -1, 96: 2, 97: 2, 98: 2, 99: 2, 112: 1, 114: 1, 115: 2, 116: 2, 117: 1, 118: 1, 119: 2,
  120: 2, 121: 1, 122: 1, 123: 2, 124: 2, 125: 1, 126: 1, 127: 1, 128: 1, 129: 2, 130: 2, 131: 2, 132: 2, 133: 1, 134: 1, 135: 2, 136: 2,
  137: 1, 138: 1, 139: 2, 140: 2, 141: 2, 142: 2, 143: 2, 144: 2, 145: 2, 146: 2, 147: 2, 148: 2, 149: 1, 152: -2, 153: 2, 154: -2,
  155: -1, 157: 5, 158: 4, 159: 2, 160: 4, 161: 2, 162: 2, 163: 5, 164: 4,
};

/** Whether `head` starts a Word for Windows 1.x or 2.0 document. */
export function isWinWordFile(head: Uint8Array): boolean {
  if (head.length < 4 || head[1] !== 0xa5) return false;
  const nFib = head[2] | (head[3] << 8);
  return (head[0] === 0x9b && nFib === 0x21) || (head[0] === 0xdb && nFib === 0x2d);
}

/** A Word for Windows 1.x or 2.0 document as an OldDocument. */
export function readWinWord(bytes: Uint8Array): OldDocument {
  return new WinWordReader(bytes).read();
}

class WinWordReader extends Word2Reader {
  protected readonly sprmSizes = SPRM_SIZES;
  private readonly fontNames: string[] = [];
  private dop: Uint8Array = new Uint8Array(0);

  constructor(bytes: Uint8Array) {
    super(bytes, false);
  }

  /**
   * The FIB: the text's start at 0x18, its lengths (main text, footnotes,
   * headers and footers) from 0x34, and from 0x58 file positions with
   * 16-bit sizes, in Word for Windows 2.0's order (Word 1 has two fewer
   * printer entries before the document's properties).
   */
  protected readFib(): Fib {
    const e = (i: number): Entry => ({ fc: this.u32(0x58 + 6 * i), cb: this.u16(0x5c + 6 * i) });
    const word2 = this.bytes[0] === 0xdb;
    this.dop = this.entry(e(word2 ? 31 : 29));
    this.readFonts(e(15));
    return {
      fcMin: this.u32(0x18), ccpText: this.u32(0x34), ccpFtn: this.u32(0x38), ccpHdd: this.u32(0x3c),
      stsh: e(1).cb ? e(1) : e(0), fndRef: e(2), fndTxt: e(3), sed: e(6), hdd: e(11), bteChpx: e(12), btePapx: e(13), clx: e(word2 ? 33 : 31),
    };
  }

  /** The font table: its size, then per font a size byte, two bytes, and the name ending in a zero byte. */
  private readFonts(entry: Entry): void {
    const table = this.entry(entry);
    let at = 2;
    while (at + 3 < table.length) {
      const size = table[at] + 1;
      const name = table.subarray(at + 3, Math.min(at + size, table.length));
      const end = name.indexOf(0);
      this.fontNames.push(CP1252.decode(end >= 0 ? name.subarray(0, end) : name));
      at += size;
    }
  }

  /** The document's properties: flags (facing pages 0x01, where the notes go in bits 5-6), the note separators there are, the default tab at 10. */
  protected readDocProps(): DocProps {
    const dop = this.dop;
    return {
      page: newSectionProps(),
      defaultTab: dop.length >= 12 ? this.n16(dop, 10) || undefined : undefined,
      facingPages: dop.length > 0 && (dop[0] & 0x01) !== 0,
      endnotes: dop.length > 0 && ((dop[0] >> 5) & 3) === 0,
      separatorStories: dop.length > 1 ? popcount(dop[1] & 0x3f) : 0,
    };
  }

  protected baseChar(): Char {
    return { fontId: 0, hps: 20, flags: new Array(8).fill(false), kul: 0, hpsPos: 0, dxaSpace: 0, ico: 0, special: false };
  }

  /**
   * A CHPX: the first bytes of a character's properties (CHP), as far as
   * they differ from the paragraph style's. Byte 0 flips bold, italic,
   * outline, small caps, all caps and hidden, byte 1 strikethrough and holds
   * fSpec; byte 2 says which of the others are set: the font, the size, the
   * spacing, the color and underline, the raise; then the picture's file
   * offset at 20.
   */
  protected applyChpx(c: Char, d: Uint8Array): void {
    if (d.length === 0) return;
    const flips: [number, number][] = [[0x01, 0], [0x02, 1], [0x08, 3], [0x20, 5], [0x40, 6], [0x80, 7]];
    for (const [bit, i] of flips) if (d[0] & bit) c.flags[i] = !c.flags[i];
    if (d.length < 2) return;
    if (d[1] & 0x04) c.flags[2] = !c.flags[2];
    c.special = (d[1] & 0x02) !== 0;
    const set = d.length >= 3 ? d[2] : 0;
    if (set & 0x02 && d.length >= 6) c.fontId = this.n16(d, 4);
    if (set & 0x04 && d.length >= 8) c.hps = this.n16(d, 6);
    if (set & 0x20) c.dxaSpace = d.length >= 9 ? (((d[8] & 0x3f) << 26) >> 26) * 5 : 0;
    if (set & 0x01) c.ico = d.length >= 10 ? d[9] & 0x1f : 0;
    if (set & 0x08) c.kul = d.length >= 10 ? d[9] >> 5 : 0;
    if (set & 0x10) c.hpsPos = d.length >= 11 ? (d[10] << 24) >> 24 : 0;
    if (d.length >= 24) c.picture = this.n32(d, 20);
  }

  protected isCharSprm(op: number): boolean {
    return op >= 53 && op <= 87;
  }

  /** A character sprm of a paragraph's or a piece's. Toggles: 0 off, 1 on, 0x80 the style's value, 0x81 its opposite. */
  protected applyCharSprm(c: Char, { op, arg }: Sprm, style: Char): void {
    if (op >= 60 && op <= 67) {
      const i = op - 60;
      const v = arg[0];
      c.flags[i] = v === 0x80 ? style.flags[i] : v === 0x81 ? !style.flags[i] : v !== 0;
      return;
    }
    switch (op) {
      case 58: Object.assign(c, { ...style, flags: [...style.flags], picture: c.picture, special: c.special }); break;
      case 68: c.fontId = this.n16(arg, 0); break;
      case 69: c.kul = arg[0]; break;
      case 71: c.dxaSpace = this.n16s(arg, 0); break;
      case 73: c.ico = arg[0]; break;
      case 74: if (arg[0]) c.hps = arg[0]; break;
      case 76: c.hpsPos = (arg[0] << 24) >> 24; break;
      default: break;
    }
  }

  protected fontName(id: number): string {
    return this.fontNames[id] ?? "Times New Roman";
  }

  /**
   * How a font's character codes become text: the Symbol font's codes are
   * its Unicode characters; a pictograph font's (Wingdings and the like)
   * stay codes, in the private-use block at U+F000 as Word keeps them; any
   * other font's are Windows ANSI.
   */
  protected decoder(fontId: number): (code: number) => string {
    const name = this.fontName(fontId);
    if (/^symbol$/i.test(name)) return (code) => symbolText(Uint8Array.of(code));
    if (/wingdings|webdings|dingbats|sorts/i.test(name)) return (code) => String.fromCharCode(0xf000 + code);
    return (code) => CP1252_CHARS[code];
  }

  /** A tab descriptor: the alignment in the low three bits, the leader in the next three. */
  protected tab(type: number): { jc: number; leader: number } {
    return { jc: Math.min(type & 7, 4), leader: (type >> 3) & 7 };
  }

  /** A piece's property modifier: with its low bit set, a grpprl's number in the rest; else a sprm in bits 1-7 and its byte. */
  protected prmSprms(prm: number, grpprls: Uint8Array[]): Sprm[] {
    if (prm & 1) return this.parseSprms(grpprls[prm >>> 1] ?? new Uint8Array(0));
    if ((prm >>> 1) & 0x7f) return this.parseSprms(Uint8Array.of((prm >>> 1) & 0x7f, prm >>> 8));
    return [];
  }

  protected styleName(bytes: Uint8Array): string {
    return CP1252.decode(bytes);
  }

  /**
   * The picture a U+0001 shows: a PIC at the file offset its run gives (its
   * size, its header's size, the metafile's mapping mode and extent, the
   * size it shows at in twips, its scale in tenths of a percent and its
   * crops), then a Windows metafile or a bitmap.
   */
  private readonly pictureIndex = new Map<number, number>();
  protected picture(char: Char): Inline | undefined {
    const at = char.picture ?? 0;
    const lcb = this.u32(at);
    const cbHeader = this.u16(at + 4);
    if (at < 0x80 || lcb < 0x3a || cbHeader < 0x3a || cbHeader >= lcb || at + lcb > this.bytes.length) return undefined;
    let index = this.pictureIndex.get(at);
    if (index === undefined) {
      const data = this.bytes.subarray(at + cbHeader, at + lcb);
      const image = pictureData(data);
      if (!image) return undefined;
      index = this.pictures.length;
      this.pictures.push(image);
      this.pictureIndex.set(at, index);
    }
    const s16 = (o: number): number => this.n16s(this.bytes, at + o);
    const mx = this.u16(at + 0x20) || 1000;
    const my = this.u16(at + 0x22) || 1000;
    const width = ((s16(0x1c) - s16(0x24) - s16(0x28)) * mx) / 1000;
    const height = ((s16(0x1e) - s16(0x26) - s16(0x2a)) * my) / 1000;
    if (width <= 0 || height <= 0) return undefined;
    return { kind: "picture", index, width, height };
  }
}

/** A picture's data: a Windows metafile (its records, with no placeable header) or a device-independent bitmap. */
function pictureData(data: Uint8Array): { data: Uint8Array; ext: "wmf" | "bmp" } | undefined {
  if (data.length < 18) return undefined;
  const v = new DataView(data.buffer, data.byteOffset, data.byteLength);
  if (data[0] === 0x42 && data[1] === 0x4d) return { data: data.slice(), ext: "bmp" };
  if (v.getUint32(0, true) === 0x9ac6cdd7) return { data: data.slice(), ext: "wmf" };
  const type = v.getUint16(0, true);
  if ((type === 1 || type === 2) && v.getUint16(2, true) === 9) return { data: data.slice(), ext: "wmf" };
  const header = v.getUint32(0, true);
  if (header === 40 || header === 12) {
    const bmp = dibToBmp(data);
    return bmp ? { data: bmp.data, ext: "bmp" } : undefined;
  }
  return undefined;
}
