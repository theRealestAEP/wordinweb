// Word for the Macintosh 4.0 and 5.x (signature FE 37 00): word2.ts's
// format, big-endian, with the text in the Mac OS Roman code page, the
// fonts by their Mac font numbers, a character's properties in the Mac's
// own layout, and pictures as Macintosh PICTs.
//
// Written from the file format as observed in Word's files, with the
// libmwaw project's notes on it (MsWrdParser) as a guide to the structures.
import { newSectionProps } from "../doc/sprm.js";
import { symbolText } from "../pdf/dib.js";
import { Inline, OldDocument } from "./model.js";
import { Char, DocProps, Entry, Fib, STANDARD_STYLES, Sprm, Word2Reader, popcount } from "./word2.js";

const MAC_ROMAN = new TextDecoder("macintosh");
const MAC_ROMAN_CHARS: string[] = Array.from({ length: 256 }, (_, i) => MAC_ROMAN.decode(Uint8Array.of(i)));

/** The Mac fonts by their system number, for a Word 4 file (no font table of its own) and a font a Word 5 table leaves out. */
export const MAC_FONTS: Record<number, string> = {
  0: "Chicago", 1: "Geneva", 2: "New York", 3: "Geneva", 4: "Monaco", 5: "Venice", 6: "London", 7: "Athens", 8: "San Francisco", 9: "Toronto",
  11: "Cairo", 12: "Los Angeles", 13: "Zapf Dingbats", 14: "Bookman", 15: "Helvetica Narrow", 16: "Palatino", 18: "Zapf Chancery", 20: "Times",
  21: "Helvetica", 22: "Courier", 23: "Symbol", 24: "Mobile", 33: "Avant Garde", 34: "New Century Schoolbook",
};

/** The date characters, by their field picture: the short date, its day, month and year, and the abbreviated and long forms. */
const DATE_PICTURES: Record<number, string> = {
  0x03: "", 0x10: "d", 0x13: "M", 0x15: "yy", 0x16: "ddd", 0x17: "dddd", 0x1a: "MMM", 0x1b: "MMMM", 0x1d: "MMM d, yyyy",
};

/** The operand sizes of the sprms Word for the Macintosh writes (Word for Windows 2.0's numbering, with Word for Windows 2.0's sizes where the Mac files do not show their own). */
export const MAC_SPRM_SIZES: Record<number, number> = {
  0x02: 1, 0x03: -1, 0x04: 1, 0x05: 1, 0x06: 1, 0x07: 1, 0x08: 1, 0x09: 1, 0x0a: 1, 0x0b: 1, 0x0c: 1, 0x0d: 1, 0x0e: 1, 0x0f: -1,
  0x10: 2, 0x11: 2, 0x12: 2, 0x13: 2, 0x14: 2, 0x15: 2, 0x16: 2, 0x17: -1, 0x18: 1, 0x19: 1, 0x1a: 2, 0x1b: 2, 0x1c: 2, 0x1d: 1,
  0x1e: 2, 0x1f: 2, 0x20: 2, 0x21: 2, 0x22: 2, 0x23: 2, 0x24: 2, 0x25: 1, 0x26: 2, 0x27: 2, 0x28: 2, 0x29: 2, 0x2a: 2, 0x2b: 2,
  0x2c: 1, 0x2d: 2, 0x2e: 2, 0x2f: 2, 0x30: 2, 0x31: 2, 0x32: 1, 0x33: 1,
  0x34: 1, 0x35: 1, 0x36: 1, 0x37: 1, 0x38: 3, 0x39: 1, 0x3a: 0, 0x3c: 1, 0x3d: 1, 0x3e: 1, 0x3f: 1, 0x40: 1, 0x41: 1, 0x42: 1, 0x43: 1,
  0x44: 2, 0x45: 1, 0x46: 3, 0x47: 1, 0x48: 2, 0x49: 1, 0x4a: 1, 0x4b: 1, 0x4c: 1, 0x4d: 1, 0x4e: -1, 0x4f: 3, 0x50: 3, 0x53: -1, 0x5e: 1,
  0x5f: -1, 0x60: 2, 0x61: 2, 0x62: 2, 0x63: 2,
  0x75: 1, 0x76: 1, 0x77: 2, 0x78: 2, 0x79: 1, 0x7a: 1, 0x7b: 2, 0x7c: 2, 0x7d: 1, 0x7e: 1, 0x7f: 1, 0x80: 1, 0x81: 2, 0x82: 2, 0x83: 2,
  0x84: 2, 0x85: 1, 0x86: 1, 0x87: 2, 0x88: 2, 0x89: 1, 0x8a: 1, 0x8b: 2, 0x8c: 2, 0x8d: 2, 0x8e: 2, 0x8f: 2, 0x90: 2, 0x91: 2,
  0x92: 2, 0x93: 2, 0x94: 2, 0x95: 1, 0x98: -2, 0x99: 2, 0x9a: -1, 0x9b: -1, 0x9d: 5, 0x9e: 4, 0x9f: 2, 0xa0: 4, 0xa1: 2, 0xa2: 2,
  0xa3: 5, 0xa4: 4,
};

/** Whether `head` starts a Word for the Macintosh 4.0 or 5.x document. */
export function isMacWordFile(head: Uint8Array): boolean {
  return head.length >= 4 && head[0] === 0xfe && head[1] === 0x37 && head[2] === 0x00 && (head[3] === 0x1c || head[3] === 0x23);
}

/** A Word for the Macintosh 4.0 or 5.x document as an OldDocument. */
export function readMacWord(bytes: Uint8Array): OldDocument {
  return new MacReader(bytes).read();
}

export class MacReader extends Word2Reader {
  protected readonly sprmSizes = MAC_SPRM_SIZES;
  private readonly fontNames = new Map<number, string>();
  protected dop: Uint8Array = new Uint8Array(0);

  constructor(bytes: Uint8Array) {
    super(bytes, true);
  }

  /**
   * The FIB: the text's start and its lengths (main text, footnotes,
   * headers and footers), then from 0x40 twenty file positions with their
   * sizes; Word 5 has a second list after two counts, its font names second.
   */
  protected readFib(): Fib {
    const e = (i: number): Entry => ({ fc: this.u32(0x40 + 6 * i), cb: this.u16(0x44 + 6 * i) });
    this.dop = this.entry(e(15));
    if (this.bytes[3] === 0x23) this.readFontNames({ fc: this.u32(0xc2), cb: this.u16(0xc6) });
    return {
      fcMin: this.u32(0x14), ccpText: this.u32(0x24), ccpFtn: this.u32(0x28), ccpHdd: this.u32(0x2c),
      stsh: e(1).cb ? e(1) : e(0), fndRef: e(2), fndTxt: e(3), sed: e(4), hdd: e(8), bteChpx: e(9), btePapx: e(10), clx: e(18),
    };
  }

  /** The font names: a count, then per font two bytes, its number and its name as a Pascal string. */
  private readFontNames(entry: Entry): void {
    const zone = this.entry(entry);
    const count = this.n16(zone, 0);
    let at = 2;
    for (let i = 0; i < count && at + 5 <= zone.length; i++) {
      const length = zone[at + 4];
      this.fontNames.set(this.n16(zone, at + 2), MAC_ROMAN.decode(zone.subarray(at + 5, at + 5 + length)));
      at += 5 + length;
    }
  }

  /**
   * The document's properties: the page's height and width and its
   * margins (top, left, bottom, right), the default tab, then flags (facing
   * pages 0x80, where the notes go in bits 1-2, 0 the document's end) and
   * the note separators there are.
   */
  protected readDocProps(): DocProps {
    const dop = this.dop;
    const page = newSectionProps();
    if (dop.length >= 12) {
      const height = this.n16(dop, 0);
      const width = this.n16(dop, 2);
      const margins = [0, 1, 2, 3].map((k) => Math.abs(this.n16s(dop, 4 + 2 * k)));
      if (height > margins[0] + margins[2] && width > margins[1] + margins[3]) {
        Object.assign(page, { yaPage: height, xaPage: width, dyaTop: margins[0], dxaLeft: margins[1], dyaBottom: margins[2], dxaRight: margins[3], landscape: width > height });
      }
    }
    return {
      page,
      defaultTab: dop.length >= 16 ? this.n16(dop, 14) || undefined : undefined,
      facingPages: dop.length > 16 && (dop[16] & 0x80) !== 0,
      endnotes: dop.length > 16 && ((dop[16] >> 1) & 3) === 0,
      separatorStories: dop.length > 17 ? popcount(dop[17] & 0x3f) : 0,
    };
  }

  protected baseChar(): Char {
    return { fontId: 2, hps: 24, flags: new Array(8).fill(false), kul: 0, hpsPos: 0, dxaSpace: 0, ico: 0, special: true };
  }

  /**
   * A CHPX: the first bytes of a character's properties, as far as they
   * differ. Byte 0 flips bold, italic, strikethrough, outline, shadow, small
   * caps, all caps and hidden; byte 1 says which of the others are set; then
   * the font number, the size in half-points, the raise in half-points, the
   * spacing in sixteenths of a point, the color and underline, and for a
   * picture its file offset.
   */
  protected applyChpx(c: Char, d: Uint8Array): void {
    if (d.length === 0) return;
    for (let i = 0; i < 8; i++) if (d[0] & (0x80 >> i)) c.flags[i] = !c.flags[i];
    const what = d.length >= 2 ? d[1] : 0;
    if (d.length >= 4 && this.n16(d, 2)) c.fontId = this.n16(d, 2);
    if (d.length >= 5 && d[4]) c.hps = d[4];
    if (what & 0x02) c.hpsPos = d.length >= 6 ? (d[5] << 24) >> 24 : 0;
    if (what & 0x01) c.dxaSpace = d.length >= 7 ? Math.round((((d[6] << 24) >> 24) * 20) / 16) : 0;
    if (what & 0x20) c.ico = d.length >= 8 ? (d[7] >> 4) + 1 : 0;
    if (what & 0x04) {
      const kul = d.length >= 8 ? (d[7] >> 1) & 7 : 0;
      c.kul = kul && c.kul === kul ? 0 : kul;
    }
    if (d.length >= 10) c.picture = (d[9] << 16) | ((d[10] ?? 0) << 8) | (d[11] ?? 0);
  }

  protected isCharSprm(op: number): boolean {
    return (op >= 0x3c && op <= 0x53) || op === 0x5e;
  }

  /** A character sprm among a paragraph's or a piece's. Toggles: 0 off, 1 on, 0x81 the style's value, 0x80 its opposite. */
  protected applyCharSprm(c: Char, { op, arg }: Sprm, style: Char): void {
    if ((op >= 0x3c && op <= 0x43) || op === 0x45) {
      const v = arg[0];
      if (op === 0x45) {
        c.kul = v === 0x81 ? style.kul : v === 0x80 ? (style.kul ? 0 : 1) : v ? 1 : 0;
        return;
      }
      const i = op - 0x3c;
      c.flags[i] = v === 0x81 ? style.flags[i] : v === 0x80 ? !style.flags[i] : v !== 0;
      return;
    }
    switch (op) {
      case 0x44: c.fontId = this.n16(arg, 0); break;
      case 0x4a: if (arg[0]) c.hps = arg[0]; break;
      case 0x4d: c.hpsPos = (arg[0] << 24) >> 24; break;
      case 0x4e:
      case 0x53: this.applyChpx(c, arg); break;
      default: break;
    }
  }

  /** Word for the Macintosh numbers the standard styles one after Word for Windows 2.0: its -1 is heading 1, its -12 header. */
  protected standardStyleName(id: number): string | undefined {
    return STANDARD_STYLES[id - 1];
  }

  protected fontName(id: number): string {
    return this.fontNames.get(id) ?? MAC_FONTS[id] ?? "Times";
  }

  /**
   * How a font's character codes become text: the Symbol font's codes are
   * its Unicode characters; a pictograph font's (Wingdings, Zapf Dingbats and
   * its Monotype Sorts copy) stay codes, in the private-use block at U+F000 as
   * Word keeps them; any other font's are Mac OS Roman.
   */
  protected decoder(fontId: number): (code: number) => string {
    const name = this.fontName(fontId);
    if (name === "Symbol") return (code) => symbolText(Uint8Array.of(code));
    if (/wingdings|webdings|dingbats|sorts/i.test(name)) return (code) => String.fromCharCode(0xf000 + code);
    return (code) => MAC_ROMAN_CHARS[code];
  }

  /** A tab descriptor: the alignment in the top three bits, the leader in bits 2-3. */
  protected tab(type: number): { jc: number; leader: number } {
    return { jc: Math.min(type >> 5, 4), leader: (type >> 2) & 3 };
  }

  /** A piece's property modifier: with its top bit set, a grpprl's number; else a sprm and its byte. */
  protected prmSprms(prm: number, grpprls: Uint8Array[]): Sprm[] {
    if (prm & 0x8000) return this.parseSprms(grpprls[(prm >>> 8) & 0x7f] ?? new Uint8Array(0));
    if (prm >>> 8) return this.parseSprms(Uint8Array.of(prm >>> 8, prm & 0xff));
    return [];
  }

  protected styleName(bytes: Uint8Array): string {
    return MAC_ROMAN.decode(bytes);
  }

  /** The Mac's special characters: the page number, dates and times, the command key and Apple signs. */
  protected special(ch: number, cp: number, start: number, char: Char, noteStart: boolean, content: Inline[]): void {
    const date = DATE_PICTURES[ch];
    if (date !== undefined) content.push({ kind: "field", code: date ? `DATE \\@ "${date}"` : "DATE", result: "" });
    else if (ch === 0x02) content.push({ kind: "field", code: "PAGE", result: "1" });
    else if (ch === 0x04 || ch === 0x18 || ch === 0x19) content.push({ kind: "field", code: "TIME", result: "" });
    else if (ch === 0x11) content.push({ kind: "text", text: "⌘" });
    else if (ch === 0x14) content.push({ kind: "text", text: "" });
    else if (ch === 0x01 || ch === 0x05) super.special(ch, cp, start, char, noteStart, content);
  }

  /**
   * The picture a U+0001 shows, at a file offset: a header (size, count,
   * flags, the frame in points), then `count` parts, each a header of 16
   * bytes and a PICT. The first part with a PICT is the picture.
   */
  private readonly pictureIndex = new Map<number, number>();
  protected picture(char: Char): Inline | undefined {
    const at = char.picture ?? 0;
    const size = this.u32(at);
    if (at < 0x100 || size < 14 || at + size > this.bytes.length) return undefined;
    let index = this.pictureIndex.get(at);
    if (index === undefined) {
      const count = this.u8(at + 4);
      let part = at + 14;
      for (let k = 0; k < count && part + 16 <= at + size; k++) {
        const partSize = this.u32(part);
        if (partSize < 16 || part + partSize > at + size) break;
        if (partSize > 16 + 10) {
          index = this.pictures.length;
          this.pictures.push({ data: this.bytes.slice(part + 16, part + partSize), ext: "pict" });
          break;
        }
        part += partSize;
      }
      if (index === undefined) return undefined;
      this.pictureIndex.set(at, index);
    }
    let width = this.n16s(this.bytes, at + 12) - this.n16s(this.bytes, at + 8);
    let height = this.n16s(this.bytes, at + 10) - this.n16s(this.bytes, at + 6);
    if (width <= 0 || height <= 0) {
      // The PICT's own frame, at 72 dpi.
      const data = this.pictures[index].data;
      height = this.n16s(data, 6) - this.n16s(data, 2);
      width = this.n16s(data, 8) - this.n16s(data, 4);
    }
    return { kind: "picture", index, width: width * 20, height: height * 20 };
  }
}
