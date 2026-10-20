// Windows Write 1.0-3.1 (.wri, signature 31 BE or 32 BE, 00 00, 00 AB): a
// 128-byte header of 128-byte page numbers, little-endian; the text from
// 0x80 in the Windows ANSI code page; character and paragraph formatting
// in 128-byte FKP pages after the text; one section; a table of font
// names. Running heads are paragraphs at the start of the text, pictures
// paragraphs whose text is the picture: a Windows metafile, a bitmap, or
// (Write 3.1) an OLE 1.0 object.
//
// Written from Microsoft's "Microsoft Windows Write file format"
// description and checked against Write files from Windows 3.x-era discs.
import { ParaProps, RunProps, SectionProps, Tab, newParaProps, newSectionProps } from "../doc/sprm.js";
import { symbolText } from "../pdf/dib.js";
import { dibToWmf } from "./dibwmf.js";
import { Block, HeaderKind, Inline, OldDocument, Paragraph, Picture, Run, Section } from "./model.js";

const CP1252 = new TextDecoder("windows-1252");

/**
 * The Far Eastern code pages, told by a font name the system wrote in
 * them: a Write file has no code page of its own, it is the system's, and
 * a Japanese, Chinese or Korean system's fonts have names in its script.
 */
const FAR_EAST: [string, RegExp][] = [
  ["shift_jis", /明朝|ゴシック|ｺﾞｼｯｸ|ﾐﾝﾁｮｳ|書体/],
  ["gbk", /宋体|黑体|楷体|仿宋/],
  ["big5", /明體|楷體|黑體/],
  ["euc-kr", /바탕|돋움|굴림|궁서|명조|고딕/],
];

/** The standard VGA colors, for a 4-bit device-dependent bitmap: blue, green, red, intensity. */
const VGA: number[] = [
  0x000000, 0x000080, 0x008000, 0x008080, 0x800000, 0x800080, 0x808000, 0xc0c0c0,
  0x808080, 0x0000ff, 0x00ff00, 0x00ffff, 0xff0000, 0xff00ff, 0xffff00, 0xffffff,
];

/** Whether `head` (at least the 128-byte header) starts a Windows Write document rather than a Word for DOS one. */
export function isWriteFile(head: Uint8Array): boolean {
  if (head.length < 128) return false;
  const u16 = (at: number): number => head[at] | (head[at + 1] << 8);
  if (u16(2) !== 0 || u16(4) !== 0xab00) return false;
  if (u16(0) === 0xbe32) return true;
  // Word for DOS has the same start: Write has the file's page count at
  // 0x60, where Word for DOS has zero or a printer name, and nothing where
  // Word for DOS has its style sheet's name.
  return u16(0) === 0xbe31 && u16(0x60) !== 0 && u16(0x62) === 0 && head.subarray(0x1e, 0x60).every((b) => b === 0);
}

/** A Windows Write document as an OldDocument. */
export function readWrite(bytes: Uint8Array): OldDocument {
  return new WriteReader(bytes).read();
}

/** A character's formatting: CHP bytes 1-5. */
interface Chp {
  bold: boolean;
  italic: boolean;
  underline: boolean;
  /** The page number in a running head, for U+0001. */
  special: boolean;
  font: number;
  /** Half-points. */
  hps: number;
  /** 0 normal, 1 superscript, 2 subscript. */
  iss: number;
}

/** A paragraph's formatting: the PAP. */
interface Pap {
  jc: number;
  dxaRight: number;
  dxaLeft: number;
  dxaLeft1: number;
  dyaLine: number;
  /** Running head code: bits 1-2 a header or footer (bit 0 a footer), bit 3 on the first page too, bit 4 a picture. */
  rhc: number;
  tabs: Tab[];
}

interface Span<T> {
  fc: number;
  lim: number;
  props: T;
}

class WriteReader {
  private readonly fcMac: number;
  private readonly pnPara: number;
  private readonly pnFntb: number;
  private readonly pnSep: number;
  private readonly pnSetb: number;
  private readonly pnFfntb: number;
  private readonly pnMac: number;
  private readonly fonts: string[] = [];
  private readonly fontDecoders: ((raw: Uint8Array) => string)[] = [];
  private readonly pictures: Picture[] = [];
  private chps: Span<Chp>[] = [];
  private chpAt = 0;

  constructor(private readonly bytes: Uint8Array) {
    this.fcMac = Math.min(this.u32(14), bytes.length);
    this.pnPara = this.u16(18);
    this.pnFntb = this.u16(20);
    this.pnSep = this.u16(22);
    this.pnSetb = this.u16(24);
    this.pnFfntb = this.u16(28);
    this.pnMac = this.u16(0x60);
    if (this.fcMac < 0x80) throw new Error("the text ends before it starts");
  }

  private u16(at: number, b: Uint8Array = this.bytes): number {
    return (b[at] ?? 0) | ((b[at + 1] ?? 0) << 8);
  }
  private s16(at: number, b: Uint8Array = this.bytes): number {
    return (this.u16(at, b) << 16) >> 16;
  }
  private u32(at: number, b: Uint8Array = this.bytes): number {
    return (this.u16(at, b) | (this.u16(at + 2, b) << 16)) >>> 0;
  }

  read(): OldDocument {
    this.readFonts();
    this.chps = this.readFkps((this.fcMac + 127) >> 7, this.pnPara, (prop) => this.chp(prop));
    const paps = this.readFkps(this.pnPara, this.pnFntb, (prop) => this.pap(prop));
    const body: Block[] = [];
    const heads: { kind: HeaderKind; firstPage: boolean; blocks: Block[] }[] = [];
    const add = (rhc: number, paragraphs: Paragraph[]): void => {
      if (rhc & 0x06) {
        const kind: HeaderKind = rhc & 0x01 ? "footer" : "header";
        let head = heads.find((h) => h.kind === kind);
        if (!head) heads.push((head = { kind, firstPage: (rhc & 0x08) !== 0, blocks: [] }));
        head.blocks.push(...paragraphs);
      } else body.push(...paragraphs);
    };
    // A paragraph ends with CR LF. Write splits a long paragraph's formatting
    // into spans of up to 255 bytes, so a span can end inside a paragraph:
    // the span that holds its end gives its formatting.
    let start = 0x80;
    let last: Pap | undefined;
    for (const span of paps) {
      if (span.props.rhc & 0x10) {
        if (last && start < span.fc) add(last.rhc, [this.paragraph(last, start, span.fc)]);
        add(span.props.rhc, this.pictureParagraphs(span));
        start = span.lim;
        continue;
      }
      for (let fc = span.fc; fc < span.lim; fc++) {
        if (this.bytes[fc] !== 0x0a) continue;
        add(span.props.rhc, [this.paragraph(span.props, start, fc + 1)]);
        start = fc + 1;
      }
      last = span.props;
    }
    if (last && start < this.fcMac) add(last.rhc, [this.paragraph(last, start, this.fcMac)]);
    const props = this.sectionProps();
    // A running head shows on the first page only when Write was told so; the first page then has its own (empty) one.
    props.titlePage = heads.some((h) => !h.firstPage);
    const section: Section = { props, blocks: body, headers: [] };
    for (const head of heads) {
      section.headers.push({ kind: head.kind, type: "default", blocks: head.blocks });
      if (head.firstPage && props.titlePage) section.headers.push({ kind: head.kind, type: "first", blocks: head.blocks });
    }
    return {
      fonts: this.fonts,
      styles: [{ name: "Normal", para: newParaProps(0), run: { toggles: {} } }],
      sections: [section],
      footnotes: [],
      pictures: this.pictures,
    };
  }

  // ---- Fonts --------------------------------------------------------------

  /**
   * The font table: a count (which some files make include the page
   * markers, so the table is read to its end instead), then per font a size
   * (0 the end, 0xFFFF the rest on the next page), a family byte and the
   * name ending in a zero.
   */
  private readFonts(): void {
    const raw: Uint8Array[] = [];
    if (this.pnFfntb < this.pnMac) {
      const end = Math.min(this.pnMac * 128, this.bytes.length);
      let at = this.pnFfntb * 128 + 2;
      while (at + 3 <= end) {
        const cb = this.u16(at);
        if (cb === 0) break;
        if (cb === 0xffff) {
          at = ((at >> 7) + 1) << 7;
          continue;
        }
        const name = this.bytes.subarray(at + 3, Math.min(at + 2 + cb, end));
        const zero = name.indexOf(0);
        raw.push(zero >= 0 ? name.subarray(0, zero) : name);
        at += 2 + cb;
      }
    }
    // The document's code page: a Far Eastern one when a font's name is in its script.
    let codePage = "windows-1252";
    for (const name of raw) {
      if (name.every((b) => b < 0x80)) continue;
      const found = FAR_EAST.find(([label, words]) => words.test(new TextDecoder(label).decode(name)));
      if (found) codePage = found[0];
    }
    const text = new TextDecoder(codePage);
    for (const name of raw) {
      const decoded = text.decode(name);
      this.fonts.push(decoded);
      this.fontDecoders.push(fontDecoder(decoded, text));
    }
  }

  // ---- Formatting pages ---------------------------------------------------

  /**
   * The FKP pages from `pnFirst` up to `pnLim`: each the file position of
   * its first character, then (fcLim, bfprop) pairs, the property bytes at
   * 4 + bfprop (a size byte, then the first bytes of the CHP or PAP; 0xFFFF
   * the defaults), and the pair count in the last byte. Returned as spans
   * covering the text, the defaults where no page covers it.
   */
  private readFkps<T>(pnFirst: number, pnLim: number, parse: (prop: Uint8Array) => T): Span<T>[] {
    const spans: Span<T>[] = [];
    let fc = 0x80;
    const add = (lim: number, props: T): void => {
      lim = Math.min(lim, this.fcMac);
      if (lim > fc) spans.push({ fc, lim, props });
      fc = Math.max(fc, lim);
    };
    for (let pn = pnFirst; pn < pnLim && (pn + 1) * 128 <= this.bytes.length; pn++) {
      const page = this.bytes.subarray(pn * 128, (pn + 1) * 128);
      const first = this.u32(0, page);
      if (first > fc) add(first, parse(new Uint8Array(0)));
      const count = Math.min(page[127], 20);
      for (let i = 0; i < count; i++) {
        const lim = this.u32(4 + 6 * i, page);
        const bfprop = this.u16(8 + 6 * i, page);
        const at = 4 + bfprop;
        const prop = bfprop === 0xffff || at >= 127 ? new Uint8Array(0) : page.subarray(at + 1, Math.min(at + 1 + page[at], 127));
        add(lim, parse(prop));
      }
    }
    add(this.fcMac, parse(new Uint8Array(0)));
    return spans;
  }

  /** The CHP: byte 1 bold, italic and the font's low six bits; 2 the size; 3 underline and the special bit; 4 the font's high bits; 5 the raise. */
  private chp(prop: Uint8Array): Chp {
    const b = Uint8Array.of(1, 0, 24, 0, 0, 0);
    b.set(prop.subarray(0, 6));
    return {
      bold: (b[1] & 0x01) !== 0,
      italic: (b[1] & 0x02) !== 0,
      font: (b[1] >> 2) | ((b[4] & 0x07) << 6),
      hps: b[2] || 24,
      underline: (b[3] & 0x01) !== 0,
      special: (b[3] & 0x40) !== 0,
      // A raise of 1-127 half-points is superscript, of 128-255 subscript.
      iss: b[5] === 0 ? 0 : b[5] < 128 ? 1 : 2,
    };
  }

  /** The PAP: 1 the alignment, 4 the right indent, 6 the left, 8 the first line's, 10 the line spacing, 16 the running head code, 22 the tabs. */
  private pap(prop: Uint8Array): Pap {
    const b = new Uint8Array(78);
    b.set([61, 0, 0x1e, 0, 0, 0, 0, 0, 0, 0, 240, 0]);
    b.set(prop.subarray(0, 78));
    const tabs: Tab[] = [];
    // Up to 14 tabs of 4 bytes: the position, then the alignment (0 left, 3 decimal).
    for (let at = 22; at + 4 <= 78; at += 4) {
      const pos = this.u16(at, b);
      if (pos === 0) break;
      tabs.push({ pos, jc: b[at + 2] === 3 ? 3 : 0, leader: 0 });
    }
    return {
      jc: b[1] & 3,
      dxaRight: this.s16(4, b),
      dxaLeft: this.s16(6, b),
      dxaLeft1: this.s16(8, b),
      dyaLine: this.u16(10, b),
      rhc: b[16],
      tabs,
    };
  }

  /** The section's page: SEP bytes 3 on, page size, first page number, top margin, text height, left margin, text width, header and footer positions. */
  private sectionProps(): SectionProps {
    const b = new Uint8Array(23);
    const defaults = [0, 15840, 12240, 0xffff, 1440, 12960, 1800, 8640, 0, 1080, 14760];
    defaults.forEach((value, i) => {
      b[1 + 2 * i] = value & 0xff;
      b[2 + 2 * i] = value >> 8;
    });
    if (this.pnSep < this.pnSetb) {
      const at = this.pnSep * 128;
      const cch = this.bytes[at] ?? 0;
      b.set(this.bytes.subarray(at + 1, at + 1 + Math.min(cch, 22)), 1);
    }
    const [yaMac, xaMac, pgnFirst, dyaText, xaLeft, dxaText] = [3, 5, 7, 11, 13, 15].map((at) => this.u16(at, b));
    const yaTop = this.s16(9, b);
    const yaHeader = this.u16(19, b);
    const yaFooter = this.u16(21, b);
    const props = newSectionProps();
    props.yaPage = yaMac;
    props.xaPage = xaMac;
    if (xaMac > yaMac) props.landscape = true;
    // Some files (converted from Word) have Word's negative top margin, the text's top whatever the
    // header's height, and a text height that makes the bottom margin negative too: their sizes are the margins.
    props.dyaTop = Math.abs(yaTop);
    props.dyaBottom = Math.abs(yaMac - yaTop - dyaText);
    props.dxaLeft = xaLeft;
    if (xaMac >= xaLeft + dxaText) props.dxaRight = xaMac - xaLeft - dxaText;
    props.dyaHdrTop = yaHeader;
    if (yaMac > yaFooter) props.dyaHdrBottom = yaMac - yaFooter;
    if (pgnFirst !== 0xffff && pgnFirst !== 1) {
      props.pgnRestart = true;
      props.pgnStart = pgnFirst;
    }
    return props;
  }

  // ---- Paragraphs ---------------------------------------------------------

  private paragraph(pap: Pap, start: number, end: number, picture?: Inline): Paragraph {
    const props: ParaProps = newParaProps(0);
    if (pap.jc) props.jc = pap.jc;
    if (pap.dxaLeft) props.dxaLeft = pap.dxaLeft;
    if (pap.dxaRight) props.dxaRight = pap.dxaRight;
    if (pap.dxaLeft1) props.dxaLeft1 = pap.dxaLeft1;
    if (pap.dyaLine > 0 && pap.dyaLine !== 240) {
      props.lineSpacing = pap.dyaLine;
      props.lineMultiple = true;
    }
    props.tabs = pap.tabs;
    const runs: Run[] = [];
    let mark: RunProps = { toggles: {} };
    if (picture) {
      const chp = this.chpFor(start);
      mark = this.runProps(chp);
      runs.push({ props: mark, content: [picture] });
    } else {
      for (let fc = start; fc < end; ) {
        const span = this.chpSpan(fc);
        const lim = Math.min(span.lim, end);
        mark = this.runProps(span.props);
        const content = this.content(fc, lim, span.props);
        if (content.length) runs.push({ props: mark, content });
        fc = lim;
      }
    }
    return { type: "paragraph", props, mark, runs };
  }

  /** The CHP span holding `fc`; paragraphs come in order, so the search goes on from the last one. */
  private chpSpan(fc: number): Span<Chp> {
    if (this.chpAt >= this.chps.length || this.chps[this.chpAt].fc > fc) this.chpAt = 0;
    while (this.chpAt < this.chps.length - 1 && this.chps[this.chpAt].lim <= fc) this.chpAt++;
    return this.chps[this.chpAt];
  }

  private chpFor(fc: number): Chp {
    return this.chpSpan(fc).props;
  }

  private runProps(chp: Chp): RunProps {
    const props: RunProps = { toggles: {}, hps: chp.hps };
    if (chp.font < this.fonts.length) props.ftc0 = props.ftc1 = props.ftc2 = chp.font;
    if (chp.bold) props.toggles.b = 1;
    if (chp.italic) props.toggles.i = 1;
    if (chp.underline) props.kul = 1;
    if (chp.iss) props.iss = chp.iss;
    return props;
  }

  /** The text from `start` to `end` in the run's font: CR and LF end the paragraph, U+0001 with the special bit is the page number. */
  private content(start: number, end: number, chp: Chp): Inline[] {
    const out: Inline[] = [];
    const decode = this.fontDecoders[chp.font] ?? ((raw: Uint8Array) => CP1252.decode(raw));
    let textStart = start;
    const flush = (fc: number): void => {
      if (fc > textStart) out.push({ kind: "text", text: decode(this.bytes.subarray(textStart, fc)) });
      textStart = fc + 1;
    };
    for (let fc = start; fc < end; fc++) {
      const ch = this.bytes[fc];
      if (ch >= 0x20) continue;
      flush(fc);
      switch (ch) {
        case 0x09: out.push({ kind: "tab" }); break;
        case 0x0b: out.push({ kind: "break" }); break;
        case 0x0c: out.push({ kind: "page" }); break;
        case 0x1f: out.push({ kind: "softHyphen" }); break;
        case 0x01: if (chp.special) out.push({ kind: "field", code: "PAGE", result: "1" }); break;
        default: break;
      }
    }
    flush(end);
    return out;
  }

  // ---- Pictures -----------------------------------------------------------

  /**
   * A picture paragraph's text is its picture: a 40-byte header (the
   * mapping mode at 0, the offset from the left margin at 8, the size in
   * twips at 10, a BITMAP at 16, the header's size at 30, the data's at 32,
   * the scale in thousandths at 36) and the data, then CR LF.
   */
  private pictureParagraphs(span: Span<Pap>): Paragraph[] {
    const out: Paragraph[] = [];
    let at = span.fc;
    while (at + 40 <= span.lim) {
      const mm = this.u16(at);
      const cbHeader = this.u16(at + 30) || 40;
      // An OLE object's size is at 16, where a picture has its BITMAP.
      const size = mm === 0xe4 ? this.u32(at + 16) : this.u32(at + 32);
      const data = this.bytes.subarray(at + cbHeader, Math.min(at + cbHeader + size, span.lim));
      const pap = { ...span.props, dxaLeft: span.props.dxaLeft + this.s16(at + 8) };
      out.push(this.paragraph(pap, at, at, this.picture(at, mm, data)));
      at += cbHeader + size;
      while (at < span.lim && (this.bytes[at] === 0x0d || this.bytes[at] === 0x0a)) at++;
    }
    return out;
  }

  private picture(at: number, mm: number, data: Uint8Array): Inline | undefined {
    const mx = (this.u16(at + 36) || 1000) / 1000;
    const my = (this.u16(at + 38) || 1000) / 1000;
    let width = this.u16(at + 10);
    let height = this.u16(at + 12);
    let picture: Picture | undefined;
    if (mm === 0xe3) {
      const bitmap = this.bytes.subarray(at + 16, at + 30);
      picture = ddbPicture(bitmap, data);
      // A bitmap's size is its pixels at the screen's 96 dots to the inch.
      if (!width || !height) {
        width = this.u16(2, bitmap) * 15;
        height = this.u16(4, bitmap) * 15;
      }
    } else if (mm === 0xe4) {
      picture = olePicture(data);
    } else if (data.length >= 18) {
      picture = { data: data.slice(), ext: "wmf" };
    }
    if (!picture || !width || !height) return undefined;
    this.pictures.push(picture);
    return { kind: "picture", index: this.pictures.length - 1, width: width * mx, height: height * my };
  }
}

/** How a font's character codes become text: the Symbol font's are its Unicode characters, a pictograph font's stay codes at U+F000, any other's are the document's code page. */
function fontDecoder(name: string, text: TextDecoder): (raw: Uint8Array) => string {
  if (/^symbol$/i.test(name)) return (raw) => symbolText(raw);
  if (/wingdings|webdings|dingbats|sorts/i.test(name)) return (raw) => Array.from(raw, (b) => String.fromCharCode(0xf000 + b)).join("");
  return (raw) => text.decode(raw);
}

/**
 * An OLE 1.0 object: a version, a format (2 embedded, 3 static) and a
 * class name; an embedded object's topic, item and native data, then its
 * presentation (version, format 5, class name); then the presentation's
 * width, height and size, and its data: a METAFILEPICT and a metafile, a
 * DIB, or a BITMAP and its bits. An embedded Paintbrush picture with no
 * presentation is its native data, a .bmp file.
 */
function olePicture(data: Uint8Array): Picture | undefined {
  const v = new DataView(data.buffer, data.byteOffset, data.byteLength);
  let at = 4;
  const u32 = (): number => {
    const value = at + 4 <= data.length ? v.getUint32(at, true) : 0;
    at += 4;
    return value;
  };
  const name = (): string => {
    const n = u32();
    const s = CP1252.decode(data.subarray(at, Math.min(at + n, data.length))).replace(/\0+$/, "");
    at += n;
    return s;
  };
  const format = u32();
  let className = name();
  if (format === 2) {
    name();
    name();
    const n = u32();
    const native = data.subarray(at, Math.min(at + n, data.length));
    at += n + 4;
    if (u32() === 0) return native[0] === 0x42 && native[1] === 0x4d ? dibPicture(native.subarray(14)) : undefined;
    className = name();
  } else if (format !== 3) return undefined;
  at += 8;
  const size = u32();
  const bits = data.subarray(at, Math.min(at + size, data.length));
  switch (className) {
    case "METAFILEPICT": return bits.length > 8 + 18 ? { data: bits.slice(8), ext: "wmf" } : undefined;
    case "DIB": return dibPicture(bits);
    case "BITMAP": return ddbPicture(bits.subarray(0, 14), bits.subarray(14));
    default: return undefined;
  }
}

/**
 * A device-dependent bitmap (a BITMAP: width, height, bytes a row, planes,
 * bits a pixel; then its rows from the top) as a DIB: monochrome, or 16
 * colors in the VGA palette, packed or in four planes (each plane's rows in
 * turn: blue, green, red, intensity).
 */
function ddbPicture(bitmap: Uint8Array, bits: Uint8Array): Picture | undefined {
  if (bitmap.length < 10) return undefined;
  const u16 = (at: number): number => bitmap[at] | (bitmap[at + 1] << 8);
  const width = u16(2);
  const height = u16(4);
  const rowBytes = u16(6);
  const planes = bitmap[8];
  const bitsPixel = bitmap[9];
  const depth = planes * bitsPixel;
  if (!width || !height || (depth !== 1 && depth !== 4) || (planes !== 1 && bitsPixel !== 1)) return undefined;
  if (bits.length < rowBytes * height * planes || rowBytes * 8 < width * bitsPixel) return undefined;
  const stride = (((width * depth + 31) >> 5) << 2);
  const colors = depth === 1 ? [0x000000, 0xffffff] : VGA;
  const dib = new Uint8Array(40 + 4 * colors.length + stride * height);
  const v = new DataView(dib.buffer);
  v.setUint32(0, 40, true);
  v.setInt32(4, width, true);
  v.setInt32(8, height, true);
  v.setUint16(12, 1, true);
  v.setUint16(14, depth, true);
  v.setUint32(20, stride * height, true);
  colors.forEach((rgb, i) => v.setUint32(40 + 4 * i, rgb, true));
  const pixels = 40 + 4 * colors.length;
  for (let y = 0; y < height; y++) {
    // A DIB's rows go from the bottom.
    const row = pixels + (height - 1 - y) * stride;
    if (planes === 1) {
      dib.set(bits.subarray(y * rowBytes, y * rowBytes + Math.min(rowBytes, stride)), row);
      continue;
    }
    for (let x = 0; x < width; x++) {
      let index = 0;
      for (let p = 0; p < 4; p++) index |= ((bits[(p * height + y) * rowBytes + (x >> 3)] >> (7 - (x & 7))) & 1) << p;
      dib[row + (x >> 1)] |= x & 1 ? index : index << 4;
    }
  }
  return dibPicture(dib);
}

/** A device-independent bitmap as a picture: a metafile that draws it, which the PDF renderer draws. */
function dibPicture(dib: Uint8Array): Picture | undefined {
  const wmf = dibToWmf(dib);
  return wmf ? { data: wmf, ext: "wmf" } : undefined;
}
