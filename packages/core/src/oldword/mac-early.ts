// Word for the Macintosh 1.0x (signature FE 32 00 00) and 3.0x (FE 34 00 00).
//
// Word 3 is the Word 4 format mac.ts reads, with a shorter header: fifteen
// file positions at 0x1E, the text's lengths at 0x7C and the document's
// properties at 0xAC; FKP pages of 128 bytes; three bytes before a PAPX's
// sprms; section sprms numbered 0x36-0x45.
//
// Word 1 is laid out as Windows Write is, big-endian: a header of 128 bytes,
// the text from 0x80, character and paragraph properties in 128-byte pages
// of runs, each run a prefix of a fixed property record; then the footnote
// table, the section's properties, the section table and the page table.
// Headers and footers are paragraphs of the text with a flag, and a picture
// is a paragraph whose bytes are a PICT.
//
// Written from the file formats as observed in Word's files, with the
// libmwaw project's notes on them (MsWrd1Parser, MsWrdParser) as a guide to
// the structures.
import { ParaProps, RunProps, SectionProps, newParaProps, newSectionProps } from "../doc/sprm.js";
import { symbolText } from "../pdf/dib.js";
import { MAC_FONTS, MAC_SPRM_SIZES, MacReader } from "./mac.js";
import { Block, HeaderKind, HeaderType, Inline, OldDocument, Paragraph, Picture, Run, Section } from "./model.js";
import { Char, Entry, Fib, Sprm } from "./word2.js";

const MAC_ROMAN = new TextDecoder("macintosh");
const MAC_ROMAN_CHARS: string[] = Array.from({ length: 256 }, (_, i) => MAC_ROMAN.decode(Uint8Array.of(i)));

/** Whether `head` starts a Word for the Macintosh 1.0x or 3.0x document. */
export function isMacEarlyFile(head: Uint8Array): boolean {
  return head.length >= 4 && head[0] === 0xfe && (head[1] === 0x32 || head[1] === 0x34) && head[2] === 0 && head[3] === 0;
}

/** A Word for the Macintosh 1.0x or 3.0x document as an OldDocument. */
export function readMacEarly(bytes: Uint8Array): OldDocument {
  return bytes[1] === 0x32 ? new Word1Reader(bytes).read() : new Word3Reader(bytes).read();
}

// ---- Word 3 ------------------------------------------------------------------

/** Word 3's section sprms (0x36-0x45) by the Word 4 sprm that says the same. */
const WORD3_SECTION_SPRMS: Record<number, number> = {
  0x36: 0x75, 0x37: 0x76, 0x38: 0x77, 0x39: 0x78, 0x3b: 0x7a, 0x41: 0x80, 0x44: 0x83, 0x45: 0x84,
};

/** Word 3's page number formats (1 arabic, 2 upper roman, 3 upper letter, 4 lower letter, 5 lower roman) as Word's. */
const WORD3_PAGE_NUMBERS: Record<number, number> = { 1: 0, 2: 1, 3: 3, 4: 4, 5: 2 };

class Word3Reader extends MacReader {
  protected readonly fkpPageSize = 128;
  protected readonly papxHeader = 3;
  /** Word 4's sizes, with Word 3's section sprms in the place of Word 4's character ones (Word 3 keeps no character sprm in a paragraph). */
  protected readonly sprmSizes: Record<number, number> = {
    ...MAC_SPRM_SIZES,
    0x36: 1, 0x37: 1, 0x38: 2, 0x39: 2, 0x3a: 1, 0x3b: 1, 0x3c: 2, 0x3d: 2, 0x3e: 1, 0x3f: 1, 0x40: 1, 0x41: 1, 0x42: 2, 0x43: 2, 0x44: 2, 0x45: 2,
  };

  /** The header: fifteen file positions with their sizes from 0x1E (Word 4's first fifteen), the text from 0x100, its lengths from 0x7C, the document's properties from 0xAC. */
  protected readFib(): Fib {
    const e = (i: number): Entry => ({ fc: this.u32(0x1e + 6 * i), cb: this.u16(0x22 + 6 * i) });
    this.dop = this.bytes.subarray(0xac, 0x100);
    return {
      fcMin: 0x100, ccpText: this.u32(0x7c), ccpFtn: this.u32(0x80), ccpHdd: this.u32(0x84),
      stsh: e(1).cb ? e(1) : e(0), fndRef: e(2), fndTxt: e(3), sed: e(4), hdd: e(8), bteChpx: e(9), btePapx: e(10), clx: e(14),
    };
  }

  /** The document, its first page numbered as the document's properties say (at 0x16 of them). */
  read(): OldDocument {
    const doc = super.read();
    const first = this.n16(this.dop, 0x16);
    if (first > 1 && first < 0x8000) Object.assign(doc.sections[0].props, { pgnStart: first, pgnRestart: true });
    return doc;
  }

  protected isCharSprm(): boolean {
    return false;
  }

  protected sectionSprms(sepx: Uint8Array): Sprm[] {
    const out: Sprm[] = [];
    for (const { op, arg } of this.parseSprms(sepx)) {
      const word4 = WORD3_SECTION_SPRMS[op];
      if (word4 === 0x7a) out.push({ op: word4, arg: Uint8Array.of(WORD3_PAGE_NUMBERS[arg[0]] ?? 0) });
      else if (word4) out.push({ op: word4, arg });
    }
    return out;
  }

  /** Word 3's formula character (the backslash its formulas start with), then the Mac's special characters. */
  protected special(ch: number, cp: number, start: number, char: Char, noteStart: boolean, content: Inline[]): void {
    if (ch === 0x06) content.push({ kind: "text", text: "\\" });
    else super.special(ch, cp, start, char, noteStart, content);
  }
}

// ---- Word 1 ------------------------------------------------------------------

/** A character's formatting: the font, its size in half-points, and the rest. */
interface Word1Char {
  font: number;
  hps: number;
  bold: boolean;
  italic: boolean;
  underline: boolean;
  caps: boolean;
  outline: boolean;
  shadow: boolean;
  /** Half-points up (negative: down). */
  hpsPos: number;
  /** A page number (U+0001) or a note reference (U+0005), not text. */
  special: boolean;
}

/** A paragraph's properties: ParaProps, the header or footer it belongs to, whether it is a picture. */
interface Word1Para {
  props: ParaProps;
  /** The running-head flags: 0x10 footer, 0x20 odd pages, 0x40 even pages, 0x80 the first page; 0 for the body. */
  running: number;
  picture: boolean;
}

/** A run of the CHP or PAP pages: the file positions [fc, end) and the property bytes after their size byte. */
interface Word1Run {
  fc: number;
  end: number;
  data?: Uint8Array;
}

/** The character formatting of text with no properties of its own: Geneva, 12 points. */
const WORD1_DEFAULT_CHAR: Word1Char = { font: 3, hps: 24, bold: false, italic: false, underline: false, caps: false, outline: false, shadow: false, hpsPos: 0, special: false };

class Word1Reader {
  private readonly textEnd: number;
  /** The page numbers (128-byte pages) of the tables after the text: paragraph properties, footnotes, the section's properties, the section table, the page table, the end. */
  private readonly pn: number[];
  private chpRuns: Word1Run[] = [];
  private papRuns: Word1Run[] = [];
  private readonly fonts: string[] = [];
  private readonly fontIndex = new Map<number, number>();
  private readonly pictures: Picture[] = [];
  /** The footnote references in the text (file positions) by note number. */
  private readonly noteRefs = new Map<number, number>();

  constructor(private readonly bytes: Uint8Array) {
    this.textEnd = Math.min(this.u32(0x0e), bytes.length);
    this.pn = [0, 1, 2, 3, 4, 5].map((i) => this.u16(0x12 + 2 * i));
  }

  private u8(at: number): number {
    return this.bytes[at] ?? 0;
  }
  private u16(at: number): number {
    return (this.u8(at) << 8) | this.u8(at + 1);
  }
  private i16(at: number): number {
    return (this.u16(at) << 16) >> 16;
  }
  private u32(at: number): number {
    return ((this.u16(at) << 16) | this.u16(at + 2)) >>> 0;
  }

  read(): OldDocument {
    if (this.textEnd < 0x80) throw new Error("the text ends before it begins");
    this.chpRuns = this.readPages(Math.ceil(this.textEnd / 128), this.pn[0]);
    this.papRuns = this.readPages(this.pn[0], this.pn[1]);
    // The notes: their texts follow the body, each from its start to the next one's.
    const notes = this.readNotes();
    const bodyEnd = notes.length ? notes[0].start : this.textEnd;
    const footnotes = notes.map((note, i) => this.blocks(note.start, i + 1 < notes.length ? notes[i + 1].start : this.textEnd));
    const sections = this.body(bodyEnd);
    return {
      fonts: this.fonts,
      styles: [{ name: "Normal", para: newParaProps(0), run: this.fullRunProps(WORD1_DEFAULT_CHAR) }],
      sections,
      footnotes,
      pictures: this.pictures,
      evenAndOddHeaders: sections.some((s) => s.headers.some((h) => h.type === "even")),
      endnotes: (this.u8(this.pn[2] * 128 + 17) & 1) !== 0 && this.pn[3] > this.pn[2],
    };
  }

  /**
   * The runs of the 128-byte property pages [first, last): each page the
   * file position its runs start at, then per run its end and the offset
   * (from byte 4) of its properties, 0xFFFF for none; its run count last.
   */
  private readPages(first: number, last: number): Word1Run[] {
    const runs: Word1Run[] = [];
    for (let pn = first; pn < last && (pn + 1) * 128 <= this.bytes.length; pn++) {
      const page = pn * 128;
      const count = this.u8(page + 127);
      if (4 + 6 * count > 127) continue;
      let fc = this.u32(page);
      for (let k = 0; k < count; k++) {
        const end = this.u32(page + 4 + 6 * k);
        const offset = this.u16(page + 8 + 6 * k);
        let data: Uint8Array | undefined;
        if (offset !== 0xffff && 4 + offset < 127) {
          const at = page + 4 + offset;
          data = this.bytes.subarray(at + 1, Math.min(at + 1 + this.u8(at), page + 127));
        }
        if (end > fc) runs.push({ fc, end, data });
        fc = end;
      }
    }
    return runs;
  }

  private runAt(runs: Word1Run[], fc: number): Word1Run | undefined {
    let lo = 0;
    let hi = runs.length - 1;
    while (lo <= hi) {
      const mid = (lo + hi) >> 1;
      if (fc < runs[mid].fc) hi = mid - 1;
      else if (fc >= runs[mid].end) lo = mid + 1;
      else return runs[mid];
    }
    return undefined;
  }

  /**
   * The footnote table: two counts, then per note the file positions (less
   * 0x80) of its reference and of its text; a last entry past the text ends
   * the list. The notes in the order of their texts.
   */
  private readNotes(): { start: number }[] {
    if (this.pn[2] <= this.pn[1]) return [];
    const at = this.pn[1] * 128;
    const count = this.u16(at);
    const notes: { ref: number; start: number }[] = [];
    for (let i = 0; i < count && at + 12 + 8 * i <= this.bytes.length; i++) {
      const ref = this.u32(at + 4 + 8 * i) + 0x80;
      const start = this.u32(at + 8 + 8 * i) + 0x80;
      if (ref < this.textEnd && start < this.textEnd && start > 0x80) notes.push({ ref, start });
    }
    notes.sort((a, b) => a.start - b.start);
    notes.forEach((note, i) => this.noteRefs.set(note.ref, i + 1));
    return notes;
  }

  /**
   * The section's properties: a flags byte (the break in bits 5-7, the page
   * number format in bits 2-4), the page's height and width, the first page
   * number (-1: go on), the top margin and the text's height, the left margin
   * and the text's width, a notes flag, the columns, the header's distance
   * from the page's top and the footer's, the space between columns.
   */
  private sectionProps(): SectionProps {
    const props = newSectionProps();
    if (this.pn[3] <= this.pn[2]) return props;
    const at = this.pn[2] * 128;
    const flags = this.u8(at + 2);
    const height = this.u16(at + 3);
    const width = this.u16(at + 5);
    const top = this.u16(at + 9);
    const textHeight = this.u16(at + 11);
    const left = this.u16(at + 13);
    const textWidth = this.u16(at + 15);
    if (height >= top + textHeight && width >= left + textWidth && textHeight > 0 && textWidth > 0) {
      Object.assign(props, { yaPage: height, xaPage: width, dyaTop: top, dyaBottom: height - top - textHeight, dxaLeft: left, dxaRight: width - left - textWidth, landscape: width > height });
      const header = this.u16(at + 19);
      const footer = this.u16(at + 21);
      if (header > 0 && header < height) props.dyaHdrTop = header;
      if (footer > 0 && footer < height) props.dyaHdrBottom = height - footer;
    }
    props.bkc = Math.min(flags >> 5, 4);
    props.nfcPgn = Math.min((flags >> 2) & 7, 4);
    const firstPage = this.i16(at + 7);
    if (firstPage > 0) Object.assign(props, { pgnStart: firstPage, pgnRestart: true });
    const columns = this.u8(at + 18);
    if (columns > 1 && columns < 10) Object.assign(props, { columns, dxaColumns: this.u16(at + 23) });
    return props;
  }

  /**
   * The body [0x80, end): its paragraphs, and its headers and footers. One
   * that comes after the body's text has begun takes effect from there: it
   * starts a new section, on the same page.
   */
  private body(end: number): Section[] {
    const props = this.sectionProps();
    const sections: Section[] = [];
    let section: Section = { props, blocks: [], headers: [] };
    let running: { flags: number; blocks: Block[] } | undefined;
    for (const { paragraph, flags } of this.paragraphs(0x80, end, false)) {
      if (!flags) {
        running = undefined;
        section.blocks.push(paragraph);
        continue;
      }
      if (running?.flags !== flags) {
        if (section.blocks.length) {
          sections.push(section);
          section = { props: { ...structuredClone(props), bkc: 0 }, blocks: [], headers: [] };
        }
        running = { flags, blocks: [] };
        this.addRunning(section.headers, flags, running.blocks);
      }
      running.blocks.push(this.fromPageEdge(paragraph, props));
    }
    sections.push(section);
    for (const s of sections) if (s.headers.some((h) => h.type === "first")) s.props.titlePage = true;
    return sections;
  }

  /** The footnote's paragraphs [start, end). */
  private blocks(start: number, end: number): Block[] {
    return this.paragraphs(start, end, true).map((p) => p.paragraph);
  }

  /** The paragraphs of the text [start, end), with their running-head flags. */
  private paragraphs(start: number, end: number, note: boolean): { paragraph: Paragraph; flags: number }[] {
    const out: { paragraph: Paragraph; flags: number }[] = [];
    let pStart = start;
    while (pStart < end) {
      const pap = this.runAt(this.papRuns, pStart);
      // A paragraph ends after its mark; a picture's, where its properties end.
      const para = this.paragraphProps(pap?.data);
      let mark = pStart;
      if (para.picture && pap) mark = Math.min(pap.end, end);
      else {
        while (mark < end && this.bytes[mark] !== 0x0d) mark++;
        mark = Math.min(mark + 1, end);
      }
      const paragraph = para.picture ? this.pictureParagraph(pStart, mark, para.props) : this.paragraph(pStart, mark, para.props, note && pStart === start, mark === end);
      out.push({ paragraph, flags: note ? 0 : para.running });
      pStart = mark;
    }
    return out;
  }

  /**
   * A header or footer paragraph, whose indents and tabs Word 1 measures
   * from the paper's edges, measured from the margins as a .docx's are.
   */
  private fromPageEdge(paragraph: Paragraph, page: SectionProps): Paragraph {
    const left = page.dxaLeft ?? 0;
    const right = page.dxaRight ?? 0;
    const p = paragraph.props;
    const props: ParaProps = { ...p, dxaLeft: (p.dxaLeft ?? 0) - left, dxaRight: (p.dxaRight ?? 0) - right, tabs: p.tabs.map((t) => ({ ...t, pos: t.pos - left })) };
    return { ...paragraph, props };
  }

  /**
   * A header or footer's blocks under each kind of page its flags name (odd
   * pages are the default ones), in the place of an earlier one there; a
   * kind of page it leaves out, and no earlier one names, gets an empty one.
   */
  private addRunning(headers: Section["headers"], flags: number, blocks: Block[]): void {
    const kind: HeaderKind = flags & 0x10 ? "footer" : "header";
    const absent = (type: HeaderType): boolean => !headers.some((h) => h.kind === kind && h.type === type);
    const put = (type: HeaderType, own: boolean): void => {
      if (!own && !absent(type)) return;
      const index = headers.findIndex((h) => h.kind === kind && h.type === type);
      if (index >= 0) headers.splice(index, 1);
      headers.push({ kind, type, blocks: own ? blocks : [] });
    };
    const odd = (flags & 0x20) !== 0;
    const even = (flags & 0x40) !== 0;
    put("default", odd);
    // Odd and even pages alike need no even header, unless an earlier one has one.
    if (!(odd && even && absent("even"))) put("even", even);
    put("first", (flags & 0x80) !== 0);
  }

  /**
   * A paragraph's properties, a prefix of: a byte of unknown use, the
   * alignment (bits 6-7) and keep flags (0x10 with the next, 0x20 its lines),
   * two bytes, the right, left and first-line indents, the line spacing, the
   * space before and after, the running-head and picture flags, five bytes,
   * then tabs of four bytes each: the position, and the alignment (bits 5-6)
   * and leader (bits 2-3).
   */
  private paragraphProps(d: Uint8Array | undefined): Word1Para {
    const props = newParaProps(0);
    const out: Word1Para = { props, running: 0, picture: false };
    if (!d) return out;
    const s16 = (at: number): number => (d.length >= at + 2 ? (((d[at] << 8) | d[at + 1]) << 16) >> 16 : 0);
    if (d.length > 1) {
      if (d[1] >> 6) props.jc = d[1] >> 6;
      if (d[1] & 0x10) props.keepNext = true;
      if (d[1] & 0x20) props.keepLines = true;
    }
    if (s16(4)) props.dxaRight = s16(4);
    if (s16(6)) props.dxaLeft = s16(6);
    if (s16(8)) props.dxaLeft1 = s16(8);
    // Positive: at least; negative: exactly; 0: single.
    if (s16(10)) Object.assign(props, { lineSpacing: s16(10), lineMultiple: false });
    if (s16(12)) props.dyaBefore = s16(12);
    if (s16(14)) props.dyaAfter = s16(14);
    if (d.length > 16) {
      out.running = d[16] & 0xe0 ? d[16] & 0xf0 : 0;
      out.picture = (d[16] & 0x08) !== 0;
    }
    for (let at = 22; at + 3 <= d.length; at += 4) {
      const pos = s16(at);
      if (pos <= 0) continue;
      props.tabs.push({ pos, jc: (d[at + 2] >> 5) & 3, leader: (d[at + 2] >> 2) & 3 });
    }
    return out;
  }

  /**
   * A character's properties, a prefix of: a byte of unknown use; bold
   * (0x80), italic (0x40) and the Mac font number; the size in half-points;
   * underline (0x80), all caps (0x0C) and a special character (0x02);
   * outline (0x10) and shadow (0x08); the raise in half-points.
   */
  private charAt(fc: number): Word1Char {
    const char = { ...WORD1_DEFAULT_CHAR };
    const d = this.runAt(this.chpRuns, fc)?.data;
    if (!d) return char;
    if (d.length > 1) {
      char.bold = (d[1] & 0x80) !== 0;
      char.italic = (d[1] & 0x40) !== 0;
      if (d[1] & 0x3f) char.font = d[1] & 0x3f;
    }
    if (d.length > 2 && d[2]) char.hps = d[2];
    if (d.length > 3) {
      char.underline = (d[3] & 0x80) !== 0;
      char.caps = (d[3] & 0x0c) === 0x0c;
      char.special = (d[3] & 0x02) !== 0;
    }
    if (d.length > 4) {
      char.outline = (d[4] & 0x10) !== 0;
      char.shadow = (d[4] & 0x08) !== 0;
    }
    if (d.length > 5) char.hpsPos = (d[5] << 24) >> 24;
    return char;
  }

  private paragraph(start: number, end: number, props: ParaProps, noteStart: boolean, last: boolean): Paragraph {
    const runs: Run[] = [];
    // The text, without its mark.
    const textEnd = this.bytes[end - 1] === 0x0d ? end - 1 : end;
    // A page break that ends the last paragraph only marks the end of the text.
    const endMark = last && this.bytes[textEnd - 1] === 0x0c ? textEnd - 1 : -1;
    let cp = start;
    while (cp < textEnd) {
      const chp = this.runAt(this.chpRuns, cp);
      const runEnd = Math.min(textEnd, chp ? chp.end : textEnd);
      const char = this.charAt(cp);
      runs.push(this.run(cp, Math.max(runEnd, cp + 1), char, noteStart && cp === start, endMark));
      cp = Math.max(runEnd, cp + 1);
    }
    return { type: "paragraph", props, mark: this.runProps(this.charAt(Math.max(start, textEnd))), runs: runs.filter((r) => r.content.length) };
  }

  private run(start: number, end: number, char: Word1Char, noteStart: boolean, endMark: number): Run {
    const content: Inline[] = [];
    let text = "";
    const flush = (): void => {
      if (text) content.push({ kind: "text", text });
      text = "";
    };
    const decode = this.decoder(char.font);
    for (let at = start; at < end; at++) {
      const ch = this.bytes[at];
      if (ch >= 0x20) {
        text += decode(ch);
        continue;
      }
      flush();
      switch (ch) {
        case 0x09: content.push({ kind: "tab" }); break;
        case 0x0b: content.push({ kind: "break" }); break;
        case 0x0c: if (at !== endMark) content.push({ kind: "page" }); break;
        case 0x1f: content.push({ kind: "softHyphen" }); break;
        case 0x01: if (char.special) content.push({ kind: "field", code: "PAGE", result: "1" }); break;
        case 0x05: {
          if (!char.special) break;
          const id = this.noteRefs.get(at);
          if (id) content.push({ kind: "footnote", id });
          else if (noteStart && at === start) content.push({ kind: "footnoteMark" });
          break;
        }
        default: break;
      }
    }
    flush();
    return { props: this.runProps(char), content };
  }

  /** A picture paragraph: six bytes, then a PICT, its frame in points. */
  private pictureParagraph(start: number, end: number, props: ParaProps): Paragraph {
    const at = start + 6;
    const runs: Run[] = [];
    if (end - at > 10) {
      const height = this.i16(at + 6) - this.i16(at + 2);
      const width = this.i16(at + 8) - this.i16(at + 4);
      this.pictures.push({ data: this.bytes.slice(at, end), ext: "pict" });
      if (width > 0 && height > 0) runs.push({ props: { toggles: {} }, content: [{ kind: "picture", index: this.pictures.length - 1, width: width * 20, height: height * 20 }] });
    }
    return { type: "paragraph", props, mark: { toggles: {} }, runs };
  }

  /** The model's font index for a Mac font number. */
  private font(id: number): number {
    let index = this.fontIndex.get(id);
    if (index === undefined) {
      index = this.fonts.length;
      this.fonts.push(MAC_FONTS[id] ?? "Geneva");
      this.fontIndex.set(id, index);
    }
    return index;
  }

  /** How a font's character codes become text: the Symbol font's are its Unicode characters, any other font's Mac OS Roman. */
  private decoder(font: number): (code: number) => string {
    if (MAC_FONTS[font] === "Symbol") return (code) => symbolText(Uint8Array.of(code));
    return (code) => MAC_ROMAN_CHARS[code];
  }

  /** A run's properties where they differ from Normal's (the default character formatting). */
  private runProps(char: Word1Char): RunProps {
    const base = WORD1_DEFAULT_CHAR;
    const props: RunProps = { toggles: {} };
    if (char.font !== base.font) props.ftc0 = props.ftc2 = this.font(char.font);
    if (char.hps !== base.hps) props.hps = char.hps;
    if (char.bold) props.toggles.b = 1;
    if (char.italic) props.toggles.i = 1;
    if (char.caps) props.toggles.caps = 1;
    if (char.outline) props.toggles.outline = 1;
    if (char.shadow) props.toggles.shadow = 1;
    if (char.underline) props.kul = 1;
    if (char.hpsPos) props.hpsPos = char.hpsPos;
    return props;
  }

  private fullRunProps(char: Word1Char): RunProps {
    return { toggles: {}, ftc0: this.font(char.font), ftc2: this.font(char.font), hps: char.hps };
  }
}
