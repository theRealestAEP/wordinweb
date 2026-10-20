// Word for MS-DOS, versions 1 to 6 (signature 31 BE 00 00 00 AB, as
// Windows Write): a 128-byte header; the text from byte 128 to fcMac in an
// OEM code page (437 or 850); then 128-byte pages of character and
// paragraph formatting (FKPs), the footnote table, the section properties
// and their table. A paragraph property can mark the paragraph a running
// head (a header or footer); footnote texts follow the main text. Styles
// live in a separate style sheet file (.STY), which a document only names.
//
// Written from Microsoft's Write file format (which shares the header and
// the FKP pages), Microsoft's KB articles on Word for MS-DOS, and the
// structures as observed in Word's files. libwps's notes on the format were
// a guide to the property bytes.
import { Border, ParaProps, RunProps, SectionProps, Tab, newParaProps, newSectionProps, newTableProps } from "../doc/sprm.js";
import { symbolText } from "../pdf/dib.js";
import { Block, Inline, OldDocument, Paragraph, Run, Section, Table } from "./model.js";

/** Code page 437 (US) and 850 (multilingual), bytes 0x80-0xFF. */
const CP437 = "ÇüéâäàåçêëèïîìÄÅÉæÆôöòûùÿÖÜ¢£¥₧ƒáíóúñÑªº¿⌐¬½¼¡«»░▒▓│┤╡╢╖╕╣║╗╝╜╛┐└┴┬├─┼╞╟╚╔╩╦╠═╬╧╨╤╥╙╘╒╓╫╪┘┌█▄▌▐▀αßΓπΣσµτΦΘΩδ∞φε∩≡±≥≤⌠⌡÷≈°∙·√ⁿ²■\u00a0";
const CP850 = "ÇüéâäàåçêëèïîìÄÅÉæÆôöòûùÿÖÜø£Ø×ƒáíóúñÑªº¿®¬½¼¡«»░▒▓│┤ÁÂÀ©╣║╗╝¢¥┐└┴┬├─┼ãÃ╚╔╩╦╠═╬¤ðÐÊËÈıÍÎÏ┘┌█▄¦Ì▀ÓßÔÒõÕµþÞÚÛÙýÝ¯´\u00ad±‗¾¶§÷¸°¨·¹³²■\u00a0";
/**
 * What a PostScript printer printed, bytes 0x80-0xFF: authors typed the
 * codes of the printer's characters (dashes, quotes) where the screen
 * showed line characters. Word 5's driver (POSTSCRP.INI) encodes its fonts
 * as code page 437's letters, then accented capitals and punctuation; Word
 * 4's prints them in Adobe's standard encoding. Codes a driver leaves out
 * are code page 437's.
 */
const PS_WORD5 = "ÇüéâäàåçêëèïîìÄÅÉæÆôöòûùÿÖÜ¢£¥₧ƒáíóúñÑªº¿⌐¬½¼¡«»ÁÂÀÃÊËÈÍÎÏÌÓÔÒÕÚÛÙŠŸ—Ž╞õšž⁄§¤'“‹›╤╥–†‡¶‚„”…█`´ˆ˜¯ß˘˙¨¸µ˝˛ˇ—ŁØıłø≡±≥≤⌠ã÷≈°•·√ⁿ²■\u00a0";
const PS_STANDARD = "ÇüéâäàåçêëèïîìÄÅÉæÆôöòûùÿÖÜ¢£¥₧ƒá¡¢£⁄¥ƒ§¤'“«‹›ﬁﬂ░–†‡·╡¶•‚„”»…‰╛¿└`´ˆ˜¯˘˙¨╔˚¸╠˝˛ˇ—╤╥╙╘╒╓╫╪┘┌█▄▌▐▀αÆΓªΣσµτŁØŒº∞φε∩≡æ≥≤⌠ı÷≈łøœßⁿ²■\u00a0";

/** The colors of a character, by number: the PC's eight text colors. */
const COLORS = ["auto", "0000FF", "00FF00", "00FFFF", "FF0000", "FF00FF", "FFFF00", "FFFFFF"];

/**
 * Whether `head` starts a Word for MS-DOS document: Write's signature, but
 * not a Write file. Write keeps its page count at 0x60 and nothing in the
 * style sheet name (0x1E-0x5F) or at 0x62; Word keeps 0 or a printer name
 * at 0x60.
 */
export function isDosWordFile(head: Uint8Array): boolean {
  if (head.length < 128 || head[0] !== 0x31 || head[1] !== 0xbe || head[2] || head[3] || head[4] || head[5] !== 0xab) return false;
  const write = (head[0x60] || head[0x61]) && head.subarray(0x1e, 0x60).every((b) => b === 0) && !head[0x62] && !head[0x63];
  return !write;
}

/** A character's formatting, as a CHP keeps it. */
interface DosChar {
  bold: boolean;
  italic: boolean;
  /** A font number: its class (modern, roman, script, foreign, decor, symbol) and its place in the class. */
  ftc: number;
  hps: number;
  /** 0 none, 1 single, 3 double. */
  kul: number;
  strike: boolean;
  caps: boolean;
  smallCaps: boolean;
  hidden: boolean;
  /** The low characters are Word's own: the page number, the date, a note's reference. */
  special: boolean;
  /** Half-points up (positive) or down. */
  hpsPos: number;
  color: number;
}

/** A paragraph's formatting: its properties, and the running head it makes. */
interface DosPara {
  props: ParaProps;
  /** Header (0) or footer (1) and the pages: odd (2), even (4), first (8); 0 for a body paragraph. */
  rhc: number;
  /** Side by side with the paragraphs next to it that are too, each in the place its indents give it. */
  sideBySide: boolean;
}

/** A row of side-by-side paragraphs: per column its left edge, its right edge and its paragraphs. */
type SideBySideRow = { left: number; paragraphs: { p: Paragraph; right: number }[] }[];

interface FkpRun {
  fc: number;
  end: number;
  data?: Uint8Array;
}

interface Note {
  ref: number;
  start: number;
  end: number;
  /** The length of the reference's own text, for a note with a mark the author typed; 0 for a numbered one. */
  markLength: number;
}

/** A Word for MS-DOS document as an OldDocument. */
export function readDosWord(bytes: Uint8Array): OldDocument {
  return new DosReader(bytes).read();
}

class DosReader {
  private readonly fcMac: number;
  private readonly codePage: string;
  private readonly chpRuns: FkpRun[];
  private readonly papRuns: FkpRun[];
  private readonly notes: Note[] = [];
  private readonly noteAt = new Map<number, number>();
  private readonly fonts: string[] = [];
  private readonly fontIndex = new Map<number, number>();
  private evenAndOdd = false;

  constructor(private readonly bytes: Uint8Array) {
    this.fcMac = Math.min(this.u32(0x0e), bytes.length);
    // The printer's driver name (0x62), and Word's version: 0 up to Word 4, 3 and above Word 5 and later.
    const printer = String.fromCharCode(...bytes.subarray(0x62, 0x6a));
    if (/^(POSTSCR|PSCRIPT|APPL)/.test(printer)) this.codePage = bytes[0x74] >= 3 ? PS_WORD5 : PS_STANDARD;
    else this.codePage = this.u16(0x7e) === 850 ? CP850 : CP437;
    const pnChar = Math.ceil(this.fcMac / 128);
    this.chpRuns = this.readFkps(pnChar, this.u16(0x12), false);
    this.papRuns = this.readFkps(this.u16(0x12), this.u16(0x14), true);
  }

  private u16(at: number): number {
    return (this.bytes[at] ?? 0) | ((this.bytes[at + 1] ?? 0) << 8);
  }

  private u32(at: number): number {
    return (this.u16(at) | (this.u16(at + 2) << 16)) >>> 0;
  }

  read(): OldDocument {
    // Font 0, Normal's, is the model's first.
    this.font(0);
    const textEnd = this.fcMac - 128;
    this.readNotes(textEnd);
    // The notes' texts follow the main text.
    const mainEnd = this.notes.reduce((end, note) => Math.min(end, note.start), textEnd);
    const sections = this.readSections(mainEnd);
    const footnotes = this.notes.map((note) => this.blocks(note.start, note.end, true));
    return {
      fonts: this.fonts,
      styles: [{ name: "Normal", para: newParaProps(0), run: { toggles: {}, ftc0: 0, ftc2: 0, hps: 24 } }],
      sections,
      footnotes,
      pictures: [],
      evenAndOddHeaders: this.evenAndOdd,
    };
  }

  // ---- Formatting pages ---------------------------------------------------

  /**
   * The runs of the FKP pages [first, limit): per page the first FC, then
   * per run its end FC and the offset of its properties from byte 4 (0xFFFF
   * for the defaults), the run count in the last byte.
   */
  private readFkps(first: number, limit: number, paragraph: boolean): FkpRun[] {
    const runs: FkpRun[] = [];
    for (let pn = first; pn < limit && (pn + 1) * 128 <= this.bytes.length; pn++) {
      const page = pn * 128;
      const count = this.bytes[page + 127];
      let fc = this.u32(page);
      for (let i = 0; i < count && 4 + 6 * i + 6 <= 127; i++) {
        const end = this.u32(page + 4 + 6 * i);
        const offset = this.u16(page + 8 + 6 * i);
        let data: Uint8Array | undefined;
        if (offset < 122) {
          const at = page + 4 + offset;
          data = this.bytes.subarray(at + 1, Math.min(at + 1 + this.bytes[at], page + 127));
        }
        if (end > fc) runs.push({ fc, end, data });
        fc = end;
      }
    }
    // A paragraph's properties are those of the run holding its mark; a run past the text ends at fcMac.
    if (paragraph) runs.sort((a, b) => a.fc - b.fc);
    return runs;
  }

  private runAt(runs: FkpRun[], fc: number): FkpRun | undefined {
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

  private charAt(cp: number): { char: DosChar; end: number } {
    const run = this.runAt(this.chpRuns, cp + 128);
    return { char: parseChar(run?.data), end: run ? run.end - 128 : cp + 1 };
  }

  // ---- Notes and sections -------------------------------------------------

  /**
   * The footnote table: a count (with the closing entry), then per note the
   * CP of its reference and of its text. Annotations are notes too, whose
   * reference is the annotation character.
   */
  private readNotes(textEnd: number): void {
    const at = this.u16(0x14) * 128;
    if (this.u16(0x16) <= this.u16(0x14)) return;
    const count = this.u16(at);
    let last = 0;
    for (let i = 0; i + 1 < count && at + 16 + 8 * i <= this.bytes.length; i++) {
      const ref = this.u32(at + 4 + 8 * i);
      const start = this.u32(at + 8 + 8 * i);
      const end = Math.min(this.u32(at + 16 + 8 * i), textEnd);
      // The notes' texts follow the main text and each other: a table that says otherwise is damaged there.
      if (ref >= start || start < last || start >= end || this.noteAt.has(ref)) continue;
      last = end;
      this.noteAt.set(ref, this.notes.length + 1);
      this.notes.push({ ref, start, end, markLength: this.charAt(ref).char.special ? 0 : this.markLength(ref, start) });
    }
  }

  /**
   * The length of a mark the author typed: the note's text starts with it
   * too, in the same formatting as the reference. At least its first
   * character, for a note whose text starts otherwise.
   */
  private markLength(ref: number, start: number): number {
    const limit = Math.min(this.charAt(ref).end - ref, this.charAt(start).end - start, 8);
    let n = 0;
    while (n < limit && this.bytes[128 + ref + n] >= 0x20 && this.bytes[128 + ref + n] === this.bytes[128 + start + n]) n++;
    return Math.max(n, this.bytes[128 + ref] >= 0x20 ? 1 : 0);
  }

  /**
   * The sections: the section table lists per section the CP after its
   * section mark and where its properties are; the table's last entry
   * closes it. Text after the last section mark is a section of the
   * defaults.
   */
  private readSections(mainEnd: number): Section[] {
    const limits: { end: number; sep?: Uint8Array }[] = [];
    if (this.u16(0x1a) > this.u16(0x18)) {
      const at = this.u16(0x18) * 128;
      const count = this.u16(at);
      for (let i = 0; i < count; i++) {
        const end = this.u32(at + 4 + 10 * i);
        const fcSep = this.u32(at + 10 + 10 * i);
        if (end > mainEnd || end <= (limits[limits.length - 1]?.end ?? 0)) continue;
        const sep = fcSep < this.bytes.length ? this.bytes.subarray(fcSep + 1, fcSep + 1 + this.bytes[fcSep]) : undefined;
        limits.push({ end, sep });
      }
    }
    // A section the table lists ends in its section mark (a 0x0C); the text after the last has none.
    const marked = limits.length;
    if (!limits.length || limits[limits.length - 1].end < mainEnd) limits.push({ end: mainEnd });
    let start = 0;
    const sections = limits.map(({ end, sep }, i) => {
      const textEnd = i < marked && this.bytes[127 + end] === 0x0c ? end - 1 : end;
      const section = this.section(start, Math.max(start, textEnd), sep);
      start = end;
      return section;
    });
    // A section on the same page keeps the page's size and margins, which Word stores a few twips apart at times: a change starts a new page.
    const keys = ["xaPage", "yaPage", "dyaTop", "dyaBottom", "dxaLeft", "dxaRight"] as const;
    for (let i = 1; i < sections.length; i++) {
      const props = sections[i].props;
      const before = sections[i - 1].props;
      if (props.bkc !== 0 || !keys.every((key) => Math.abs((props[key] ?? 0) - (before[key] ?? 0)) <= 20)) continue;
      for (const key of keys) props[key] = before[key];
    }
    return sections;
  }

  /**
   * A section's properties (offsets after the size byte): flags, with how
   * the section starts in bits 1-3 (1 on the same page, 2 on a new one);
   * the page height and width; the first page number (0xFFFF to continue);
   * the top margin and the text's height; the left margin and the text's (a
   * column's) width; the column count; the running heads' distance from the
   * top; the column spacing.
   */
  private sectionProps(sep: Uint8Array | undefined): SectionProps {
    const d = new Uint8Array(32);
    // The defaults: a new page, a Letter page, margins of 1 inch top and bottom and 1.25 inches left and right.
    d.set([0xd4, 0x02, 0xe0, 0x3d, 0xd0, 0x2f, 0xff, 0xff, 0xa0, 0x05, 0xa0, 0x32, 0x08, 0x07, 0xc0, 0x21, 0, 1, 0xd0, 0x02, 0x10, 0x3b, 0xd0, 0x02]);
    if (sep) d.set(sep.subarray(0, 32));
    const word = (at: number): number => d[at] | (d[at + 1] << 8);
    const props = newSectionProps();
    if (((d[0] >> 1) & 7) === 1) props.bkc = 0;
    const height = word(2);
    const width = word(4);
    const top = word(8);
    const textHeight = word(10);
    const left = word(12);
    const columnWidth = word(14);
    const columns = Math.max(1, Math.min(d[17], 10));
    const spacing = word(22);
    if (height >= top + textHeight && width >= left + columnWidth && textHeight > 0 && columnWidth > 0) {
      props.yaPage = height;
      props.xaPage = width;
      props.dyaTop = top;
      props.dyaBottom = height - top - textHeight;
      props.dxaLeft = left;
      props.dxaRight = Math.max(0, width - left - columns * columnWidth - (columns - 1) * spacing);
      props.landscape = width > height;
    }
    if (word(6) !== 0xffff) {
      props.pgnRestart = true;
      props.pgnStart = word(6);
    }
    if (columns > 1) {
      props.columns = columns;
      props.dxaColumns = spacing;
    }
    const header = word(18);
    const footer = word(20);
    if (header > 0 && header < height) props.dyaHdrTop = header;
    if (footer > 0 && footer < height) props.dyaHdrBottom = height - footer;
    return props;
  }

  /** The section [start, end): its paragraphs, its running heads taken out of them. */
  private section(start: number, end: number, sep: Uint8Array | undefined): Section {
    const props = this.sectionProps(sep);
    const blocks: Block[] = [];
    // Each running head's paragraphs, by kind and page: the first group of adjacent running-head paragraphs that defines it.
    const heads = new Map<string, { group: number; blocks: Block[] }>();
    let group = 0;
    let pageBreak = false;
    const rows: SideBySideRow[] = [];
    // The width of the text (of a column): side-by-side paragraphs' right indents are from its right edge.
    const columns = props.columns ?? 1;
    const textWidth = Math.round(((props.xaPage ?? 12240) - (props.dxaLeft ?? 1800) - (props.dxaRight ?? 1800) - (columns - 1) * (props.dxaColumns ?? 720)) / columns);
    for (const { para, start: pStart, end: pEnd } of this.paragraphs(start, end)) {
      if (para.rhc & 0x0e) {
        const p = this.paragraph(pStart, pEnd, para, false);
        const kind = para.rhc & 1 ? "footer" : "header";
        for (const [bit, type] of [[2, "default"], [4, "even"], [8, "first"]] as const) {
          if (!(para.rhc & bit)) continue;
          const key = `${kind} ${type}`;
          const head = heads.get(key) ?? { group, blocks: [] };
          heads.set(key, head);
          if (head.group === group) head.blocks.push(p);
        }
        continue;
      }
      group++;
      // A paragraph that is only a page break: the next paragraph starts the page.
      if (pEnd - pStart === 1 && this.bytes[128 + pStart] === 0x0c) {
        pageBreak = true;
        continue;
      }
      const p = this.paragraph(pStart, pEnd, para, false);
      if (pageBreak) p.props.pageBreakBefore = true;
      pageBreak = false;
      if (para.sideBySide) {
        addSideBySide(rows, p, textWidth);
        continue;
      }
      for (const block of [...sideBySideBlocks(rows.splice(0)), p]) blocks.push(block);
    }
    for (const block of sideBySideBlocks(rows)) blocks.push(block);
    const headers: Section["headers"] = [];
    for (const kind of ["header", "footer"] as const) {
      const odd = heads.get(`${kind} default`)?.blocks ?? [];
      const even = heads.get(`${kind} even`)?.blocks ?? [];
      const first = heads.get(`${kind} first`)?.blocks ?? [];
      if (!odd.length && !even.length && !first.length) continue;
      // The even pages' too: a .docx with different even pages (a document-wide setting) shows only it on them.
      headers.push({ kind, type: "default", blocks: odd }, { kind, type: "even", blocks: even });
      if (!sameBlocks(even, odd)) this.evenAndOdd = true;
      if (!sameBlocks(first, odd)) props.titlePage = true;
    }
    if (props.titlePage) {
      for (const kind of ["header", "footer"] as const) headers.push({ kind, type: "first", blocks: heads.get(`${kind} first`)?.blocks ?? [] });
    }
    return { props, blocks, headers };
  }

  // ---- Paragraphs -----------------------------------------------------------

  /** The paragraphs that end in [start, end), each with its properties. */
  private *paragraphs(start: number, end: number): Generator<{ para: DosPara; start: number; end: number }> {
    let cp = start;
    while (cp < end) {
      const run = this.runAt(this.papRuns, cp + 128);
      const runEnd = Math.max(Math.min(run ? run.end - 128 : end, end), cp + 1);
      const para = parsePara(run?.data);
      // A run of the same properties holds paragraphs that end in a CR LF each.
      while (cp < runEnd) {
        let pEnd = cp;
        while (pEnd < runEnd && !(this.bytes[127 + pEnd] === 0x0d && this.bytes[128 + pEnd] === 0x0a)) pEnd++;
        pEnd = Math.min(pEnd + 1, runEnd);
        yield { para: { ...para, props: structuredClone(para.props) }, start: cp, end: pEnd };
        cp = pEnd;
      }
    }
  }

  /** The blocks of a note's text. */
  private blocks(start: number, end: number, note: boolean): Block[] {
    const out: Block[] = [];
    for (const { para, start: pStart, end: pEnd } of this.paragraphs(start, end)) {
      out.push(this.paragraph(pStart, pEnd, para, note && pStart === start));
    }
    return out;
  }

  /** The paragraph of the text [start, end), its mark (CR LF) left out; `noteStart` for a note's first, whose reference mark starts it. */
  private paragraph(start: number, end: number, para: DosPara, noteStart: boolean): Paragraph {
    let textEnd = end;
    if (this.bytes[127 + textEnd] === 0x0a) textEnd--;
    if (textEnd > start && this.bytes[127 + textEnd] === 0x0d) textEnd--;
    const runs: Run[] = [];
    let cp = start;
    if (noteStart && this.charAt(cp).char.special && this.bytes[128 + cp] < 0x20) {
      // The note's number; a mark the author typed is text.
      runs.push({ props: this.runProps(this.charAt(cp).char), content: [{ kind: "footnoteMark" }] });
      cp++;
    }
    while (cp < textEnd) {
      const { char, end: runEnd } = this.charAt(cp);
      const stop = Math.min(Math.max(runEnd, cp + 1), textEnd);
      const run = this.run(cp, stop, char);
      if (run.content.length) runs.push(run);
      cp = stop;
    }
    const mark = this.charAt(Math.max(start, end - 1)).char;
    return { type: "paragraph", props: para.props, mark: this.runProps(mark), runs };
  }

  private run(start: number, end: number, char: DosChar): Run {
    const content: Inline[] = [];
    let text = "";
    const flush = (): void => {
      if (text) content.push({ kind: "text", text });
      text = "";
    };
    const symbol = this.fonts[this.font(char.ftc)] === "Symbol";
    for (let cp = start; cp < end; cp++) {
      // A note's reference: a mark the author typed is the note's own text.
      const note = this.noteAt.get(cp);
      if (note) {
        flush();
        const { markLength } = this.notes[note - 1];
        const mark = markLength ? this.decode(cp, Math.min(end, cp + markLength)) : undefined;
        content.push({ kind: "footnote", id: note, customMark: mark });
        if (markLength) cp = Math.min(end, cp + markLength) - 1;
        continue;
      }
      const ch = this.bytes[128 + cp];
      if (char.special && ch < 0x20) {
        flush();
        if (ch === 0x01 || ch === 0x09) content.push({ kind: "field", code: "PAGE", result: "1" });
        else if (ch === 0x02) content.push({ kind: "field", code: "DATE", result: "" });
        else if (ch === 0x03) content.push({ kind: "field", code: "TIME", result: "" });
        continue;
      }
      if (ch === 0xc4 && !symbol && this.isNoBreakHyphen(cp)) {
        flush();
        content.push({ kind: "noBreakHyphen" });
        continue;
      }
      if (ch >= 0x20) {
        text += symbol ? symbolText(Uint8Array.of(ch)) : this.char(ch);
        continue;
      }
      flush();
      switch (ch) {
        case 0x09: content.push({ kind: "tab" }); break;
        case 0x0b: content.push({ kind: "break" }); break;
        // A line feed with no carriage return before it, from text Word took in: a new line.
        case 0x0a: if (this.bytes[127 + cp] !== 0x0d) content.push({ kind: "break" }); break;
        case 0x0c: content.push({ kind: "page" }); break;
        case 0x0e: content.push({ kind: "column" }); break;
        case 0x1f: content.push({ kind: "softHyphen" }); break;
        default: break;
      }
    }
    flush();
    return { props: this.runProps(char), content };
  }

  /**
   * Whether the 0xC4 at `cp` is Word's nonbreaking hyphen: it is the code
   * page's horizontal line too, which draws boxes and rules. A hyphen is
   * next to a letter or digit and to no other line character.
   */
  private isNoBreakHyphen(cp: number): boolean {
    const before = this.bytes[127 + cp];
    const after = this.bytes[129 + cp];
    const line = (b: number): boolean => b >= 0xb0 && b <= 0xdf;
    const word = (b: number): boolean => (b >= 0x30 && b <= 0x39) || (b >= 0x41 && b <= 0x5a) || (b >= 0x61 && b <= 0x7a) || (b >= 0x80 && b <= 0xa5);
    return !line(before) && !line(after) && (word(before) || word(after));
  }

  /** A text character (0x20 and above) in the document's code page. */
  private char(ch: number): string {
    return ch < 0x80 ? String.fromCharCode(ch) : this.codePage[ch - 0x80];
  }

  private decode(start: number, end: number): string {
    let s = "";
    for (let cp = start; cp < end; cp++) {
      const ch = this.bytes[128 + cp];
      if (ch >= 0x20) s += this.char(ch);
    }
    return s;
  }

  /** A run's properties over the Normal style's (Courier, 12 points). */
  private runProps(c: DosChar): RunProps {
    const props: RunProps = { toggles: {} };
    if (c.ftc !== 0) props.ftc0 = props.ftc2 = this.font(c.ftc);
    if (c.hps !== 24) props.hps = c.hps;
    if (c.bold) props.toggles.b = 1;
    if (c.italic) props.toggles.i = 1;
    if (c.strike) props.toggles.strike = 1;
    if (c.caps) props.toggles.caps = 1;
    if (c.smallCaps) props.toggles.smallCaps = 1;
    if (c.hidden) props.toggles.vanish = 1;
    if (c.kul) props.kul = c.kul;
    // Word for MS-DOS offers superscript and subscript only: the raise is its fixed amount for them.
    if (c.hpsPos) props.iss = c.hpsPos > 0 ? 1 : 2;
    if (c.color) props.color = COLORS[c.color];
    return props;
  }

  /** The model's font index for a font number. */
  private font(ftc: number): number {
    let index = this.fontIndex.get(ftc);
    if (index === undefined) {
      index = this.fonts.length;
      this.fonts.push(fontName(ftc));
      this.fontIndex.set(ftc, index);
    }
    return index;
  }
}

/**
 * The font a font number stands for. The printer driver names the fonts of
 * each class; these are the usual ones: Courier and Helvetica among the
 * modern fonts (0-15), Times among the roman (16-31) and the script,
 * foreign and decorative ones, Symbol for the symbol class (56-63).
 */
function fontName(ftc: number): string {
  if (ftc < 8) return "Courier New";
  if (ftc < 16) return "Arial";
  if (ftc >= 56) return "Symbol";
  return "Times New Roman";
}

/**
 * A CHP: byte 0 the character style (bit 0 set, the number above it);
 * byte 1 bold (bit 0), italic (bit 1) and the font number; byte 2 the size
 * in half-points; byte 3 underline (bit 0), strikethrough (1), double
 * underline (2), capitals (bits 4-5: 1 all caps, 3 small caps), Word's own
 * characters (6), hidden (7); byte 5 the raise in half-points (signed);
 * byte 6 the color. Bytes the CHP leaves out are 0, the size 12 points.
 */
function parseChar(data: Uint8Array | undefined): DosChar {
  const d = new Uint8Array(7);
  d[2] = 24;
  if (data) d.set(data.subarray(0, 7));
  const caps = (d[3] >> 4) & 3;
  return {
    bold: (d[1] & 1) !== 0,
    italic: (d[1] & 2) !== 0,
    ftc: d[1] >> 2,
    hps: d[2] || 24,
    kul: d[3] & 4 ? 3 : d[3] & 1 ? 1 : 0,
    strike: (d[3] & 2) !== 0,
    caps: caps === 1,
    smallCaps: caps === 3,
    hidden: (d[3] & 0x80) !== 0,
    special: (d[3] & 0x40) !== 0,
    hpsPos: (d[5] << 24) >> 24,
    color: d[6] & 7,
  };
}

/** The border lines of PAP byte 16's top bits: normal, bold, double, thick. */
const BORDER_LINES: Border[] = [
  { type: 1, width: 4, color: "auto", space: 1, shadow: false, frame: false },
  { type: 1, width: 12, color: "auto", space: 1, shadow: false, frame: false },
  { type: 3, width: 4, color: "auto", space: 1, shadow: false, frame: false },
  { type: 1, width: 18, color: "auto", space: 1, shadow: false, frame: false },
];

/**
 * A PAP: byte 0 the paragraph style (bit 0 set, the number above it), which
 * only the style sheet defines; byte 1 the alignment (bits 0-1) and keep together (bit 2), keep with next (bit
 * 3); then signed words from byte 4: the right indent, the left indent, the
 * first line's indent, the line spacing (not positive: as the font needs),
 * the space before and after, all in twips; byte 16 the running head (bits
 * 0-3) and the border (bits 4-5: 1 a box, 2 the sides byte 17 lists; bits
 * 6-7 the line); byte 17 the sides (bit 0 left, 1 right, 2 top, 3 bottom);
 * byte 18 the shading in percent; from byte 22 the tabs, four bytes each:
 * the position, then the alignment (bits 0-2) and leader (bits 3-5).
 */
function parsePara(data: Uint8Array | undefined): DosPara {
  const d = new Uint8Array(22 + 4 * 20);
  d[10] = 240;
  if (data) d.set(data.subarray(0, d.length));
  const word = (at: number): number => ((d[at] | (d[at + 1] << 8)) << 16) >> 16;
  const p = newParaProps(0);
  if (d[1] & 3) p.jc = d[1] & 3;
  if (d[1] & 4) p.keepLines = true;
  if (d[1] & 8) p.keepNext = true;
  if (word(4)) p.dxaRight = word(4);
  if (word(6)) p.dxaLeft = word(6);
  if (word(8)) p.dxaLeft1 = word(8);
  const line = word(10);
  if (line > 0 && line !== 240) {
    p.lineSpacing = line;
    p.lineMultiple = false;
  }
  if (word(12)) p.dyaBefore = word(12);
  if (word(14)) p.dyaAfter = word(14);
  const border = (d[16] >> 4) & 3;
  if (border) {
    const line = BORDER_LINES[d[16] >> 6];
    const sides = border === 1 ? 0x0f : d[17] & 0x0f;
    if (sides & 1) p.borders.left = line;
    if (sides & 2) p.borders.right = line;
    if (sides & 4) p.borders.top = line;
    if (sides & 8) p.borders.bottom = line;
  }
  const shade = d[18] & 0x7f;
  if (shade) p.shading = { pattern: 0, fore: "auto", fill: grey(Math.min(shade, 100)) };
  const tabs: Tab[] = [];
  const tabBytes = data ? Math.min(data.length, d.length) : 0;
  for (let at = 22; at + 2 <= tabBytes; at += 4) {
    const pos = word(at);
    if (pos <= 0) break;
    tabs.push({ pos, jc: Math.min(d[at + 2] & 7, 3), leader: Math.min((d[at + 2] >> 3) & 7, 3) });
  }
  p.tabs = tabs;
  return { props: p, rhc: d[16] & 0x0f, sideBySide: (d[1] & 0x10) !== 0 };
}

/** A shade of grey: `percent` of black. */
function grey(percent: number): string {
  const v = Math.round(255 * (1 - percent / 100));
  const h = v.toString(16).toUpperCase().padStart(2, "0");
  return h + h + h;
}

/**
 * Adds a side-by-side paragraph to the rows: a paragraph with the left
 * indent of the column before it joins that column, one further right
 * starts the next column, one further left starts a new row.
 */
function addSideBySide(rows: SideBySideRow[], p: Paragraph, textWidth: number): void {
  const left = p.props.dxaLeft ?? 0;
  const right = Math.max(left + 1, textWidth - (p.props.dxaRight ?? 0));
  const row = rows[rows.length - 1];
  const column = row?.[row.length - 1];
  if (column && left === column.left) column.paragraphs.push({ p, right });
  else if (column && left > column.left) row.push({ left, paragraphs: [{ p, right }] });
  else rows.push([{ left, paragraphs: [{ p, right }] }]);
}

/**
 * Side-by-side paragraphs as a table with no borders: a column a cell,
 * which reaches to the next column's left edge (the last, to its widest
 * paragraph's right edge); each paragraph keeps the space between its own
 * right edge and the cell's as its right indent. Rows of one column are
 * the paragraphs as they are.
 */
function sideBySideBlocks(rows: SideBySideRow[]): Block[] {
  if (rows.every((row) => row.length === 1)) return rows.flatMap((row) => row[0].paragraphs.map(({ p }) => p));
  const table: Table = { type: "table", rows: [] };
  for (const row of rows) {
    const props = newTableProps();
    props.left = row[0].left;
    props.cellMargins = { left: 0, right: 0 };
    const cells: Block[][] = [];
    row.forEach((column, i) => {
      const edge = i + 1 < row.length ? row[i + 1].left : column.paragraphs.reduce((edge, { right }) => Math.max(edge, right), 0);
      props.cells.push({
        right: edge, horzMerge: 0, vertMerge: 0, vertAlign: 0, textFlow: 0, ftsWidth: 0, width: 0,
        fitText: false, noWrap: false, hideMark: false, borders: {}, margins: {},
      });
      for (const { p, right } of column.paragraphs) {
        p.props.dxaLeft = undefined;
        p.props.dxaRight = edge > right ? edge - right : undefined;
      }
      cells.push(column.paragraphs.map(({ p }) => p));
    });
    table.rows.push({ props, cells });
  }
  return [table];
}

function sameBlocks(a: Block[], b: Block[]): boolean {
  return a.length === b.length && a.every((block, i) => block === b[i]);
}
