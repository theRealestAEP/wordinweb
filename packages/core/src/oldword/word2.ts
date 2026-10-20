// What Word for the Macintosh 4.0/5.x and Word for Windows 1.x/2.0 share:
// one file format, big-endian on the Mac and little-endian on Windows. A
// header (the FIB) of file positions; the text in pieces when the file was
// fast-saved; character and paragraph formatting in 512-byte FKP pages; a
// style sheet of names, character and paragraph properties and a based-on
// list; sections; footnotes, headers and footers as stories after the main
// text; paragraph, section and table properties as one-byte sprms (Word for
// Windows 2.0's numbering). mac.ts and winword.ts read what differs: the
// FIB's layout, the character properties, the fonts, the code page and the
// pictures.
import { Border, CellProps, ParaProps, RunProps, SectionProps, TableProps, newParaProps, newTableProps } from "../doc/sprm.js";
import { Block, Inline, OldDocument, Paragraph, Picture, Row, Run, Section, Style } from "./model.js";

/** A character's formatting: a font number (the file's own), and the rest. */
export interface Char {
  fontId: number;
  hps: number;
  /** bold, italic, strike, outline, shadow, small caps, all caps, hidden. */
  flags: boolean[];
  /** A Kul: 0 none, 1 single, 2 words, 3 double, 4 dotted. */
  kul: number;
  hpsPos: number;
  /** Twips. */
  dxaSpace: number;
  /** A Word color number (0 auto, 1 black ... 8 white). */
  ico: number;
  /** The file offset of the picture a U+0001 shows. */
  picture?: number;
  /** Whether the run's control characters are special ones (a picture, a field, a note reference). */
  special: boolean;
}

/** A paragraph's properties: ParaProps for the XML writer, with the character changes its sprms make. */
export interface Para {
  props: ParaProps;
  /** The character sprms among the paragraph's, applied to its text. */
  charSprms: Sprm[];
  /** Old-style border (sprm 0x0a): the line every side sprm 0x0b turns on gets. */
  borderStyle?: Border;
}

export interface Sprm {
  op: number;
  arg: Uint8Array;
}

export interface Entry {
  fc: number;
  cb: number;
}

/** What the FIB gives: the text's start and lengths, and where each table of the file is. */
export interface Fib {
  fcMin: number;
  ccpText: number;
  ccpFtn: number;
  ccpHdd: number;
  stsh: Entry;
  fndRef: Entry;
  fndTxt: Entry;
  sed: Entry;
  hdd: Entry;
  bteChpx: Entry;
  btePapx: Entry;
  clx: Entry;
}

/** The document-wide properties: the page (Word for the Macintosh's), the default tab and the flags. */
export interface DocProps {
  page: SectionProps;
  defaultTab?: number;
  facingPages: boolean;
  /** Footnotes at the end of the document. */
  endnotes: boolean;
  /** How many stories of note separators the header stories begin with. */
  separatorStories: number;
}

interface Piece {
  cp: number;
  end: number;
  fc: number;
  /** The piece's property modifier: sprms applied to its paragraphs and text. */
  sprms: Sprm[];
}

interface FkpRun {
  fc: number;
  end: number;
  /** The CHPX or PAPX bytes, after their size byte; undefined for the default. */
  data?: Uint8Array;
}

interface StyleEntry {
  name: string;
  para: Para;
  char: Char;
  next: number;
  index: number;
}

/** The standard styles by their number (stc - 256), as Word for Windows 2.0 numbers them. */
export const STANDARD_STYLES: Record<number, string> = {
  [-1]: "Normal Indent", [-2]: "heading 1", [-3]: "heading 2", [-4]: "heading 3", [-5]: "heading 4", [-6]: "heading 5", [-7]: "heading 6",
  [-8]: "heading 7", [-9]: "heading 8", [-10]: "heading 9", [-11]: "footnote text", [-12]: "footnote reference", [-13]: "header",
  [-14]: "footer", [-15]: "index heading", [-16]: "line number", [-17]: "index 1", [-18]: "index 2", [-19]: "index 3", [-20]: "index 4",
  [-21]: "index 5", [-22]: "index 6", [-23]: "index 7", [-24]: "toc 1", [-25]: "toc 2", [-26]: "toc 3", [-27]: "toc 4", [-28]: "toc 5",
  [-29]: "toc 6", [-30]: "toc 7", [-31]: "toc 8", [-32]: "annotation text", [-33]: "annotation reference",
};

/** The based-on number that means none. */
const NO_STYLE = -34;

const FLAG_NAMES = ["b", "i", "strike", "outline", "shadow", "smallCaps", "caps", "vanish"] as const;

/** Word's colors by number, 1 black to 16 light gray (doc/sprm.ts's Ico). */
const ICO: string[] = ["auto", "000000", "0000FF", "00FFFF", "00FF00", "FF00FF", "FF0000", "FFFF00", "FFFFFF", "000080", "008080", "008000", "800080", "800000", "808000", "808080", "C0C0C0"];

export abstract class Word2Reader {
  private pieces: Piece[] = [];
  private chpRuns: FkpRun[] = [];
  private papRuns: FkpRun[] = [];
  private text = new Uint8Array(0);
  private readonly fonts: string[] = [];
  private readonly fontIndex = new Map<number, number>();
  /** The styles by their Word number (negative for the standard ones). */
  private readonly styles = new Map<number, StyleEntry>();
  private readonly modelStyles: Style[] = [];
  /** The resolved character formatting of each model style. */
  private readonly styleChars: Char[] = [];
  protected readonly pictures: Picture[] = [];
  private readonly noteRefs = new Map<number, number>();
  private readonly footnotes: Block[][] = [];
  private fib!: Fib;
  private docProps!: DocProps;

  constructor(protected readonly bytes: Uint8Array, private readonly bigEndian: boolean) {}

  // ---- What each format reads its own way ----------------------------------

  /** The size of an FKP page: 512 bytes, 128 in Word for the Macintosh 3. */
  protected readonly fkpPageSize: number = 512;
  /** The bytes of a PAPX before its sprms: the style number and six more (three in all in Word for the Macintosh 3). */
  protected readonly papxHeader: number = 7;
  /** A SEPX's sprms, in word2.ts's numbering of the section sprms. */
  protected sectionSprms(sepx: Uint8Array): Sprm[] {
    return this.parseSprms(sepx);
  }
  /** The name of a standard style by its number. */
  protected standardStyleName(id: number): string | undefined {
    return STANDARD_STYLES[id];
  }

  protected abstract readFib(): Fib;
  protected abstract readDocProps(): DocProps;
  /** The operand sizes of the format's sprms: -1 for a size byte first, -2 for a size word that counts one byte more than follows. */
  protected abstract readonly sprmSizes: Record<number, number>;
  /** The character formatting a style based on none starts from. */
  protected abstract baseChar(): Char;
  /** Apply a CHPX (the bytes after its size byte). */
  protected abstract applyChpx(c: Char, chpx: Uint8Array): void;
  /** Apply a character sprm of a paragraph's or a piece's; `style` is the paragraph style's formatting. */
  protected abstract applyCharSprm(c: Char, sprm: Sprm, style: Char): void;
  /** Whether a paragraph sprm is a character one (kept for the paragraph's text). */
  protected abstract isCharSprm(op: number): boolean;
  /** The font name of a font number; called once per font. */
  protected abstract fontName(fontId: number): string;
  /** How the font's character codes (0x20 and up) become text. */
  protected abstract decoder(fontId: number): (code: number) => string;
  /** A tab descriptor byte: its alignment and leader. */
  protected abstract tab(type: number): { jc: number; leader: number };
  /** A piece's property modifier as sprms. */
  protected abstract prmSprms(prm: number, grpprls: Uint8Array[]): Sprm[];
  /** A style name in the file's code page. */
  protected abstract styleName(bytes: Uint8Array): string;
  /** The picture the run's U+0001 shows. */
  protected abstract picture(char: Char): Inline | undefined;

  // ---- Bytes ----------------------------------------------------------------

  protected u8(at: number): number {
    return this.bytes[at] ?? 0;
  }
  protected u16(at: number): number {
    return this.n16(this.bytes, at);
  }
  protected u32(at: number): number {
    return this.n32(this.bytes, at);
  }
  /** A 16-bit number of `b` in the file's byte order. */
  protected n16(b: Uint8Array, at: number): number {
    const x = b[at] ?? 0;
    const y = b[at + 1] ?? 0;
    return this.bigEndian ? (x << 8) | y : x | (y << 8);
  }
  protected n16s(b: Uint8Array, at: number): number {
    return (this.n16(b, at) << 16) >> 16;
  }
  protected n32(b: Uint8Array, at: number): number {
    const hi = this.n16(b, this.bigEndian ? at : at + 2);
    const lo = this.n16(b, this.bigEndian ? at + 2 : at);
    return ((hi << 16) | lo) >>> 0;
  }
  protected entry({ fc, cb }: Entry): Uint8Array {
    return cb > 0 && fc + cb <= this.bytes.length ? this.bytes.subarray(fc, fc + cb) : new Uint8Array(0);
  }

  // ---- The document -------------------------------------------------------

  read(): OldDocument {
    this.fib = this.readFib();
    this.docProps = this.readDocProps();
    this.readPieces();
    this.chpRuns = this.readFkps(this.fib.bteChpx, false);
    this.papRuns = this.readFkps(this.fib.btePapx, true);
    this.readStyles();
    this.readNoteRefs();
    const sections = this.readSections();
    this.readFootnotes();
    return {
      fonts: this.fonts,
      styles: this.modelStyles,
      sections,
      footnotes: this.footnotes,
      pictures: this.pictures,
      defaultTab: this.docProps.defaultTab,
      evenAndOddHeaders: this.docProps.facingPages,
      endnotes: this.docProps.endnotes,
    };
  }

  /** The text's pieces: the clx of a fast-saved file, else one piece from fcMin. */
  private readPieces(): void {
    const { ccpText, ccpFtn, ccpHdd, fcMin } = this.fib;
    const total = ccpText + ccpFtn + ccpHdd + (ccpFtn || ccpHdd ? 1 : 0);
    const clx = this.entry(this.fib.clx);
    const grpprls: Uint8Array[] = [];
    let at = 0;
    while (at < clx.length && clx[at] === 1) {
      const cb = this.n16(clx, at + 1);
      grpprls.push(clx.subarray(at + 3, at + 3 + cb));
      at += 3 + cb;
    }
    if (at < clx.length && clx[at] === 2) {
      const cb = this.n16(clx, at + 1);
      const plc = clx.subarray(at + 3, at + 3 + cb);
      const n = Math.floor((plc.length - 4) / 12);
      for (let i = 0; i < n; i++) {
        const cp = this.n32(plc, 4 * i);
        const end = this.n32(plc, 4 * (i + 1));
        const pcd = 4 * (n + 1) + 8 * i;
        if (end > cp) this.pieces.push({ cp, end, fc: this.n32(plc, pcd + 2), sprms: this.prmSprms(this.n16(plc, pcd + 6), grpprls) });
      }
    }
    if (this.pieces.length === 0) this.pieces.push({ cp: 0, end: total, fc: fcMin, sprms: [] });
    const length = this.pieces[this.pieces.length - 1].end;
    this.text = new Uint8Array(length);
    for (const piece of this.pieces) {
      const from = Math.min(piece.fc, this.bytes.length);
      const to = Math.min(piece.fc + piece.end - piece.cp, this.bytes.length);
      if (piece.cp < length) this.text.set(this.bytes.subarray(from, Math.min(to, from + length - piece.cp)), piece.cp);
    }
  }

  private pieceAt(cp: number): Piece | undefined {
    let lo = 0;
    let hi = this.pieces.length - 1;
    while (lo <= hi) {
      const mid = (lo + hi) >> 1;
      const p = this.pieces[mid];
      if (cp < p.cp) hi = mid - 1;
      else if (cp >= p.end) lo = mid + 1;
      else return p;
    }
    return undefined;
  }

  private fcOf(cp: number): number {
    const piece = this.pieceAt(cp);
    return piece ? piece.fc + cp - piece.cp : cp + this.fib.fcMin;
  }

  /**
   * The runs of the FKP pages a bin table lists. A page that ends before the
   * next listed one begins is followed by the pages after it in the file,
   * which the table leaves out.
   */
  private readFkps(table: Entry, paragraph: boolean): FkpRun[] {
    const bte = this.entry(table);
    const n = Math.floor((bte.length - 4) / 6);
    const runs: FkpRun[] = [];
    for (let i = 0; i < n; i++) {
      const limit = this.n32(bte, 4 * (i + 1));
      let pn = this.n16(bte, 4 * (n + 1) + 2 * i);
      let expected = this.n32(bte, 4 * i);
      for (let pages = 0; pages < 256; pages++, pn++) {
        const size = this.fkpPageSize;
        const page = pn * size;
        if (page + size > this.bytes.length) break;
        const crun = this.u8(page + size - 1);
        if (crun === 0 || 5 * crun + 4 > size - 1 || this.u32(page) !== expected) break;
        expected = this.u32(page + 4 * crun);
        // In a 512-byte page, run offsets count words, and so does a PAPX's size; in a 128-byte page (Word for the Macintosh 3), bytes.
        const unit = size === 512 ? 2 : 1;
        for (let k = 0; k < crun; k++) {
          const fc = this.u32(page + 4 * k);
          const end = this.u32(page + 4 * (k + 1));
          const offset = this.u8(page + 4 * (crun + 1) + k) * unit;
          let data: Uint8Array | undefined;
          if (offset) {
            const count = this.u8(page + offset);
            // A CHPX's size counts bytes.
            const length = paragraph ? unit * count : count;
            data = this.bytes.subarray(page + offset + 1, Math.min(page + offset + 1 + length, page + size - 1));
          }
          if (end > fc) runs.push({ fc, end, data });
        }
        if (expected >= limit) break;
      }
    }
    runs.sort((a, b) => a.fc - b.fc);
    return runs;
  }

  private runAt(runs: FkpRun[], fc: number): FkpRun | undefined {
    let lo = 0;
    let hi = runs.length - 1;
    while (lo <= hi) {
      const mid = (lo + hi) >> 1;
      const r = runs[mid];
      if (fc < r.fc) hi = mid - 1;
      else if (fc >= r.end) lo = mid + 1;
      else return r;
    }
    return undefined;
  }

  // ---- Fonts and styles ---------------------------------------------------

  /** The model's font index for a font number. */
  private font(id: number): number {
    let index = this.fontIndex.get(id);
    if (index === undefined) {
      index = this.fonts.length;
      this.fonts.push(this.fontName(id));
      this.fontIndex.set(id, index);
    }
    return index;
  }

  /**
   * The style sheet: the count of standard styles; the names, the
   * character and the paragraph properties, each a list with an entry per
   * style; and per style the next style and the one it is based on. Style
   * number i - (count of standard styles) is entry i.
   */
  private readStyles(): void {
    const stsh = this.entry(this.fib.stsh);
    const defined: { id: number; name?: string; chpx?: Uint8Array; papx?: Uint8Array; next: number; basedOn: number }[] = [];
    if (stsh.length >= 8) {
      const standard = this.n16s(stsh, 0);
      const zones: Uint8Array[] = [];
      let at = 2;
      for (let z = 0; z < 3 && at + 2 <= stsh.length; z++) {
        const size = this.n16(stsh, at);
        zones.push(stsh.subarray(at + 2, Math.min(at + size, stsh.length)));
        at += Math.max(size, 2);
      }
      const names = entriesOf(zones[0] ?? new Uint8Array(0));
      const chpxs = entriesOf(zones[1] ?? new Uint8Array(0));
      const papxs = entriesOf(zones[2] ?? new Uint8Array(0));
      const count = at + 2 <= stsh.length ? this.n16(stsh, at) : 0;
      for (let i = 0; i < Math.max(names.length, count); i++) {
        const name = names[i];
        if (name === null) continue;
        const listed = at + 3 + 2 * i < stsh.length;
        defined.push({
          id: i - standard,
          name: name && name.length ? this.styleName(name) : undefined,
          chpx: chpxs[i] ?? undefined,
          papx: papxs[i] ?? undefined,
          next: listed ? (stsh[at + 2 + 2 * i] << 24) >> 24 : i - standard,
          basedOn: listed ? (stsh[at + 3 + 2 * i] << 24) >> 24 : NO_STYLE,
        });
      }
    }
    if (!defined.some((s) => s.id === 0)) defined.push({ id: 0, next: 0, basedOn: NO_STYLE });
    const byId = new Map(defined.map((s) => [s.id, s]));
    const resolving = new Set<number>();
    const resolve = (id: number): StyleEntry => {
      const done = this.styles.get(id);
      if (done) return done;
      const style = byId.get(id);
      let para: Para = { props: newParaProps(0), charSprms: [] };
      let char = this.baseChar();
      resolving.add(id);
      if (style && style.basedOn !== NO_STYLE && style.basedOn !== id && byId.has(style.basedOn) && !resolving.has(style.basedOn)) {
        const parent = resolve(style.basedOn);
        para = clonePara(parent.para);
        char = cloneChar(parent.char);
      }
      resolving.delete(id);
      if (style?.chpx) this.applyChpx(char, style.chpx);
      // A style's paragraph entry: its own number, six bytes, then its sprms.
      if (style?.papx && style.papx.length > this.papxHeader) {
        const sprms = this.parseSprms(style.papx.subarray(this.papxHeader));
        this.applyParaSprms(para, sprms, false);
        for (const sprm of para.charSprms) this.applyCharSprm(char, sprm, char);
      }
      para.props.istd = 0;
      para.charSprms = [];
      const name = id === 0 ? "Normal" : style?.name ?? this.standardStyleName(id) ?? `Style ${id}`;
      const entry = { name, para, char, next: style?.next ?? id, index: -1 };
      this.styles.set(id, entry);
      return entry;
    };
    // Normal first: the model's style 0.
    for (const id of [0, ...defined.map((s) => s.id).filter((id) => id !== 0)]) {
      const style = resolve(id);
      style.index = this.modelStyles.length;
      this.modelStyles.push({ name: style.name, para: { ...style.para.props, istd: style.index }, run: this.fullRunProps(style.char) });
      this.styleChars.push(style.char);
    }
    for (const style of this.styles.values()) {
      const next = this.styles.get(style.next);
      if (next) this.modelStyles[style.index].next = next.index;
    }
  }

  private style(id: number): StyleEntry {
    return this.styles.get(id) ?? this.styles.get(0)!;
  }

  // ---- Stories ------------------------------------------------------------

  /** The footnote references in the main text: CP to note number. */
  private readNoteRefs(): void {
    const refs = this.entry(this.fib.fndRef);
    const n = Math.floor((refs.length - 4) / 6);
    for (let i = 0; i < n; i++) this.noteRefs.set(this.n32(refs, 4 * i), i + 1);
  }

  private readFootnotes(): void {
    const txt = this.entry(this.fib.fndTxt);
    const cps: number[] = [];
    for (let at = 0; at + 4 <= txt.length; at += 4) cps.push(this.n32(txt, at));
    const base = this.fib.ccpText;
    for (let i = 0; i < this.noteRefs.size && i + 1 < cps.length; i++) {
      const start = base + cps[i];
      const end = Math.min(base + cps[i + 1], base + this.fib.ccpFtn);
      this.footnotes.push(end > start ? this.blocks(start, end, null, true) : []);
    }
  }

  /** The sections of the main text, with their headers and footers. */
  private readSections(): Section[] {
    const mainEnd = this.fib.ccpText;
    const sed = this.entry(this.fib.sed);
    const n = Math.floor((sed.length - 4) / 10);
    const limits: { cp: number; end: number; sepx?: Uint8Array }[] = [];
    for (let i = 0; i < n; i++) {
      const cp = this.n32(sed, 4 * i);
      const end = Math.min(this.n32(sed, 4 * (i + 1)), mainEnd);
      const fcSepx = this.n32(sed, 4 * (n + 1) + 6 * i + 2);
      let sepx: Uint8Array | undefined;
      if (fcSepx !== 0xffffffff && fcSepx < this.bytes.length) sepx = this.bytes.subarray(fcSepx + 1, fcSepx + 1 + this.u8(fcSepx));
      if (end > cp) limits.push({ cp, end, sepx });
    }
    if (limits.length === 0) limits.push({ cp: 0, end: mainEnd });
    limits[limits.length - 1].end = mainEnd;
    const sectionEnds = new Set(limits.slice(0, -1).map((l) => l.end - 1));
    const hdd = this.entry(this.fib.hdd);
    const stories: number[] = [];
    for (let at = 0; at + 4 <= hdd.length; at += 4) stories.push(this.n32(hdd, at));
    const hddBase = this.fib.ccpText + this.fib.ccpFtn;
    // The stories of the note separators come first.
    let story = this.docProps.separatorStories;
    return limits.map((limit) => {
      const props: SectionProps = { ...structuredClone(this.docProps.page) };
      let headerFlags = 0;
      // A fast-saved file changes a section by the property modifier of the piece that holds its end.
      const sprms = [...this.sectionSprms(limit.sepx ?? new Uint8Array(0)), ...(this.pieceAt(limit.end - 1)?.sprms ?? [])];
      for (const { op, arg } of sprms) {
        switch (op) {
          case 0x75: props.bkc = arg[0]; break;
          case 0x76: props.titlePage = arg[0] !== 0; break;
          case 0x77: props.columns = this.n16(arg, 0) + 1; break;
          case 0x78: props.dxaColumns = this.n16(arg, 0); break;
          case 0x7a: props.nfcPgn = arg[0]; break;
          case 0x7d: props.pgnRestart = arg[0] !== 0; break;
          case 0x7f: props.lnc = arg[0]; break;
          case 0x80: headerFlags = arg[0]; break;
          case 0x81: props.nLnnMod = this.n16(arg, 0); break;
          case 0x82: props.dxaLnn = this.n16(arg, 0); break;
          case 0x87: props.lnnMin = this.n16(arg, 0); break;
          case 0x83: props.dyaHdrTop = this.n16(arg, 0); break;
          case 0x84: props.dyaHdrBottom = this.n16(arg, 0); break;
          case 0x85: props.lineBetween = arg[0] !== 0; break;
          case 0x86: props.vjc = arg[0]; break;
          case 0x88: props.pgnStart = this.n16(arg, 0); break;
          case 0x89: props.landscape = arg[0] === 2; break;
          case 0x8b: props.xaPage = this.n16(arg, 0); break;
          case 0x8c: props.yaPage = this.n16(arg, 0); break;
          case 0x8d: props.dxaLeft = this.n16(arg, 0); break;
          case 0x8e: props.dxaRight = this.n16(arg, 0); break;
          case 0x8f: props.dyaTop = Math.abs(this.n16s(arg, 0)); break;
          case 0x90: props.dyaBottom = Math.abs(this.n16s(arg, 0)); break;
          case 0x91: props.dzaGutter = this.n16(arg, 0); break;
          default: break;
        }
      }
      const headers: Section["headers"] = [];
      // The header and footer stories, in the order of their flags: even, odd and first page headers and footers.
      const kinds = [["header", "even"], ["header", "default"], ["footer", "even"], ["footer", "default"], ["header", "first"], ["footer", "first"]] as const;
      for (let bit = 0; bit < 6; bit++) {
        if (!(headerFlags & (1 << bit))) continue;
        if (story + 2 >= stories.length) break;
        const start = hddBase + stories[story];
        const end = hddBase + stories[story + 1];
        story++;
        const [kind, type] = kinds[bit];
        if (this.storyHasText(start, end)) headers.push({ kind, type, blocks: this.blocks(start, end, null, false) });
      }
      return { props, blocks: this.blocks(limit.cp, limit.end, sectionEnds, false), headers };
    });
  }

  private storyHasText(start: number, end: number): boolean {
    for (let cp = start; cp < end; cp++) if (this.text[cp] !== 0x0d) return true;
    return false;
  }

  /** The blocks of the story [start, end): paragraphs, and tables made of the paragraphs in cells. */
  private blocks(start: number, end: number, sectionEnds: Set<number> | null, footnote: boolean): Block[] {
    const paragraphs: { p: Paragraph; inTable: boolean; cellEnd: boolean; rowEnd: boolean; table?: TableProps }[] = [];
    let pStart = start;
    for (let cp = start; cp < end; cp++) {
      const ch = this.text[cp];
      if (ch !== 0x0d && ch !== 0x07 && !(ch === 0x0c && sectionEnds?.has(cp)) && cp !== end - 1) continue;
      const para = this.paragraphProps(cp);
      const cellEnd = ch === 0x07;
      const rowEnd = cellEnd && !!para.props.ttp;
      paragraphs.push({ p: this.paragraph(pStart, cp, para, footnote && paragraphs.length === 0), inTable: !!para.props.inTable || cellEnd, cellEnd, rowEnd, table: para.props.table });
      pStart = cp + 1;
    }
    const out: Block[] = [];
    let i = 0;
    while (i < paragraphs.length) {
      if (!paragraphs[i].inTable) {
        out.push(paragraphs[i++].p);
        continue;
      }
      const rows: Row[] = [];
      let cells: Block[][] = [];
      let cell: Block[] = [];
      while (i < paragraphs.length && paragraphs[i].inTable) {
        const item = paragraphs[i++];
        if (item.rowEnd) {
          if (cell.length) cells.push(cell);
          // A row mark with no cell edges of its own takes the row above's.
          const props = item.table?.cells.length ? item.table : rows.length ? rows[rows.length - 1].props : item.table ?? newTableProps();
          rows.push({ props, cells });
          cells = [];
          cell = [];
          continue;
        }
        cell.push(item.p);
        if (item.cellEnd) {
          cells.push(cell);
          cell = [];
        }
      }
      if (rows.length) out.push({ type: "table", rows });
      // Cells after the last row mark (a damaged table): as paragraphs.
      for (const rest of [...cells, cell]) out.push(...rest);
    }
    return out;
  }

  /** The properties of the paragraph whose mark is at `cp`: its style's, its PAPX's, its piece's. */
  private paragraphProps(cp: number): Para {
    const piece = this.pieceAt(cp);
    const papx = this.runAt(this.papRuns, this.fcOf(cp))?.data;
    const styleId = papx && papx.length ? (papx[0] << 24) >> 24 : 0;
    const style = this.style(styleId);
    const para = clonePara(style.para);
    para.props.istd = style.index;
    // A PAPX: the style number, six bytes (the paragraph's line heights), then its sprms.
    if (papx && papx.length > this.papxHeader) this.applyParaSprms(para, this.parseSprms(papx.subarray(this.papxHeader)), true);
    if (piece?.sprms.length) this.applyParaSprms(para, piece.sprms, true);
    return para;
  }

  private paragraph(start: number, mark: number, para: Para, footnote: boolean): Paragraph {
    const style = this.modelStyles[para.props.istd];
    const styleChar = this.styleChars[para.props.istd];
    const runs: Run[] = [];
    let cp = start;
    while (cp < mark) {
      const piece = this.pieceAt(cp);
      const fc = this.fcOf(cp);
      const chp = this.runAt(this.chpRuns, fc);
      let end = mark;
      if (piece) end = Math.min(end, piece.end);
      if (chp) end = Math.min(end, cp + chp.end - fc);
      end = Math.max(end, cp + 1);
      const char = this.charAt(styleChar, para, chp?.data, piece);
      runs.push(this.run(cp, end, char, styleChar, footnote && cp === start));
      cp = end;
    }
    const markChar = this.charAt(styleChar, para, this.runAt(this.chpRuns, this.fcOf(mark))?.data, this.pieceAt(mark));
    return { type: "paragraph", props: diffPara(para.props, style.para), mark: this.runProps(markChar, styleChar), runs: runs.filter((r) => r.content.length) };
  }

  private charAt(styleChar: Char, para: Para, chpx: Uint8Array | undefined, piece: Piece | undefined): Char {
    const char = cloneChar(styleChar);
    for (const sprm of para.charSprms) this.applyCharSprm(char, sprm, styleChar);
    if (chpx) this.applyChpx(char, chpx);
    if (piece) for (const sprm of piece.sprms) if (this.isCharSprm(sprm.op)) this.applyCharSprm(char, sprm, styleChar);
    return char;
  }

  private run(start: number, end: number, char: Char, styleChar: Char, noteStart: boolean): Run {
    const content: Inline[] = [];
    let text = "";
    const flush = (): void => {
      if (text) content.push({ kind: "text", text });
      text = "";
    };
    const decode = this.decoder(char.fontId);
    for (let cp = start; cp < end; cp++) {
      const ch = this.text[cp];
      const note = ch >= 0x20 ? this.noteRefs.get(cp) : undefined;
      if (note) {
        // A footnote reference with a mark of its own: the character at the reference.
        flush();
        content.push({ kind: "footnote", id: note, customMark: decode(ch) });
        continue;
      }
      if (ch >= 0x20) {
        text += decode(ch);
        continue;
      }
      flush();
      switch (ch) {
        case 0x09: content.push({ kind: "tab" }); break;
        case 0x0b: content.push({ kind: "break" }); break;
        case 0x0c: content.push({ kind: "page" }); break;
        case 0x0e: content.push({ kind: "column" }); break;
        case 0x1e: content.push({ kind: "noBreakHyphen" }); break;
        case 0x1f: content.push({ kind: "softHyphen" }); break;
        default:
          if (char.special) this.special(ch, cp, start, char, noteStart, content);
          break;
      }
    }
    flush();
    return { props: this.runProps(char, styleChar), content };
  }

  /** A special character: a picture, a note reference, a field. */
  protected special(ch: number, cp: number, start: number, char: Char, noteStart: boolean, content: Inline[]): void {
    switch (ch) {
      case 0x01: {
        const picture = char.picture !== undefined ? this.picture(char) : undefined;
        if (picture) content.push(picture);
        break;
      }
      case 0x02:
      case 0x05: {
        const id = this.noteRefs.get(cp);
        if (id) content.push({ kind: "footnote", id });
        else if (noteStart && cp === start) content.push({ kind: "footnoteMark" });
        break;
      }
      case 0x13: content.push({ kind: "fieldBegin" }); break;
      case 0x14: content.push({ kind: "fieldSeparate" }); break;
      case 0x15: content.push({ kind: "fieldEnd" }); break;
      default: break;
    }
  }

  /** A run's properties: where its character formatting differs from its paragraph style's. */
  private runProps(char: Char, style: Char): RunProps {
    const props: RunProps = { toggles: {} };
    if (char.fontId !== style.fontId) props.ftc0 = props.ftc2 = this.font(char.fontId);
    if (char.hps !== style.hps) props.hps = char.hps;
    FLAG_NAMES.forEach((name, i) => {
      if (char.flags[i] !== style.flags[i]) props.toggles[name] = char.flags[i] ? 1 : 0;
    });
    if (char.kul !== style.kul) props.kul = char.kul;
    if (char.hpsPos !== style.hpsPos) props.hpsPos = char.hpsPos;
    if (char.dxaSpace !== style.dxaSpace) props.dxaSpace = char.dxaSpace;
    if (char.ico !== style.ico) props.color = ICO[char.ico] ?? "auto";
    return props;
  }

  /** A style's character formatting, all of it. */
  private fullRunProps(char: Char): RunProps {
    const props: RunProps = { toggles: {}, ftc0: this.font(char.fontId), ftc2: this.font(char.fontId), hps: char.hps };
    FLAG_NAMES.forEach((name, i) => {
      if (char.flags[i]) props.toggles[name] = 1;
    });
    if (char.kul) props.kul = char.kul;
    if (char.hpsPos) props.hpsPos = char.hpsPos;
    if (char.dxaSpace) props.dxaSpace = char.dxaSpace;
    if (char.ico) props.color = ICO[char.ico] ?? "auto";
    return props;
  }

  // ---- Sprms --------------------------------------------------------------

  /** A grpprl's sprms, up to the first one whose size is unknown. */
  protected parseSprms(b: Uint8Array): Sprm[] {
    const out: Sprm[] = [];
    let at = 0;
    while (at < b.length) {
      const op = b[at];
      const size = this.sprmSizes[op];
      if (size === undefined) break;
      let start = at + 1;
      let length = size;
      if (size === -1) {
        length = b[at + 1] ?? 0;
        start = at + 2;
      } else if (size === -2) {
        length = Math.max(0, this.n16(b, at + 1) - 1);
        start = at + 3;
      }
      if (start + length > b.length) break;
      out.push({ op, arg: b.subarray(start, start + length) });
      at = start + length;
    }
    return out;
  }

  /** A border word (BRC10): line widths and the space between them in the low 9 bits, the space to the text, a shadow bit. */
  protected border10(v: number): Border | undefined {
    const line = v & 0x1ff;
    const base = { color: "auto", space: (v >>> 9) & 0x1f, shadow: (v & 0x4000) !== 0, frame: false };
    switch (line) {
      case 0: return undefined;
      case 0x49: return { ...base, type: 3, width: 4 };
      case 0x80: return { ...base, type: 1, width: 12 };
      case 0x180: return { ...base, type: 6, width: 4 };
      case 0x1c0: return { ...base, type: 5, width: 2 };
      default: return { ...base, type: 1, width: 4 };
    }
  }

  /** A border word (BRC): a line width in 0.75 points, a type, a shadow bit, a color and the space to the text. */
  protected border(v: number): Border | undefined {
    const width = v & 7;
    const type = (v >>> 3) & 3;
    if (type === 0 || width === 0) return undefined;
    return { type: type === 2 ? 2 : type === 3 ? 3 : 1, width: width * 6, color: ICO[(v >>> 6) & 0x1f] ?? "auto", space: (v >>> 11) & 0x1f, shadow: (v & 0x20) !== 0, frame: false };
  }

  private applyParaSprms(para: Para, sprms: Sprm[], styled: boolean): void {
    const p = para.props;
    for (const sprm of sprms) {
      const { op, arg } = sprm;
      if (this.isCharSprm(op)) {
        para.charSprms.push(sprm);
        continue;
      }
      if (op >= 0x92) {
        this.applyTableSprm(p, op, arg);
        continue;
      }
      switch (op) {
        case 0x02: if (styled) p.istd = this.style((arg[0] << 24) >> 24).index; break;
        case 0x05: p.jc = arg[0] & 3; break;
        case 0x07: p.keepLines = arg[0] !== 0; break;
        case 0x08: p.keepNext = arg[0] !== 0; break;
        case 0x09: p.pageBreakBefore = arg[0] !== 0; break;
        case 0x0a: {
          const line = [1, 2, 3, 1][arg[0]] ?? 1;
          para.borderStyle = { type: line, width: line === 2 ? 12 : 4, color: "auto", space: 0, shadow: arg[0] === 3, frame: false };
          break;
        }
        case 0x0b: {
          const border = para.borderStyle ?? { type: 1, width: 4, color: "auto", space: 0, shadow: false, frame: false };
          const v = arg[0];
          if (v & 1) p.borders.top = border;
          if (v & 2) p.borders.bottom = border;
          if (v & 4 || v === 0x10) p.borders.left = border;
          if (v & 8) p.borders.right = border;
          break;
        }
        case 0x0f:
        case 0x17: this.applyTabs(p, arg, op === 0x17 ? 4 : 2); break;
        case 0x10: p.dxaRight = this.n16s(arg, 0); break;
        case 0x11: p.dxaLeft = this.n16s(arg, 0); break;
        case 0x13: p.dxaLeft1 = this.n16s(arg, 0); break;
        case 0x14: {
          // Positive: at least; negative: exactly; 0: single, as the font needs.
          const v = this.n16s(arg, 0);
          p.lineSpacing = v || 240;
          p.lineMultiple = v === 0;
          break;
        }
        case 0x15: p.dyaBefore = this.n16s(arg, 0); break;
        case 0x16: p.dyaAfter = this.n16s(arg, 0); break;
        case 0x18: p.inTable = arg[0] !== 0; break;
        case 0x19: p.ttp = arg[0] !== 0; break;
        case 0x1a: p.frame.x = this.n16(arg, 0); break;
        case 0x1b: p.frame.y = this.n16(arg, 0); break;
        case 0x1c: p.frame.width = this.n16(arg, 0); break;
        case 0x1d: p.frame.pcHorz = arg[0] & 3; p.frame.pcVert = arg[0] & 4 ? 1 : 0; break;
        case 0x1e: p.borders.top = this.border10(this.n16(arg, 0)); break;
        case 0x1f: p.borders.left = this.border10(this.n16(arg, 0)); break;
        case 0x20: p.borders.bottom = this.border10(this.n16(arg, 0)); break;
        case 0x21: p.borders.right = this.border10(this.n16(arg, 0)); break;
        case 0x22: p.borders.between = this.border10(this.n16(arg, 0)); break;
        case 0x24: p.frame.dxaFromText = this.n16(arg, 0); break;
        case 0x26: p.borders.top = this.border(this.n16(arg, 0)); break;
        case 0x27: p.borders.left = this.border(this.n16(arg, 0)); break;
        case 0x28: p.borders.bottom = this.border(this.n16(arg, 0)); break;
        case 0x29: p.borders.right = this.border(this.n16(arg, 0)); break;
        case 0x2a: p.borders.between = this.border(this.n16(arg, 0)); break;
        case 0x2c: p.noAutoHyph = arg[0] !== 0; break;
        case 0x33: p.widowControl = arg[0] !== 0; break;
        default: break;
      }
    }
  }

  /** A tab change: the positions cleared (each `deletedSize` bytes), then the positions set and a descriptor byte per tab. */
  private applyTabs(p: ParaProps, arg: Uint8Array, deletedSize: number): void {
    let at = 0;
    const deleted = arg[at++] ?? 0;
    for (let i = 0; i < deleted; i++) {
      const pos = this.n16s(arg, at + i * deletedSize);
      p.tabs = p.tabs.filter((t) => Math.abs(t.pos - pos) > 1);
    }
    at += deleted * deletedSize;
    const added = arg[at++] ?? 0;
    for (let i = 0; i < added && at + 2 * added + i < arg.length; i++) {
      const pos = this.n16s(arg, at + 2 * i);
      p.tabs = p.tabs.filter((t) => Math.abs(t.pos - pos) > 1);
      p.tabs.push({ pos, ...this.tab(arg[at + 2 * added + i]) });
    }
  }

  private applyTableSprm(p: ParaProps, op: number, arg: Uint8Array): void {
    const t = (p.table ??= newTableProps());
    switch (op) {
      case 0x92: t.jc = this.n16(arg, 0) & 3; break;
      case 0x94: t.dxaGapHalf = this.n16s(arg, 0); break;
      case 0x99: t.dyaRowHeight = this.n16s(arg, 0); break;
      case 0x98:
      case 0x9a: {
        // The cell count, the cell edges, then per cell a flag word and four borders (top, left, bottom, right):
        // 0x98's in the old border words, 0x9a's (Word for Windows 2.0's sprmTDefTable; a Mac's 0x9a is shading) in the new.
        if (op === 0x9a && this.sprmSizes[op] !== -2) break;
        const border = (v: number): Border | undefined => (op === 0x98 ? this.border10(v) : this.border(v));
        const count = arg[0] ?? 0;
        t.left = this.n16s(arg, 1);
        t.cells = [];
        const cellsAt = 1 + 2 * (count + 1);
        for (let i = 0; i < count; i++) {
          const cell: CellProps = {
            right: this.n16s(arg, 1 + 2 * (i + 1)), horzMerge: 0, vertMerge: 0, vertAlign: 0, textFlow: 0, ftsWidth: 0, width: 0,
            fitText: false, noWrap: false, hideMark: false, borders: {}, margins: {},
          };
          const at = cellsAt + 10 * i;
          if (at + 10 <= arg.length) {
            cell.borders.top = border(this.n16(arg, at + 2));
            cell.borders.left = border(this.n16(arg, at + 4));
            cell.borders.bottom = border(this.n16(arg, at + 6));
            cell.borders.right = border(this.n16(arg, at + 8));
          }
          t.cells.push(cell);
        }
        break;
      }
      case 0x9d: {
        // Shading for the cells [first, last): a share of black in hundredths of a percent.
        const fill = grey(this.n16(arg, 3));
        for (let i = arg[1]; i < arg[2] && i < t.cells.length; i++) t.cells[i].shading = { pattern: 0, fore: "auto", fill };
        break;
      }
      case 0xa3: {
        const border = this.border10(this.n16(arg, 3));
        for (let i = arg[0]; i < arg[1] && i < t.cells.length; i++) {
          const sides = t.cells[i].borders;
          if (arg[2] & 1) sides.top = border;
          if (arg[2] & 2) sides.left = border;
          if (arg[2] & 4) sides.bottom = border;
          if (arg[2] & 8) sides.right = border;
        }
        break;
      }
      default: break;
    }
  }
}

// ---- Helpers -----------------------------------------------------------------

export function cloneChar(c: Char): Char {
  return { ...c, flags: [...c.flags] };
}

function clonePara(p: Para): Para {
  return { props: structuredClone(p.props), charSprms: [...p.charSprms], borderStyle: p.borderStyle };
}

function grey(tenThousandths: number): string {
  const v = Math.max(0, Math.min(255, Math.round(255 * (1 - tenThousandths / 10000))));
  const h = v.toString(16).toUpperCase().padStart(2, "0");
  return h + h + h;
}

/** A paragraph's properties where they differ from its style's: what its pPr must say. */
function diffPara(p: ParaProps, style: ParaProps): ParaProps {
  const out = newParaProps(p.istd);
  const keys = new Set([...Object.keys(p), ...Object.keys(style)]) as Set<keyof ParaProps>;
  for (const key of keys) {
    if (key === "istd" || key === "tabs" || key === "borders" || key === "frame" || key === "table") continue;
    if (p[key] !== style[key] && p[key] !== undefined) (out as unknown as Record<string, unknown>)[key] = p[key];
  }
  // The line spacing is its amount and its rule together.
  if (out.lineSpacing !== undefined || out.lineMultiple !== undefined) {
    out.lineSpacing = p.lineSpacing;
    out.lineMultiple = p.lineMultiple;
  }
  if (JSON.stringify(p.tabs) !== JSON.stringify(style.tabs)) {
    out.tabs = [...p.tabs];
    for (const t of style.tabs) if (!p.tabs.some((own) => own.pos === t.pos)) out.tabs.push({ pos: t.pos, jc: "clear", leader: 0 });
  }
  for (const side of Object.keys(p.borders) as (keyof ParaProps["borders"])[]) {
    if (JSON.stringify(p.borders[side]) !== JSON.stringify(style.borders[side])) out.borders[side] = p.borders[side];
  }
  // Only a frame with a position of its own: one with no position stays in the text's flow.
  if ((p.frame.x || p.frame.y) && JSON.stringify(p.frame) !== JSON.stringify(style.frame)) out.frame = p.frame;
  if (p.table) out.table = p.table;
  return out;
}

export function popcount(v: number): number {
  let n = 0;
  for (; v; v &= v - 1) n++;
  return n;
}

/** A style sheet list: per entry a size byte and that many bytes; 0xFF is an entry not in use (null). */
function entriesOf(zone: Uint8Array): (Uint8Array | null)[] {
  const out: (Uint8Array | null)[] = [];
  let at = 0;
  while (at < zone.length) {
    const size = zone[at];
    if (size === 0xff) {
      out.push(null);
      at++;
      continue;
    }
    out.push(zone.subarray(at + 1, at + 1 + size));
    at += 1 + size;
  }
  return out;
}
