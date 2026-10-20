// Word for the Macintosh 4.0 and 5.x (signature FE 37 00): a header of
// file positions much like Word for Windows 2.0's FIB, big-endian; the text
// in the Mac OS Roman code page, in pieces when the file was fast-saved;
// character and paragraph formatting in 512-byte FKP pages; a style sheet;
// sections; footnotes, headers and footers as stories after the main text;
// pictures as Macintosh PICTs.
//
// Written from the file format as observed in Word's files, with the
// libmwaw project's notes on it (MsWrdParser) as a guide to the structures.
import { ParaProps, RunProps, SectionProps, TableProps, Border, CellProps, newParaProps, newSectionProps, newTableProps } from "../doc/sprm.js";
import { symbolText } from "../pdf/dib.js";
import { Block, Inline, OldDocument, Paragraph, Picture, Row, Run, Section, Style } from "./model.js";

const MAC_ROMAN = new TextDecoder("macintosh");

/** The Mac fonts by their system number, for a Word 4 file (no font table of its own) and a font a Word 5 table leaves out. */
const MAC_FONTS: Record<number, string> = {
  0: "Chicago", 1: "Geneva", 2: "New York", 3: "Geneva", 4: "Monaco", 5: "Venice", 6: "London", 7: "Athens", 8: "San Francisco", 9: "Toronto",
  11: "Cairo", 12: "Los Angeles", 13: "Zapf Dingbats", 14: "Bookman", 15: "Helvetica Narrow", 16: "Palatino", 18: "Zapf Chancery", 20: "Times",
  21: "Helvetica", 22: "Courier", 23: "Symbol", 24: "Mobile", 33: "Avant Garde", 34: "New Century Schoolbook",
};

/** The standard styles by their negative number: Word for Windows 2.0's list, which Word for the Macintosh shares. */
const STANDARD_STYLES: Record<number, string> = {
  [-1]: "Normal Indent", [-2]: "heading 1", [-3]: "heading 2", [-4]: "heading 3", [-5]: "heading 4", [-6]: "heading 5", [-7]: "heading 6",
  [-8]: "heading 7", [-9]: "heading 8", [-10]: "heading 9", [-11]: "footnote text", [-12]: "footnote reference", [-13]: "annotation text",
  [-14]: "annotation reference", [-15]: "header", [-16]: "footer", [-17]: "index heading", [-18]: "line number",
};

/** Whether `head` starts a Word for the Macintosh 4.0 or 5.x document. */
export function isMacWordFile(head: Uint8Array): boolean {
  return head.length >= 4 && head[0] === 0xfe && head[1] === 0x37 && head[2] === 0x00 && (head[3] === 0x1c || head[3] === 0x23);
}

/** A character's formatting as Word for the Macintosh keeps it: a Mac font number, and the rest. */
interface MacChar {
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
}

/** A paragraph's properties: ParaProps for the XML writer, with the table row definition and the character changes a paragraph sprm makes. */
interface MacPara {
  props: ParaProps;
  /** The character sprms (0x3c-0x53) among the paragraph's, applied to its text. */
  charSprms: Sprm[];
  /** Old-style border (sprm 0x0a): the line every side sprm 0x0b turns on gets. */
  borderStyle?: Border;
}

interface Sprm {
  op: number;
  arg: Uint8Array;
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

const FLAG_NAMES = ["b", "i", "strike", "outline", "shadow", "smallCaps", "caps", "vanish"] as const;

/** A Word for the Macintosh 4.0 or 5.x document as an OldDocument. */
export function readMacWord(bytes: Uint8Array): OldDocument {
  return new MacReader(bytes).read();
}

class MacReader {
  private readonly version: number;
  private readonly fcMin: number;
  private readonly ccpText: number;
  private readonly ccpFtn: number;
  private readonly ccpHdd: number;
  private readonly entries: { fc: number; cb: number }[] = [];
  private readonly extraEntries: { fc: number; cb: number }[] = [];
  private pieces: Piece[] = [];
  private chpRuns: FkpRun[] = [];
  private papRuns: FkpRun[] = [];
  private text = new Uint8Array(0);
  private readonly fonts: string[] = [];
  private readonly fontIndex = new Map<number, number>();
  private readonly macFontNames = new Map<number, string>();
  /** The styles by their Word number (negative for the standard ones): the resolved properties and the model index. */
  private readonly styles = new Map<number, { name: string; para: MacPara; char: MacChar; next: number; index: number }>();
  private readonly modelStyles: Style[] = [];
  /** The resolved character formatting of each model style. */
  private readonly styleChars: MacChar[] = [];
  private readonly pictures: Picture[] = [];
  private readonly pictureIndex = new Map<number, number>();
  private readonly noteRefs = new Map<number, number>();
  private readonly footnotes: Block[][] = [];

  constructor(private readonly bytes: Uint8Array) {
    this.version = bytes[3] === 0x1c ? 4 : 5;
    this.fcMin = this.u32(0x14);
    this.ccpText = this.u32(0x24);
    this.ccpFtn = this.u32(0x28);
    this.ccpHdd = this.u32(0x2c);
    for (let i = 0; i < 20; i++) this.entries.push({ fc: this.u32(0x40 + 6 * i), cb: this.u16(0x44 + 6 * i) });
    // Word 5 has a second list after two counts: object names, font names, objects, ...
    if (this.version === 5) for (let i = 0; i < 11; i++) this.extraEntries.push({ fc: this.u32(0xbc + 6 * i), cb: this.u16(0xc0 + 6 * i) });
  }

  // ---- Bytes --------------------------------------------------------------

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
    return ((this.u8(at) << 24) | (this.u8(at + 1) << 16) | (this.u8(at + 2) << 8) | this.u8(at + 3)) >>> 0;
  }
  private entry(index: number): Uint8Array {
    const { fc, cb } = this.entries[index];
    return cb > 0 && fc + cb <= this.bytes.length ? this.bytes.subarray(fc, fc + cb) : new Uint8Array(0);
  }

  // ---- The document -------------------------------------------------------

  read(): OldDocument {
    this.readPieces();
    this.chpRuns = this.readFkps(9);
    this.papRuns = this.readFkps(10);
    this.readFontNames();
    this.readStyles();
    const mainEnd = this.ccpText;
    this.readNoteRefs();
    const sections = this.readSections(mainEnd);
    this.readFootnotes();
    const dop = this.entry(15);
    return {
      fonts: this.fonts,
      styles: this.modelStyles,
      sections,
      footnotes: this.footnotes,
      pictures: this.pictures,
      defaultTab: dop.length >= 16 ? be16(dop, 14) || undefined : undefined,
    };
  }

  /** The text's pieces: the clx of a fast-saved file, else one piece from fcMin. */
  private readPieces(): void {
    const total = this.ccpText + this.ccpFtn + this.ccpHdd + (this.ccpFtn || this.ccpHdd ? 1 : 0);
    const clx = this.entry(18);
    const grpprls: Uint8Array[] = [];
    let at = 0;
    while (at < clx.length && clx[at] === 1) {
      const cb = be16(clx, at + 1);
      grpprls.push(clx.subarray(at + 3, at + 3 + cb));
      at += 3 + cb;
    }
    if (at < clx.length && clx[at] === 2) {
      const cb = be16(clx, at + 1);
      const plc = clx.subarray(at + 3, at + 3 + cb);
      const n = Math.floor((plc.length - 4) / 12);
      for (let i = 0; i < n; i++) {
        const cp = be32(plc, 4 * i);
        const end = be32(plc, 4 * (i + 1));
        const pcd = 4 * (n + 1) + 8 * i;
        const prm = be16(plc, pcd + 6);
        let sprms: Sprm[] = [];
        if (prm & 0x8000) sprms = parseSprms(grpprls[(prm >>> 8) & 0x7f] ?? new Uint8Array(0));
        else if (prm >>> 8) sprms = parseSprms(Uint8Array.of(prm >>> 8, prm & 0xff));
        if (end > cp) this.pieces.push({ cp, end, fc: be32(plc, pcd + 2), sprms });
      }
    }
    if (this.pieces.length === 0) this.pieces.push({ cp: 0, end: total, fc: this.fcMin, sprms: [] });
    const length = this.pieces[this.pieces.length - 1].end;
    this.text = new Uint8Array(length);
    for (const piece of this.pieces) {
      const from = Math.min(piece.fc, this.bytes.length);
      const to = Math.min(piece.fc + piece.end - piece.cp, this.bytes.length);
      this.text.set(this.bytes.subarray(from, to), piece.cp);
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

  /**
   * The runs of the FKP pages the bin table at entry `index` lists. A page
   * that ends before the next listed one begins is followed by the pages
   * after it in the file, which the table leaves out.
   */
  private readFkps(index: number): FkpRun[] {
    const bte = this.entry(index);
    const n = Math.floor((bte.length - 4) / 6);
    const runs: FkpRun[] = [];
    const paragraph = index === 10;
    for (let i = 0; i < n; i++) {
      const limit = be32(bte, 4 * (i + 1));
      let pn = be16(bte, 4 * (n + 1) + 2 * i);
      let expected = be32(bte, 4 * i);
      for (let pages = 0; pages < 256; pages++, pn++) {
        const page = pn * 512;
        if (page + 512 > this.bytes.length) break;
        const crun = this.u8(page + 511);
        if (crun === 0 || 5 * crun + 4 > 511 || this.u32(page) !== expected) break;
        expected = this.u32(page + 4 * crun);
        for (let k = 0; k < crun; k++) {
          const fc = this.u32(page + 4 * k);
          const end = this.u32(page + 4 * (k + 1));
          const offset = this.u8(page + 4 * (crun + 1) + k) * 2;
          let data: Uint8Array | undefined;
          if (offset) {
            const size = this.u8(page + offset);
            // A PAPX's size counts words after its size byte; a CHPX's, bytes.
            const length = paragraph ? 2 * size : size;
            data = this.bytes.subarray(page + offset + 1, Math.min(page + offset + 1 + length, page + 511));
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

  private readFontNames(): void {
    const names = this.version === 5 ? this.extraEntries[1] : undefined;
    if (!names || names.cb < 2 || names.fc + names.cb > this.bytes.length) return;
    const zone = this.bytes.subarray(names.fc, names.fc + names.cb);
    const count = be16(zone, 0);
    let at = 2;
    for (let i = 0; i < count && at + 5 <= zone.length; i++) {
      const id = be16(zone, at + 2);
      const length = zone[at + 4];
      this.macFontNames.set(id, MAC_ROMAN.decode(zone.subarray(at + 5, at + 5 + length)));
      at += 5 + length;
    }
  }

  /** The model's font index for a Mac font number. */
  private font(id: number): number {
    let index = this.fontIndex.get(id);
    if (index === undefined) {
      index = this.fonts.length;
      this.fonts.push(this.macFontNames.get(id) ?? MAC_FONTS[id] ?? "Times");
      this.fontIndex.set(id, index);
    }
    return index;
  }

  /**
   * The style sheet: the count of standard styles; the names, the
   * character and the paragraph properties, each a list with an entry per
   * style; and per style the next style and the one it is based on.
   */
  private readStyles(): void {
    const stsh = this.entry(1).length ? this.entry(1) : this.entry(0);
    const base: MacChar = { fontId: 2, hps: 24, flags: new Array(8).fill(false), kul: 0, hpsPos: 0, dxaSpace: 0, ico: 0 };
    const defined: { id: number; name?: string; chpx?: Uint8Array; papx?: Uint8Array; next: number; basedOn: number }[] = [];
    if (stsh.length >= 8) {
      const standard = be16s(stsh, 0);
      const zones: Uint8Array[] = [];
      let at = 2;
      for (let z = 0; z < 3 && at + 2 <= stsh.length; z++) {
        const size = be16(stsh, at);
        zones.push(stsh.subarray(at + 2, Math.min(at + size, stsh.length)));
        at += Math.max(size, 2);
      }
      const names = entriesOf(zones[0] ?? new Uint8Array(0));
      const chpxs = entriesOf(zones[1] ?? new Uint8Array(0));
      const papxs = entriesOf(zones[2] ?? new Uint8Array(0));
      const count = at + 2 <= stsh.length ? be16(stsh, at) : 0;
      for (let i = 0; i < Math.max(names.length, count); i++) {
        const name = names[i];
        if (name === null) continue;
        const next = at + 3 + 2 * i < stsh.length ? (stsh[at + 2 + 2 * i] << 24) >> 24 : i - standard;
        const basedOn = at + 3 + 2 * i < stsh.length ? (stsh[at + 3 + 2 * i] << 24) >> 24 : -34;
        defined.push({
          id: i - standard,
          name: name && name.length ? MAC_ROMAN.decode(name) : undefined,
          chpx: chpxs[i] ?? undefined,
          papx: papxs[i] ?? undefined,
          next,
          basedOn,
        });
      }
    }
    if (!defined.some((s) => s.id === 0)) defined.push({ id: 0, next: 0, basedOn: -34 });
    const byId = new Map(defined.map((s) => [s.id, s]));
    const resolving = new Set<number>();
    const resolve = (id: number): { para: MacPara; char: MacChar } => {
      const done = this.styles.get(id);
      if (done) return done;
      const style = byId.get(id);
      let para: MacPara = { props: newParaProps(0), charSprms: [] };
      let char: MacChar = { ...base, flags: [...base.flags] };
      if (style && style.basedOn !== -34 && style.basedOn !== id && byId.has(style.basedOn) && !resolving.has(style.basedOn)) {
        resolving.add(id);
        const parent = resolve(style.basedOn);
        resolving.delete(id);
        para = clonePara(parent.para);
        char = cloneChar(parent.char);
      }
      if (style?.chpx) applyChpx(char, style.chpx);
      // A style's paragraph entry: its own number, two bytes, then (Word 4 and 5) four more, then its sprms.
      if (style?.papx && style.papx.length > 7) {
        const sprms = parseSprms(style.papx.subarray(7));
        applyParaSprms(para, sprms);
        for (const sprm of sprms) applyCharSprm(char, sprm, char);
      }
      para.props.istd = 0;
      para.charSprms = [];
      const name = id === 0 ? "Normal" : style?.name ?? STANDARD_STYLES[id] ?? `Style ${id}`;
      const entry = { name, para, char, next: style?.next ?? id, index: -1 };
      this.styles.set(id, entry);
      return entry;
    };
    // Normal first: the model's style 0.
    for (const id of [0, ...defined.map((s) => s.id).filter((id) => id !== 0)]) {
      const style = resolve(id);
      this.styles.get(id)!.index = this.modelStyles.length;
      this.modelStyles.push({ name: this.styles.get(id)!.name, para: style.para.props, run: this.fullRunProps(style.char) });
      this.styleChars.push(style.char);
    }
    for (const [, style] of this.styles) {
      const next = this.styles.get(style.next);
      if (next) this.modelStyles[style.index].next = next.index;
    }
    for (const [, style] of this.styles) this.modelStyles[style.index].para = { ...style.para.props, istd: style.index };
  }

  private style(id: number): { name: string; para: MacPara; char: MacChar; next: number; index: number } {
    return this.styles.get(id) ?? this.styles.get(0)!;
  }

  // ---- Stories ------------------------------------------------------------

  /** The footnote references in the main text: CP to note number. */
  private readNoteRefs(): void {
    const refs = this.entry(2);
    const n = Math.floor((refs.length - 4) / 6);
    for (let i = 0; i < n; i++) this.noteRefs.set(be32(refs, 4 * i), i + 1);
  }

  private readFootnotes(): void {
    const txt = this.entry(3);
    const cps: number[] = [];
    for (let at = 0; at + 4 <= txt.length; at += 4) cps.push(be32(txt, at));
    const base = this.ccpText;
    for (let i = 0; i < this.noteRefs.size && i + 1 < cps.length; i++) {
      const start = base + cps[i];
      const end = Math.min(base + cps[i + 1], base + this.ccpFtn);
      this.footnotes.push(end > start ? this.blocks(start, end, null, true) : []);
    }
  }

  /** The sections of the main text, with their headers and footers. */
  private readSections(mainEnd: number): Section[] {
    const sed = this.entry(4);
    const n = Math.floor((sed.length - 4) / 10);
    const limits: { cp: number; end: number; sepx?: Uint8Array }[] = [];
    for (let i = 0; i < n; i++) {
      const cp = be32(sed, 4 * i);
      const end = Math.min(be32(sed, 4 * (i + 1)), mainEnd);
      const fcSepx = be32(sed, 4 * (n + 1) + 6 * i + 2);
      let sepx: Uint8Array | undefined;
      if (fcSepx !== 0xffffffff && fcSepx < this.bytes.length) sepx = this.bytes.subarray(fcSepx + 1, fcSepx + 1 + this.u8(fcSepx));
      if (end > cp) limits.push({ cp, end, sepx });
    }
    if (limits.length === 0) limits.push({ cp: 0, end: mainEnd });
    limits[limits.length - 1].end = mainEnd;
    const sectionEnds = new Set(limits.slice(0, -1).map((l) => l.end - 1));
    const hdd = this.entry(8);
    const stories: number[] = [];
    for (let at = 0; at + 4 <= hdd.length; at += 4) stories.push(be32(hdd, at));
    const hddBase = this.ccpText + this.ccpFtn;
    let story = 0;
    return limits.map((limit) => {
      const props = this.pageProps();
      let headerFlags = 0;
      for (const { op, arg } of parseSprms(limit.sepx ?? new Uint8Array(0))) {
        switch (op) {
          case 0x75: props.bkc = arg[0]; break;
          case 0x76: props.titlePage = arg[0] !== 0; break;
          case 0x77: props.columns = be16(arg, 0) + 1; break;
          case 0x78: props.dxaColumns = be16(arg, 0); break;
          case 0x80: headerFlags = arg[0]; break;
          case 0x83: props.dyaHdrTop = be16(arg, 0); break;
          case 0x84: props.dyaHdrBottom = be16(arg, 0); break;
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

  /** The page size and margins, the document's (Word for the Macintosh has them once, in the document's properties). */
  private pageProps(): SectionProps {
    const props = newSectionProps();
    const dop = this.entry(15);
    if (dop.length >= 12) {
      const height = be16(dop, 0);
      const width = be16(dop, 2);
      const margins = [0, 1, 2, 3].map((k) => Math.abs((be16(dop, 4 + 2 * k) << 16) >> 16));
      if (height > margins[0] + margins[2] && width > margins[1] + margins[3]) {
        props.yaPage = height;
        props.xaPage = width;
        props.dyaTop = margins[0];
        props.dxaLeft = margins[1];
        props.dyaBottom = margins[2];
        props.dxaRight = margins[3];
        props.landscape = width > height;
      }
    }
    return props;
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
      const inTable = !!para.props.inTable;
      const cellEnd = ch === 0x07;
      const rowEnd = cellEnd && !!para.props.ttp;
      paragraphs.push({ p: this.paragraph(pStart, cp, para, footnote && paragraphs.length === 0), inTable: inTable || cellEnd, cellEnd, rowEnd, table: para.props.table });
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
          rows.push({ props: item.table ?? newTableProps(), cells });
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
  private paragraphProps(cp: number): MacPara {
    const piece = this.pieceAt(cp);
    const fc = piece ? piece.fc + cp - piece.cp : cp + this.fcMin;
    const papx = this.runAt(this.papRuns, fc)?.data;
    const styleId = papx && papx.length ? (papx[0] << 24) >> 24 : 0;
    const para = clonePara(this.style(styleId).para);
    para.props.istd = this.style(styleId).index;
    if (papx && papx.length > 7) applyParaSprms(para, parseSprms(papx.subarray(7)), this);
    if (piece?.sprms.length) applyParaSprms(para, piece.sprms, this);
    return para;
  }

  /** Called by applyParaSprms for sprm 0x02: a style number to the model's index. */
  styleIndex(id: number): number {
    return this.style(id).index;
  }

  private paragraph(start: number, mark: number, para: MacPara, footnote: boolean): Paragraph {
    const style = this.modelStyles[para.props.istd];
    const styleChar = this.styleChars[para.props.istd];
    const runs: Run[] = [];
    let cp = start;
    while (cp < mark) {
      const piece = this.pieceAt(cp);
      const fc = piece ? piece.fc + cp - piece.cp : cp + this.fcMin;
      const chp = this.runAt(this.chpRuns, fc);
      let end = mark;
      if (piece) end = Math.min(end, piece.end);
      if (chp) end = Math.min(end, cp + chp.end - fc);
      const char = this.charAt(styleChar, para, chp?.data, piece);
      runs.push(this.run(cp, Math.max(end, cp + 1), char, styleChar, footnote && cp === start));
      cp = Math.max(end, cp + 1);
    }
    const markChar = this.charAt(styleChar, para, this.runAt(this.chpRuns, this.fcOf(mark))?.data, this.pieceAt(mark));
    return { type: "paragraph", props: diffPara(para.props, style.para), mark: this.runProps(markChar, styleChar), runs: runs.filter((r) => r.content.length) };
  }

  private fcOf(cp: number): number {
    const piece = this.pieceAt(cp);
    return piece ? piece.fc + cp - piece.cp : cp + this.fcMin;
  }

  private charAt(styleChar: MacChar, para: MacPara, chpx: Uint8Array | undefined, piece: Piece | undefined): MacChar {
    const char = cloneChar(styleChar);
    for (const sprm of para.charSprms) applyCharSprm(char, sprm, styleChar);
    if (chpx) applyChpx(char, chpx);
    if (piece) for (const sprm of piece.sprms) applyCharSprm(char, sprm, styleChar);
    return char;
  }

  private run(start: number, end: number, char: MacChar, styleChar: MacChar, noteStart: boolean): Run {
    const content: Inline[] = [];
    let text = "";
    const flush = (): void => {
      if (text) content.push({ kind: "text", text });
      text = "";
    };
    const fontName = this.fonts[this.font(char.fontId)];
    const symbol = fontName === "Symbol";
    for (let cp = start; cp < end; cp++) {
      const ch = this.text[cp];
      if (ch >= 0x20) {
        text += symbol ? symbolText(Uint8Array.of(ch)) : MAC_ROMAN_CHARS[ch];
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
        case 0x11: text += "⌘"; break;
        case 0x14: text += ""; break;
        case 0x02: content.push({ kind: "field", code: "PAGE", result: "1" }); break;
        case 0x03: case 0x10: case 0x13: case 0x15: case 0x16: case 0x17: case 0x1a: case 0x1b: case 0x1d:
          content.push({ kind: "field", code: "DATE", result: "" });
          break;
        case 0x04: case 0x18: case 0x19: content.push({ kind: "field", code: "TIME", result: "" }); break;
        case 0x05: {
          const id = this.noteRefs.get(cp);
          if (id) content.push({ kind: "footnote", id });
          else if (noteStart && cp === start) content.push({ kind: "footnoteMark" });
          break;
        }
        case 0x01: {
          const picture = char.picture !== undefined ? this.picture(char.picture) : undefined;
          if (picture) content.push(picture);
          break;
        }
        default: break;
      }
    }
    flush();
    return { props: this.runProps(char, styleChar), content };
  }

  /** A run's properties: where its character formatting differs from its paragraph style's. */
  private runProps(char: MacChar, style: MacChar): RunProps {
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
  private fullRunProps(char: MacChar): RunProps {
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

  /**
   * The picture at file offset `at`: a header (size, count, flags, the
   * frame in points), then `count` parts, each a header of 16 bytes and a
   * PICT. The first part with a PICT is the picture.
   */
  private picture(at: number): Inline | undefined {
    let index = this.pictureIndex.get(at);
    const size = this.u32(at);
    if (size < 14 || at + size > this.bytes.length) return undefined;
    const top = this.i16(at + 6);
    const left = this.i16(at + 8);
    const bottom = this.i16(at + 10);
    const right = this.i16(at + 12);
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
    let width = right - left;
    let height = bottom - top;
    if (width <= 0 || height <= 0) {
      // The PICT's own frame, at 72 dpi.
      const data = this.pictures[index].data;
      height = be16(data, 6) - be16(data, 2);
      width = be16(data, 8) - be16(data, 4);
    }
    return { kind: "picture", index, width: width * 20, height: height * 20 };
  }
}

// ---- Character formatting --------------------------------------------------

const MAC_ROMAN_CHARS: string[] = Array.from({ length: 256 }, (_, i) => MAC_ROMAN.decode(Uint8Array.of(i)));

/** Word's colors by number, 1 black to 8 white (doc/sprm.ts's Ico). */
const ICO: Record<number, string> = { 0: "auto", 1: "000000", 2: "0000FF", 3: "00FFFF", 4: "00FF00", 5: "FF00FF", 6: "FF0000", 7: "FFFF00", 8: "FFFFFF" };

function cloneChar(c: MacChar): MacChar {
  return { ...c, flags: [...c.flags] };
}

/**
 * A CHPX: the first bytes of a character's properties, as far as they
 * differ. Byte 0 flips bold, italic, strikethrough, outline, shadow, small
 * caps, all caps and hidden; byte 1 says which of the others are set; then
 * the font number, the size in half-points, the raise in half-points, the
 * spacing, the color and underline, and for a picture its file offset.
 */
function applyChpx(c: MacChar, d: Uint8Array): void {
  if (d.length === 0) return;
  for (let i = 0; i < 8; i++) if (d[0] & (0x80 >> i)) c.flags[i] = !c.flags[i];
  const what = d.length >= 2 ? d[1] : 0;
  if (d.length >= 4 && be16(d, 2)) c.fontId = be16(d, 2);
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

/** A character sprm among a paragraph's or a piece's. Toggles: 0 off, 1 on, 0x81 the style's value, 0x80 its opposite. */
function applyCharSprm(c: MacChar, { op, arg }: Sprm, style: MacChar): void {
  if ((op >= 0x3c && op <= 0x43) || op === 0x45) {
    const v = arg[0];
    const i = op - 0x3c;
    if (op === 0x45) {
      c.kul = v === 0x81 ? style.kul : v === 0x80 ? (style.kul ? 0 : 1) : v ? 1 : 0;
      return;
    }
    c.flags[i] = v === 0x81 ? style.flags[i] : v === 0x80 ? !style.flags[i] : v !== 0;
    return;
  }
  switch (op) {
    case 0x44: c.fontId = be16(arg, 0); break;
    case 0x4a: if (arg[0]) c.hps = arg[0]; break;
    case 0x4d: c.hpsPos = (arg[0] << 24) >> 24; break;
    case 0x4e:
    case 0x53: applyChpx(c, arg.subarray(1)); break;
    default: break;
  }
}

// ---- Paragraph formatting --------------------------------------------------

function clonePara(p: MacPara): MacPara {
  return { props: structuredClone(p.props), charSprms: [...p.charSprms], borderStyle: p.borderStyle };
}

/** The operand sizes of the sprms Word for the Macintosh writes (Word for Windows 2.0's numbering); -1 for a size byte first, -2 for a size word. */
const SPRM_SIZES: Record<number, number> = {
  0x02: 1, 0x03: -1, 0x04: 1, 0x05: 1, 0x06: 1, 0x07: 1, 0x08: 1, 0x09: 1, 0x0a: 1, 0x0b: 1, 0x0c: 1, 0x0d: 1, 0x0e: 1, 0x0f: -1,
  0x10: 2, 0x11: 2, 0x12: 2, 0x13: 2, 0x14: 2, 0x15: 2, 0x16: 2, 0x17: -1, 0x18: 1, 0x19: 1, 0x1a: 2, 0x1b: 2, 0x1c: 2, 0x1d: 1,
  0x1e: 2, 0x1f: 2, 0x20: 2, 0x21: 2, 0x22: 2, 0x23: 2, 0x24: 2,
  0x34: 1, 0x38: 3, 0x39: 1, 0x3a: 0, 0x3c: 1, 0x3d: 1, 0x3e: 1, 0x3f: 1, 0x40: 1, 0x41: 1, 0x42: 1, 0x43: 1, 0x44: 2, 0x45: 1,
  0x47: 1, 0x49: 1, 0x4a: 1, 0x4c: 1, 0x4d: 1, 0x4e: -1, 0x4f: 3, 0x50: 3, 0x53: -1, 0x5e: 1, 0x5f: -1,
  0x75: 1, 0x76: 1, 0x77: 2, 0x78: 2, 0x79: 1, 0x7a: 1, 0x7b: 2, 0x7c: 2, 0x7d: 1, 0x7e: 1, 0x7f: 1, 0x80: 1, 0x81: 2, 0x82: 2, 0x83: 2,
  0x84: 2, 0x85: 1, 0x86: 1, 0x87: 2, 0x88: 2,
  0x92: 2, 0x93: 2, 0x94: 2, 0x98: -2, 0x99: 2, 0x9a: -1, 0x9d: 5, 0x9e: 4, 0x9f: 2, 0xa0: 4, 0xa3: 5,
};

/** A grpprl's sprms, up to the first one whose size is unknown. */
function parseSprms(b: Uint8Array): Sprm[] {
  const out: Sprm[] = [];
  let at = 0;
  while (at < b.length) {
    const op = b[at];
    const size = SPRM_SIZES[op];
    if (size === undefined) break;
    let start = at + 1;
    let length = size;
    if (size === -1) {
      length = b[at + 1] ?? 0;
      start = at + 2;
    } else if (size === -2) {
      length = be16(b, at + 1);
      start = at + 3;
    }
    if (start + length > b.length) break;
    // A full font (0x4e, 0x53) keeps its size byte: applyChpx reads after it.
    out.push({ op, arg: op === 0x4e || op === 0x53 ? b.subarray(at + 1, start + length) : b.subarray(start, start + length) });
    at = start + length;
  }
  return out;
}

/** A border word: the line in the low 9 bits, the space to the text in points above them, a shadow bit. */
function macBorder(v: number): Border | undefined {
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

function applyParaSprms(para: MacPara, sprms: Sprm[], reader?: MacReader): void {
  const p = para.props;
  for (const sprm of sprms) {
    const { op, arg } = sprm;
    if ((op >= 0x3c && op <= 0x53) || op === 0x5e) {
      para.charSprms.push(sprm);
      continue;
    }
    if (op >= 0x92) {
      applyTableSprm(p, op, arg);
      continue;
    }
    switch (op) {
      case 0x02: if (reader) p.istd = reader.styleIndex((arg[0] << 24) >> 24); break;
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
      case 0x17: applyTabs(p, arg, op === 0x17 ? 4 : 2); break;
      case 0x10: p.dxaRight = be16s(arg, 0); break;
      case 0x11: p.dxaLeft = be16s(arg, 0); break;
      case 0x13: p.dxaLeft1 = be16s(arg, 0); break;
      case 0x14: {
        // Positive: at least; negative: exactly; 0: single, as the font needs.
        const v = be16s(arg, 0);
        p.lineSpacing = v || 240;
        p.lineMultiple = v === 0;
        break;
      }
      case 0x15: p.dyaBefore = be16s(arg, 0); break;
      case 0x16: p.dyaAfter = be16s(arg, 0); break;
      case 0x18: p.inTable = arg[0] !== 0; break;
      case 0x19: p.ttp = arg[0] !== 0; break;
      case 0x1a: p.frame.x = be16(arg, 0); break;
      case 0x1b: p.frame.y = be16(arg, 0); break;
      case 0x1c: p.frame.width = be16(arg, 0); break;
      case 0x1d: p.frame.pcHorz = arg[0] & 3; p.frame.pcVert = arg[0] & 4 ? 1 : 0; break;
      case 0x24: p.frame.dxaFromText = be16(arg, 0); break;
      case 0x1e: p.borders.top = macBorder(be16(arg, 0)); break;
      case 0x1f: p.borders.left = macBorder(be16(arg, 0)); break;
      case 0x20: p.borders.bottom = macBorder(be16(arg, 0)); break;
      case 0x21: p.borders.right = macBorder(be16(arg, 0)); break;
      case 0x22: p.borders.between = macBorder(be16(arg, 0)); break;
      default: break;
    }
  }
}

/** A tab change: the positions cleared (each `deletedSize` bytes), then the positions set and a byte of alignment and leader per tab. */
function applyTabs(p: ParaProps, arg: Uint8Array, deletedSize: number): void {
  let at = 0;
  const deleted = arg[at++] ?? 0;
  for (let i = 0; i < deleted; i++) {
    const pos = be16s(arg, at + i * deletedSize);
    p.tabs = p.tabs.filter((t) => Math.abs(t.pos - pos) > 1);
  }
  at += deleted * deletedSize;
  const added = arg[at++] ?? 0;
  for (let i = 0; i < added && at + 2 * added + i < arg.length; i++) {
    const pos = be16s(arg, at + 2 * i);
    const type = arg[at + 2 * added + i];
    p.tabs = p.tabs.filter((t) => Math.abs(t.pos - pos) > 1);
    p.tabs.push({ pos, jc: Math.min(type >> 5, 4), leader: (type >> 2) & 3 });
  }
}

function applyTableSprm(p: ParaProps, op: number, arg: Uint8Array): void {
  const t = (p.table ??= newTableProps());
  switch (op) {
    case 0x92: t.jc = be16(arg, 0) & 3; break;
    case 0x93: t.dxaLeft = be16s(arg, 0); break;
    case 0x94: t.dxaGapHalf = be16s(arg, 0); break;
    case 0x99: t.dyaRowHeight = be16s(arg, 0); break;
    case 0x98: {
      // The cell count, the cell edges, then per cell a flag word and four borders (top, left, bottom, right).
      const count = arg[0] ?? 0;
      t.left = be16s(arg, 1);
      t.cells = [];
      const cellsAt = 1 + 2 * (count + 1);
      for (let i = 0; i < count; i++) {
        const cell: CellProps = {
          right: be16s(arg, 1 + 2 * (i + 1)), horzMerge: 0, vertMerge: 0, vertAlign: 0, textFlow: 0, ftsWidth: 0, width: 0,
          fitText: false, noWrap: false, hideMark: false, borders: {}, margins: {},
        };
        const at = cellsAt + 10 * i;
        if (at + 10 <= arg.length) {
          cell.borders.top = macBorder(be16(arg, at + 2));
          cell.borders.left = macBorder(be16(arg, at + 4));
          cell.borders.bottom = macBorder(be16(arg, at + 6));
          cell.borders.right = macBorder(be16(arg, at + 8));
        }
        t.cells.push(cell);
      }
      break;
    }
    case 0x9d: {
      // Shading for the cells [first, last): a percentage of black in hundredths.
      const fill = grey(be16(arg, 3));
      for (let i = arg[1]; i < arg[2] && i < t.cells.length; i++) t.cells[i].shading = { pattern: 0, fore: "auto", fill };
      break;
    }
    case 0xa3: {
      const border = macBorder(be16(arg, 3));
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

// ---- Bytes ----------------------------------------------------------------

function be16(b: Uint8Array, at: number): number {
  return ((b[at] ?? 0) << 8) | (b[at + 1] ?? 0);
}

function be16s(b: Uint8Array, at: number): number {
  return (be16(b, at) << 16) >> 16;
}

function be32(b: Uint8Array, at: number): number {
  return (((b[at] ?? 0) << 24) | ((b[at + 1] ?? 0) << 16) | ((b[at + 2] ?? 0) << 8) | (b[at + 3] ?? 0)) >>> 0;
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
