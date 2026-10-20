// A Word 97-2003 .doc as a .docx package: the text of every document part
// with its direct formatting, the styles, lists, sections, headers and
// footers, footnotes and endnotes, tables, fields and bookmarks, written as
// WordprocessingML for the same parser, layout and compare a .docx gets.
import { parsePlc, parseSttb, u16, u32 } from "./bytes.js";
import { FC, Fib, PartStarts, parseFib, partStarts, tableBlob } from "./fib.js";
import { ffDataXml } from "./formfields.js";
import { numberingXml } from "./numbering.js";
import { CharRun, PropertyTables } from "./props.js";
import {
  ParaProps,
  RunProps,
  SectionProps,
  TableProps,
  XmlContext,
  applyParaSprms,
  applyRunSprms,
  applySectSprms,
  escapeAttr,
  escapeText,
  newParaProps,
  newRunProps,
  newSectionProps,
  notePrXml,
  pPrXml,
  rPrXml,
  sectPrXml,
  tblPrXml,
  tcPrXml,
  toggleValue,
  trPrXml,
} from "./sprm.js";
import { DocStreams, openDocStreams } from "./streams.js";
import { Font, StyleSheet, parseFonts } from "./styles.js";
import { PieceTable } from "./text.js";
import { Word6Lists } from "./numbering6.js";
import { highHalf, parseSttb6, word6Grpprl } from "./word6.js";

/** The parsed structures of a .doc, for every module that reads one. */
export interface DocReader {
  streams: DocStreams;
  fib: Fib;
  /** Where each document part begins in the one CP space. */
  parts: PartStarts;
  pieces: PieceTable;
  props: PropertyTables;
  fonts: Font[];
  styles: StyleSheet;
  /** Names the XML writers need: style ids, font names, toggles of the style stack. */
  ctx: XmlContext;
}

export function openDoc(bytes: Uint8Array): DocReader {
  const streams = openDocStreams(bytes);
  const fib = parseFib(streams.wordDocument);
  const pieces = new PieceTable(tableBlob(fib, streams.table, FC.clx), streams.wordDocument, fib.word6);
  const props = new PropertyTables(streams.wordDocument, streams.table, streams.data, fib, pieces);
  const fonts = parseFonts(fib, streams.table);
  const styles = new StyleSheet(fib, streams.table);
  const ctx: XmlContext = {
    styleId: (istd) => styles.id(istd),
    fontName: (ftc) => fonts[ftc]?.name,
    styleToggle: (prop, istdPara, istdChar) => styles.toggle(prop, istdPara, istdChar),
  };
  return { streams, fib, parts: partStarts(fib), pieces, props, fonts, styles, ctx };
}

/**
 * What the other readers of a .doc add to the conversion: review marks and
 * drawings. Every hook is optional; without it the construct is left out.
 */
export interface ConvertHooks {
  /** XML to put before the character at `cp` (a comment range mark); range ends come before starts at one CP. */
  markers?: (cp: number) => string | undefined;
  /**
   * The run XML for a special character the text holds: U+0001 (a
   * picture; props.picLocation is its offset in the Data stream), U+0005
   * (a comment reference) or U+0008 (a floating shape's anchor). `rPr` is
   * the run's rPr body. `paragraphs` writes the text at [cpStart, cpEnd)
   * as formatted w:p elements (a text box's text). Undefined drops the
   * character.
   */
  special?: (ch: number, cp: number, props: RunProps, rPr: string, paragraphs: (cpStart: number, cpEnd: number) => string) => string | undefined;
  /** Extra XML at the end of a run's rPr (an rPrChange) for the text at [cp, end). */
  runExtra?: (cp: number, end: number, props: RunProps) => string | undefined;
  /** A finished run's XML (a special character's run too), wrapped (w:ins, w:del) as its properties say. */
  wrapRun?: (cp: number, end: number, props: RunProps, runXml: string) => string;
  /**
   * Extra XML for a paragraph mark at `cp`: `rPr` goes at the end of the
   * mark's rPr (w:ins, w:del), `pPr` at the end of the pPr (a pPrChange).
   */
  markExtra?: (cp: number, props: ParaProps, mark: RunProps) => { rPr?: string; pPr?: string } | undefined;
  /** Extra XML at the end of the trPr of the row whose row mark is at `cp` (w:ins, w:del, trPrChange). */
  rowExtra?: (cp: number, table: TableProps) => string | undefined;
  /** A finished paragraph's XML, rewritten (adjacent w:ins/w:del of one change joined). */
  paragraph?: (xml: string) => string;
  /**
   * Extra package parts (comments.xml, media), with their relationships
   * from document.xml and their content types. `paragraphs` writes the text
   * at [cpStart, cpEnd) as formatted w:p elements (a comment's text).
   */
  parts?: (paragraphs: (cpStart: number, cpEnd: number) => string) => ExtraParts;
}

export interface ExtraParts {
  files: Record<string, Uint8Array>;
  documentRels: { id: string; type: string; target: string }[];
  contentTypes: { partName: string; contentType: string }[];
}

/** The .docx package's parts for an opened .doc, with the given hooks. */
export function convertDoc(reader: DocReader, hooks: ConvertHooks): Record<string, Uint8Array> {
  return new Converter(reader, hooks).files();
}

/** The .docx package's parts for a .doc, with only the given hooks (doc/hooks.ts adds every reader's). */
export function docToDocxFiles(bytes: Uint8Array, hooks: ConvertHooks = {}): Record<string, Uint8Array> {
  return convertDoc(openDoc(bytes), hooks);
}

const W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const XML_HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

interface Paragraph {
  start: number;
  /** The CP of the paragraph mark (or cell, row or section mark). */
  mark: number;
  props: ParaProps;
  depth: number;
  cellEnd: boolean;
  rowEnd: boolean;
}

interface Row {
  cells: Paragraph[][];
  table: TableProps;
  mark: Paragraph;
}

interface NoteRef {
  kind: "footnote" | "endnote";
  id: number;
  custom: boolean;
}

/** Where a paragraph's writer is among its fields: whether each open field is in its code, and the CP before which text was already written (a ruby). */
interface FieldState {
  inCode: boolean[];
  skipTo: number;
}

interface Section {
  start: number;
  end: number;
  props: SectionProps;
}

class Converter {
  private readonly text: string;
  private readonly noteRefs = new Map<number, NoteRef>();
  private readonly markers = new Map<number, { order: number; xml: string }[]>();
  private readonly sections: Section[] = [];
  private readonly headerStories: number[] = [];
  private readonly parts: Record<string, Uint8Array> = {};
  private readonly documentRels: { id: string; type: string; target: string }[] = [];
  private readonly contentTypes: { partName: string; contentType: string }[] = [];
  private nextRel = 1;
  private headerCount = 0;
  private footerCount = 0;
  private readonly rPrCache = new Map<string, string>();
  private readonly encoder = new TextEncoder();
  /** A Word 6.0/95 document's lists, which its paragraphs make as they are read. */
  private readonly lists6: Word6Lists | undefined;
  private readonly styleAutonumbers = new Map<number, { nLvlAnm?: number; anld?: Uint8Array }>();

  constructor(private readonly reader: DocReader, private readonly hooks: ConvertHooks) {
    // The whole text once: every story is a slice of it.
    this.text = reader.pieces.text(0, Math.min(reader.pieces.end, reader.parts.end));
    if (reader.fib.word6) this.lists6 = new Word6Lists(reader.fonts, reader.fib.word6.mac, reader.ctx);
    this.readSections();
    this.readHeaderStories();
    this.readNoteRefs();
    this.readBookmarks();
  }

  files(): Record<string, Uint8Array> {
    const { reader } = this;
    const styles = reader.styles.xml(reader.ctx, reader.fonts);
    this.addPart("word/styles.xml", styles, "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml", `${REL_NS}/styles`);
    const numbering = numberingXml(reader.fib, reader.streams.table, reader.ctx);
    if (numbering) this.addPart("word/numbering.xml", numbering, "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml", `${REL_NS}/numbering`);
    const footnotes = this.notesXml("footnote");
    if (footnotes) this.addPart("word/footnotes.xml", footnotes, "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml", `${REL_NS}/footnotes`);
    const endnotes = this.notesXml("endnote");
    if (endnotes) this.addPart("word/endnotes.xml", endnotes, "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml", `${REL_NS}/endnotes`);
    this.addPart("word/settings.xml", this.settingsXml(), "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml", `${REL_NS}/settings`);
    // The body last: its sections register their header and footer parts.
    const body = this.bodyXml();
    const extra = this.hooks.parts?.((cpStart, cpEnd) => this.storyXml(cpStart, cpEnd, null, null, null));
    if (extra) {
      Object.assign(this.parts, extra.files);
      this.documentRels.push(...extra.documentRels);
      this.contentTypes.push(...extra.contentTypes);
    }
    const lists6 = this.lists6?.xml();
    if (lists6) this.addPart("word/numbering.xml", lists6, "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml", `${REL_NS}/numbering`);
    this.parts["word/document.xml"] = this.encoder.encode(`${XML_HEAD}<w:document ${W_NS}><w:body>${body}</w:body></w:document>`);
    this.parts["word/_rels/document.xml.rels"] = this.encoder.encode(
      `${XML_HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${this.documentRels.map((rel) => `<Relationship Id="${rel.id}" Type="${rel.type}" Target="${escapeAttr(rel.target)}"/>`).join("")}</Relationships>`,
    );
    this.parts["_rels/.rels"] = this.encoder.encode(
      `${XML_HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL_NS}/officeDocument" Target="word/document.xml"/></Relationships>`,
    );
    const overrides = this.contentTypes.map((c) => `<Override PartName="${escapeAttr(c.partName)}" ContentType="${c.contentType}"/>`).join("");
    this.parts["[Content_Types].xml"] = this.encoder.encode(
      `${XML_HEAD}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Default Extension="jpeg" ContentType="image/jpeg"/><Default Extension="jpg" ContentType="image/jpeg"/><Default Extension="gif" ContentType="image/gif"/><Default Extension="bmp" ContentType="image/bmp"/><Default Extension="tiff" ContentType="image/tiff"/><Default Extension="emf" ContentType="image/x-emf"/><Default Extension="wmf" ContentType="image/x-wmf"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>${overrides}</Types>`,
    );
    return this.parts;
  }

  private addPart(name: string, xml: string, contentType: string, relType: string): string {
    this.parts[name] = this.encoder.encode(xml);
    this.contentTypes.push({ partName: "/" + name, contentType });
    const id = `rId${this.nextRel++}`;
    this.documentRels.push({ id, type: relType, target: name.slice("word/".length) });
    return id;
  }

  // ---- The document's tables of positions ------------------------------

  private readSections(): void {
    const { fib, streams } = this.reader;
    const plc = parsePlc(tableBlob(fib, streams.table, FC.plcfSed), 12);
    const word = streams.wordDocument;
    for (let i = 0; i < plc.data.length; i++) {
      const props = newSectionProps();
      const fcSepx = u32(plc.data[i], 2);
      if (fcSepx !== 0xffffffff && fcSepx + 2 <= word.length) {
        const cb = u16(word, fcSepx);
        const grpprl = word.subarray(fcSepx + 2, fcSepx + 2 + cb);
        applySectSprms(fib.word6 ? word6Grpprl(grpprl) : grpprl, props);
      }
      this.sections.push({ start: plc.cps[i], end: Math.min(plc.cps[i + 1], fib.ccpText), props });
    }
    if (this.sections.length === 0) this.sections.push({ start: 0, end: fib.ccpText, props: newSectionProps() });
  }

  private readHeaderStories(): void {
    const { fib, streams } = this.reader;
    const blob = tableBlob(fib, streams.table, FC.plcfHdd);
    for (let at = 0; at + 4 <= blob.length; at += 4) this.headerStories.push(u32(blob, at));
  }

  private readNoteRefs(): void {
    const { fib, streams } = this.reader;
    const refs = parsePlc(tableBlob(fib, streams.table, FC.plcffndRef), 2);
    for (let i = 0; i < refs.data.length; i++) this.noteRefs.set(refs.cps[i], { kind: "footnote", id: i + 1, custom: u16(refs.data[i], 0) === 0 });
    const endRefs = parsePlc(tableBlob(fib, streams.table, FC.plcfendRef), 2);
    for (let i = 0; i < endRefs.data.length; i++) this.noteRefs.set(endRefs.cps[i], { kind: "endnote", id: i + 1, custom: u16(endRefs.data[i], 0) === 0 });
  }

  private readBookmarks(): void {
    const { fib, streams } = this.reader;
    const sttb = tableBlob(fib, streams.table, FC.sttbfBkmk);
    const names = (fib.word6 ? parseSttb6(sttb, 0) : parseSttb(sttb)).strings;
    const starts = parsePlc(tableBlob(fib, streams.table, FC.plcfBkf), 4);
    const ends = parsePlc(tableBlob(fib, streams.table, FC.plcfBkl), 0);
    for (let i = 0; i < starts.data.length && i < names.length; i++) {
      const ibkl = u16(starts.data[i], 0);
      if (ibkl >= ends.cps.length - 1) continue;
      const start = `<w:bookmarkStart w:id="${i}" w:name="${escapeAttr(names[i])}"/>`;
      const end = `<w:bookmarkEnd w:id="${i}"/>`;
      // Another bookmark's end comes before a start at the same CP; an empty bookmark's own end comes after it.
      if (ends.cps[ibkl] === starts.cps[i]) {
        this.addMarker(starts.cps[i], 1, start + end);
      } else {
        this.addMarker(starts.cps[i], 1, start);
        this.addMarker(ends.cps[ibkl], 0, end);
      }
    }
  }

  private addMarker(cp: number, order: number, xml: string): void {
    let list = this.markers.get(cp);
    if (!list) this.markers.set(cp, (list = []));
    list.push({ order, xml });
  }

  private markersAt(cp: number): string {
    const own = this.markers.get(cp);
    const hooked = this.hooks.markers?.(cp);
    if (!own && !hooked) return "";
    let out = "";
    if (own) {
      for (const m of [...own].sort((a, b) => a.order - b.order)) out += m.xml;
    }
    return out + (hooked ?? "");
  }

  // ---- Stories ----------------------------------------------------------

  /** The paragraphs of the text at [start, end); `sectionEnds` marks the CPs whose 0x0C is a section mark. */
  private paragraphs(start: number, end: number, sectionEnds: Set<number> | null): Paragraph[] {
    const out: Paragraph[] = [];
    let pStart = start;
    this.lists6?.startStory();
    for (let cp = start; cp < end; cp++) {
      const ch = this.text.charCodeAt(cp);
      const isMark = ch === 0x0d || ch === 0x07 || (ch === 0x0c && sectionEnds?.has(cp));
      if (!isMark && cp !== end - 1) continue;
      const mark = cp;
      const direct = this.reader.props.paragraphProps(mark);
      const props = newParaProps(direct.istd);
      applyParaSprms(direct.grpprl, props);
      if (this.lists6) {
        const style = this.styleAutonumber(props.istd);
        const section = this.sections.find((s) => mark < s.end) ?? this.sections[0];
        this.lists6.number(props, props.nLvlAnm ?? style.nLvlAnm ?? 0, props.anld ?? style.anld, section.props.olst);
      }
      let depth = props.itap ?? (props.inTable ? 1 : 0);
      if (props.dtap) depth += props.dtap;
      if (depth < 0) depth = 0;
      const cellChar = ch === 0x07;
      const rowEnd = depth > 0 && (depth === 1 ? cellChar && !!props.ttp : ch === 0x0d && !!props.innerTtp);
      const cellEnd = depth > 0 && !rowEnd && (depth === 1 ? cellChar : ch === 0x0d && !!props.innerTableCell);
      out.push({ start: pStart, mark, props, depth, cellEnd, rowEnd });
      pStart = cp + 1;
    }
    return out;
  }

  /** A Word 6 paragraph style's autonumbering, through the styles it is based on. */
  private styleAutonumber(istd: number): { nLvlAnm?: number; anld?: Uint8Array } {
    let found = this.styleAutonumbers.get(istd);
    if (!found) {
      const chain: Uint8Array[] = [];
      for (let i = istd, seen = 0; i !== 0x0fff && seen < 32; seen++) {
        const style = this.reader.styles.style(i);
        if (!style) break;
        if (style.papx) chain.unshift(style.papx);
        i = style.base;
      }
      const props = newParaProps(istd);
      for (const papx of chain) applyParaSprms(papx, props);
      found = { nLvlAnm: props.nLvlAnm, anld: props.anld };
      this.styleAutonumbers.set(istd, found);
    }
    return found;
  }

  /** The blocks of a story as XML: its paragraphs and tables. */
  private storyXml(start: number, end: number, sectionEnds: Set<number> | null, sectPrAt: Map<number, string> | null, noteRef: NoteRef | null): string {
    if (end <= start) return "<w:p/>";
    const paras = this.paragraphs(start, end, sectionEnds);
    if (paras.length === 0) return "<w:p/>";
    return this.blocksXml(paras, 0, sectPrAt, noteRef);
  }

  private blocksXml(paras: Paragraph[], depth: number, sectPrAt: Map<number, string> | null, noteRef: NoteRef | null): string {
    const out: string[] = [];
    let i = 0;
    while (i < paras.length) {
      if (paras[i].depth <= depth) {
        out.push(this.paragraphXml(paras[i], sectPrAt?.get(paras[i].mark), noteRef && i === 0 ? noteRef : null));
        i++;
        continue;
      }
      const rows: Row[] = [];
      let cells: Paragraph[][] = [];
      let cellStart = i;
      let j = i;
      while (j < paras.length && paras[j].depth > depth) {
        const p = paras[j];
        if (p.depth === depth + 1 && p.rowEnd) {
          if (cellStart < j) cells.push(paras.slice(cellStart, j));
          rows.push({ cells, table: p.props.table ?? { cells: [], borders: {}, cellMargins: {} }, mark: p });
          cells = [];
          cellStart = j + 1;
        } else if (p.depth === depth + 1 && p.cellEnd) {
          cells.push(paras.slice(cellStart, j + 1));
          cellStart = j + 1;
        }
        j++;
      }
      if (rows.length) out.push(this.tableXml(rows, depth + 1, sectPrAt, noteRef));
      // Paragraphs after the last row mark (a damaged table): as plain paragraphs.
      for (let k = cellStart; k < j; k++) out.push(this.paragraphXml(paras[k], undefined, null));
      i = j;
    }
    return out.join("");
  }

  private tableXml(rows: Row[], depth: number, sectPrAt: Map<number, string> | null, noteRef: NoteRef | null): string {
    // The grid: every cell edge of every row, merged within two twips.
    const edges: number[] = [];
    for (const row of rows) {
      edges.push(row.table.left ?? 0);
      for (const cell of row.table.cells) edges.push(cell.right);
    }
    edges.sort((a, b) => a - b);
    const grid: number[] = [];
    for (const edge of edges) if (grid.length === 0 || edge - grid[grid.length - 1] > 2) grid.push(edge);
    const column = (edge: number): number => {
      let best = 0;
      for (let k = 1; k < grid.length; k++) if (Math.abs(grid[k] - edge) < Math.abs(grid[best] - edge)) best = k;
      return best;
    };
    const first = rows[0].table;
    const gapHalf = first.dxaGapHalf ?? 0;
    const tblInd = grid[0] + gapHalf;
    const out: string[] = [`<w:tbl><w:tblPr>${tblPrXml(first, this.reader.ctx, tblInd)}</w:tblPr>`];
    out.push(`<w:tblGrid>${grid.slice(1).map((edge, k) => `<w:gridCol w:w="${edge - grid[k]}"/>`).join("")}</w:tblGrid>`);
    for (const row of rows) {
      const t = row.table;
      const left = t.left ?? grid[0];
      const gridBefore = column(left);
      const lastRight = t.cells.length ? t.cells[t.cells.length - 1].right : left;
      const gridAfter = grid.length - 1 - column(lastRight);
      let tr = `<w:tr>`;
      const trPr = trPrXml(t, gridBefore, gridAfter) + (this.hooks.rowExtra?.(row.mark.mark, t) ?? "");
      if (trPr) tr += `<w:trPr>${trPr}</w:trPr>`;
      let prevEdge = left;
      let cellIndex = 0;
      for (let c = 0; c < t.cells.length; c++) {
        const cell = t.cells[c];
        // A horizontally merged cell joins the one before it.
        if (cell.horzMerge === 1 && c > 0) {
          continue;
        }
        let right = cell.right;
        let k = c + 1;
        while (k < t.cells.length && t.cells[k].horzMerge === 1) right = t.cells[k++].right;
        const span = Math.max(1, column(right) - column(prevEdge));
        const blocks = row.cells[cellIndex];
        cellIndex++;
        const content = blocks && blocks.length ? this.blocksXml(blocks, depth, sectPrAt, noteRef) : "<w:p/>";
        tr += `<w:tc><w:tcPr>${tcPrXml(cell, right - prevEdge, span, !!first.autofit)}</w:tcPr>${content}</w:tc>`;
        prevEdge = right;
      }
      // Cell content with no cell definition (a damaged row) still shows.
      for (; cellIndex < row.cells.length; cellIndex++) {
        tr += `<w:tc><w:tcPr><w:tcW w:w="0" w:type="auto"/></w:tcPr>${this.blocksXml(row.cells[cellIndex], depth, sectPrAt, noteRef)}</w:tc>`;
      }
      if (cellIndex === 0 && t.cells.length === 0) tr += "<w:tc><w:p/></w:tc>";
      out.push(tr + "</w:tr>");
    }
    return out.join("") + "</w:tbl>";
  }

  private paragraphXml(p: Paragraph, sectPr: string | undefined, noteRef: NoteRef | null): string {
    const { ctx } = this.reader;
    const runs = this.reader.props.charRuns(p.start, p.mark + 1);
    // The mark's own properties: the paragraph's rPr.
    const markRun = runs.length ? runs[runs.length - 1] : null;
    const markProps = newRunProps();
    if (markRun) applyRunSprms(markRun.grpprl, markProps);
    const ppr = pPrXml(p.props, ctx);
    const extra = this.hooks.markExtra?.(p.mark, p.props, markProps);
    const markRPr = rPrXml(markProps, ctx, p.props.istd) + (extra?.rPr ?? "");
    let pPr = ppr.before;
    if (markRPr) pPr += `<w:rPr>${markRPr}</w:rPr>`;
    if (sectPr) pPr += sectPr;
    if (extra?.pPr) pPr += extra.pPr;
    const out: string[] = [pPr ? `<w:p><w:pPr>${pPr}</w:pPr>` : "<w:p>"];
    const field: FieldState = { inCode: [], skipTo: 0 };
    for (const run of runs) {
      const end = Math.min(run.end, p.mark);
      if (run.cp >= end && run.cp !== p.mark) continue;
      if (run.cp >= p.mark) break;
      out.push(this.runsXml(run, end, p.props.istd, field, noteRef));
    }
    out.push(this.markersAt(p.mark));
    const xml = out.join("") + "</w:p>";
    return this.hooks.paragraph ? this.hooks.paragraph(xml) : xml;
  }

  /** The runs for the text of one character run, split at special characters and markers. */
  private runsXml(run: CharRun, end: number, istdPara: number, field: FieldState, noteRef: NoteRef | null): string {
    const { ctx, fonts } = this.reader;
    const props = newRunProps();
    applyRunSprms(run.grpprl, props);
    const key = `${istdPara}:${latin1(run.grpprl)}`;
    let rPr = this.rPrCache.get(key);
    if (rPr === undefined) {
      rPr = rPrXml(props, ctx, istdPara);
      this.rPrCache.set(key, rPr);
    }
    const rPrFull = rPr + (this.hooks.runExtra?.(run.cp, end, props) ?? "");
    const rPrXmlWrapped = rPrFull ? `<w:rPr>${rPrFull}</w:rPr>` : "";
    const special = toggleValue(props, "fSpec", ctx, istdPara) === true;
    // Word 6 text is bytes in the character set of the font the run shows in.
    const word6 = this.reader.fib.word6;
    const ftc = props.ftc0 ?? (word6 ? this.reader.styles.font(istdPara, props.istd) : undefined);
    const symbolFont = run.compressed && ftc !== undefined && fonts[ftc]?.charset === 2;
    const highChars = word6 && highHalf(ftc !== undefined ? fonts[ftc]?.charset ?? 0 : 0, props.chse !== undefined ? props.chse === 256 : word6.mac);
    const out: string[] = [];
    let text = "";
    let content = "";
    const flushText = (): void => {
      if (!text) return;
      const inCode = field.inCode.length > 0 && field.inCode[field.inCode.length - 1];
      content += inCode ? `<w:instrText xml:space="preserve">${escapeText(text)}</w:instrText>` : `<w:t xml:space="preserve">${escapeText(text)}</w:t>`;
      text = "";
    };
    let runStart = run.cp;
    const flushRun = (cpEnd: number): void => {
      flushText();
      if (content) {
        const xml = `<w:r>${rPrXmlWrapped}${content}</w:r>`;
        out.push(this.hooks.wrapRun ? this.hooks.wrapRun(runStart, cpEnd, props, xml) : xml);
        content = "";
      }
      runStart = cpEnd;
    };
    for (let cp = Math.max(run.cp, field.skipTo); cp < end; cp++) {
      const markers = this.markersAt(cp);
      if (markers) {
        flushRun(cp);
        out.push(markers);
      }
      let ch = this.text.charCodeAt(cp);
      if (ch >= 0x20) {
        // A custom note mark of ordinary text: the PlcffndRef/PlcfendRef CP makes it a reference.
        const refAt = this.noteRefs.get(cp);
        if (refAt) {
          flushText();
          content += `<w:${refAt.kind}Reference w:customMarkFollows="1" w:id="${refAt.id}"/>`;
        }
        if (ch === 0x28 && special && props.symbol) {
          flushText();
          const font = fonts[props.symbol.ftc]?.name ?? "Symbol";
          content += `<w:sym w:font="${escapeAttr(font)}" w:char="${props.symbol.char.toString(16).toUpperCase().padStart(4, "0")}"/>`;
          continue;
        }
        if (special && (ch === 0x3c || ch === 0x3e)) continue;
        if (symbolFont && ch < 0x100) ch += 0xf000;
        else if (highChars && ch >= 0x80 && ch < 0x100) ch = highChars.charCodeAt(ch - 0x80);
        if (ch === 0x7f) continue;
        text += String.fromCharCode(ch);
        continue;
      }
      switch (ch) {
        case 0x09: flushText(); content += "<w:tab/>"; break;
        case 0x0b: flushText(); content += "<w:br/>"; break;
        case 0x0c: flushText(); content += '<w:br w:type="page"/>'; break;
        case 0x0e: flushText(); content += '<w:br w:type="column"/>'; break;
        case 0x1e: flushText(); content += "<w:noBreakHyphen/>"; break;
        case 0x1f: flushText(); content += "<w:softHyphen/>"; break;
        case 0x13: {
          flushText();
          // A ruby (an EQ \o\ad field), as Word converts it: the field gives way to a w:ruby.
          const ruby = this.rubyXml(cp, istdPara);
          if (ruby) {
            flushRun(cp);
            out.push(ruby.xml);
            // What the field's code held besides text: range marks, and comment references.
            for (let k = cp + 1; k <= ruby.end; k++) {
              out.push(this.markersAt(k));
              if (this.text.charCodeAt(k) === 0x05) out.push(this.hooks.special?.(0x05, k, props, rPrFull, (s, e) => this.storyXml(s, e, null, null, null)) ?? "");
            }
            field.skipTo = ruby.end + 1;
            runStart = ruby.end + 1;
            cp = ruby.end;
            break;
          }
          const ffData = this.formFieldData(cp);
          content += ffData ? `<w:fldChar w:fldCharType="begin">${ffData}</w:fldChar>` : '<w:fldChar w:fldCharType="begin"/>';
          field.inCode.push(true);
          break;
        }
        case 0x14: flushText(); content += '<w:fldChar w:fldCharType="separate"/>'; if (field.inCode.length) field.inCode[field.inCode.length - 1] = false; break;
        case 0x15: flushText(); content += '<w:fldChar w:fldCharType="end"/>'; field.inCode.pop(); break;
        case 0x02: {
          flushText();
          const ref = this.noteRefs.get(cp);
          if (ref) content += `<w:${ref.kind}Reference w:id="${ref.id}"/>`;
          else if (noteRef) content += `<w:${noteRef.kind}Ref/>`;
          break;
        }
        case 0x01:
        case 0x05:
        case 0x08: {
          const refAt = ch !== 0x01 ? this.noteRefs.get(cp) : undefined;
          if (refAt) {
            // A custom note mark drawn with this character.
            flushText();
            content += `<w:${refAt.kind}Reference w:customMarkFollows="1" w:id="${refAt.id}"/>`;
            break;
          }
          flushRun(cp);
          const xml = this.hooks.special?.(ch, cp, props, rPrFull, (s, e) => this.storyXml(s, e, null, null, null));
          if (xml) out.push(this.hooks.wrapRun ? this.hooks.wrapRun(cp, cp + 1, props, xml) : xml);
          runStart = cp + 1;
          break;
        }
        default:
          break;
      }
    }
    flushRun(end);
    return out.join("");
  }

  /**
   * The w:ffData of the field that begins at `cp`: a form field (FORMTEXT,
   * FORMCHECKBOX, FORMDROPDOWN) holds a U+0001 in its code whose run points
   * at the data (sprmCFData, sprmCPicLocation).
   */
  private formFieldData(cp: number): string {
    for (let k = cp + 1; k < this.text.length; k++) {
      const ch = this.text.charCodeAt(k);
      if (ch === 0x13 || ch === 0x14 || ch === 0x15) return "";
      if (ch !== 0x01) continue;
      const [run] = this.reader.props.charRuns(k, k + 1);
      const props = newRunProps();
      if (run) applyRunSprms(run.grpprl, props);
      return props.fData && props.picLocation !== undefined ? ffDataXml(this.reader.streams.data, props.picLocation) : "";
    }
    return "";
  }

  /**
   * A ruby field (EQ \* jc2 \* hps10 \o\ad(\s\up 9(guide),base)) at
   * `cp` as a w:ruby run, and the CP of its end character; null for any
   * other field.
   */
  private rubyXml(cp: number, istdPara: number): { xml: string; end: number } | null {
    let depth = 1;
    let separator = -1;
    let end = -1;
    for (let k = cp + 1; k < this.text.length && end < 0; k++) {
      const ch = this.text.charCodeAt(k);
      if (ch === 0x13) return null;
      if (ch === 0x14 && depth === 1 && separator < 0) separator = k;
      if (ch === 0x15 && --depth === 0) end = k;
    }
    if (end < 0) return null;
    const codeStart = cp + 1;
    const code = this.text.slice(codeStart, separator >= 0 ? separator : end);
    const m = /^\s*EQ\s[\s\S]*?\\o\s*\\a[dclr]\s*\(\s*\\s\s*\\up\s*(\d+)\s*\(([^)]*)\)\s*,([^)]*)\)\s*$/i.exec(code);
    if (!m) return null;
    // \up is in points; hpsRaise, like \* hps, in half-points.
    const [, raise, guide, base] = m;
    const guideAt = codeStart + code.indexOf(guide, code.search(/\\up/i));
    const baseAt = codeStart + code.lastIndexOf(base);
    const propsAt = (at: number): RunProps => {
      const props = newRunProps();
      const [run] = this.reader.props.charRuns(at, at + 1);
      if (run) applyRunSprms(run.grpprl, props);
      return props;
    };
    const baseProps = propsAt(baseAt);
    const guideProps = propsAt(guideAt);
    const hps = Number(/\\\*\s*hps(\d+)/i.exec(code)?.[1] ?? guideProps.hps ?? 10);
    guideProps.hps = hps;
    const align = ["center", "distributeLetter", "distributeSpace", "left", "right", "rightVertical"][Number(/\\\*\s*jc(\d+)/i.exec(code)?.[1] ?? 0)] ?? "center";
    const run = (props: RunProps, text: string): string => {
      const rPr = rPrXml(props, this.reader.ctx, istdPara);
      // A comment reference or other special character inside the field code is left out.
      return `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ""}<w:t xml:space="preserve">${escapeText(text.replace(/[\u0000-\u001f]/g, ""))}</w:t></w:r>`;
    };
    const rubyPr = `<w:rubyAlign w:val="${align}"/><w:hps w:val="${hps}"/><w:hpsRaise w:val="${Number(raise) * 2}"/>${baseProps.hps !== undefined ? `<w:hpsBaseText w:val="${baseProps.hps}"/>` : ""}`;
    return { xml: `<w:r><w:ruby><w:rubyPr>${rubyPr}</w:rubyPr><w:rt>${run(guideProps, guide)}</w:rt><w:rubyBase>${run(baseProps, base)}</w:rubyBase></w:ruby></w:r>`, end };
  }

  // ---- The body, headers and footers, notes ----------------------------

  private bodyXml(): string {
    const { fib } = this.reader;
    const sectionEnds = new Set<number>();
    const sectPrAt = new Map<number, string>();
    let last = "";
    this.sections.forEach((section, i) => {
      const xml = `<w:sectPr>${sectPrXml(section.props, this.headerReferences(i), notePrXml("footnotePr", section.props), notePrXml("endnotePr", section.props))}</w:sectPr>`;
      if (i < this.sections.length - 1 && section.end > 0) {
        sectionEnds.add(section.end - 1);
        sectPrAt.set(section.end - 1, xml);
      } else {
        last = xml;
      }
    });
    const body = this.storyXml(0, fib.ccpText, sectionEnds, sectPrAt, null);
    return body + last;
  }

  /**
   * The header story of kind `k` (in Word 97's order: even header, odd
   * header, even footer, odd footer, first header, first footer) of section
   * `index` in a Word 6 file, which stores only the stories the document
   * (its note separators, DOP grpfIhdt) and each section (SEP grpfIhdt)
   * have, in that order. Undefined when the section has none: it keeps the
   * previous section's.
   */
  private word6HeaderStory(index: number, k: number): number | undefined {
    const dop = tableBlob(this.reader.fib, this.reader.streams.table, FC.dop);
    // The footnote and endnote separators, continuation separators and continuation notices.
    let story = bitCount((dop[1] ?? 0) & 0x3f);
    for (let s = 0; s < index; s++) story += bitCount((this.sections[s].props.grpfIhdt ?? 0) & 0x3f);
    const bits = this.sections[index].props.grpfIhdt ?? 0;
    if (!(bits & (1 << k))) return undefined;
    return story + bitCount(bits & ((1 << k) - 1));
  }

  /** The headerReference and footerReference elements of section `index`, registering the parts. */
  private headerReferences(index: number): string {
    const stories = this.headerStories;
    const base = this.reader.parts.headers;
    const kinds = [
      ["header", "even"], ["header", "default"], ["footer", "even"], ["footer", "default"], ["header", "first"], ["footer", "first"],
    ] as const;
    let out = "";
    for (let k = 0; k < 6; k++) {
      const i = this.reader.fib.word6 ? this.word6HeaderStory(index, k) : 6 + index * 6 + k;
      if (i === undefined) continue;
      if (i + 1 >= stories.length) break;
      const start = stories[i];
      const end = stories[i + 1];
      if (end <= start) continue;
      const [kind, type] = kinds[k];
      const n = kind === "header" ? ++this.headerCount : ++this.footerCount;
      const name = `word/${kind}${n}.xml`;
      // The story's last character is the guard paragraph mark.
      const content = this.storyXml(base + start, base + end - 1, null, null, null);
      const root = kind === "header" ? "hdr" : "ftr";
      const id = this.addPart(name, `${XML_HEAD}<w:${root} ${W_NS}>${content}</w:${root}>`, `application/vnd.openxmlformats-officedocument.wordprocessingml.${kind}+xml`, `${REL_NS}/${kind}`);
      out += `<w:${kind}Reference w:type="${type}" r:id="${id}"/>`;
    }
    return out;
  }

  private notesXml(kind: "footnote" | "endnote"): string | undefined {
    const { fib, streams, parts } = this.reader;
    const base = kind === "footnote" ? parts.footnotes : parts.endnotes;
    const blob = tableBlob(fib, streams.table, kind === "footnote" ? FC.plcffndTxt : FC.plcfendTxt);
    const cps: number[] = [];
    for (let at = 0; at + 4 <= blob.length; at += 4) cps.push(u32(blob, at));
    // The last two CPs end the last note and are a guard.
    const count = Math.max(0, cps.length - 2);
    if (count === 0) return undefined;
    const root = kind === "footnote" ? "footnotes" : "endnotes";
    const out: string[] = [`${XML_HEAD}<w:${root} ${W_NS}>`];
    out.push(`<w:${kind} w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:${kind}>`);
    out.push(`<w:${kind} w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:${kind}>`);
    for (let i = 0; i < count; i++) {
      const refCustom = [...this.noteRefs.values()].some((ref) => ref.kind === kind && ref.id === i + 1 && ref.custom);
      const content = this.storyXml(base + cps[i], base + cps[i + 1], null, null, { kind, id: i + 1, custom: refCustom });
      out.push(`<w:${kind} w:id="${i + 1}">${content}</w:${kind}>`);
    }
    out.push(`</w:${root}>`);
    return out.join("");
  }

  private settingsXml(): string {
    const { fib, streams } = this.reader;
    const dop = tableBlob(fib, streams.table, FC.dop);
    const out: string[] = [`${XML_HEAD}<w:settings ${W_NS}>`];
    if (dop.length > 0 && dop[0] & 0x01) out.push("<w:evenAndOddHeaders/>");
    if (dop.length >= 12) {
      const dxaTab = u16(dop, 10);
      if (dxaTab > 0 && dxaTab !== 720) out.push(`<w:defaultTabStop w:val="${dxaTab}"/>`);
    }
    if (dop.length >= 1) {
      const fpc = (dop[0] >>> 5) & 3;
      if (fpc === 0 || fpc === 2) out.push(`<w:footnotePr><w:pos w:val="${fpc === 0 ? "sectEnd" : "beneathText"}"/></w:footnotePr>`);
    }
    // Word lays a .doc out in Word 2003 compatibility mode.
    out.push('<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="11"/></w:compat>');
    out.push("</w:settings>");
    return out.join("");
  }
}

function bitCount(v: number): number {
  let n = 0;
  for (; v; v &= v - 1) n++;
  return n;
}

function latin1(bytes: Uint8Array): string {
  let s = "";
  for (let i = 0; i < bytes.length; i++) s += String.fromCharCode(bytes[i]);
  return s;
}
