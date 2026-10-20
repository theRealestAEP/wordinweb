// A Word 97-2003 .doc as a .docx package: the text of every document part
// with its direct formatting, the styles, lists, sections, headers and
// footers, footnotes and endnotes, tables, fields and bookmarks, written as
// WordprocessingML for the same parser, layout and compare a .docx gets.
import { zipSync } from "fflate";
import { FIXED_ZIP_MTIME } from "../zip.js";
import { parsePlc, parseSttb, u16, u32 } from "./bytes.js";
import { FC, Fib, PartStarts, parseFib, partStarts, tableBlob } from "./fib.js";
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
  const pieces = new PieceTable(tableBlob(fib, streams.table, FC.clx), streams.wordDocument);
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
   * the run's rPr body. Undefined drops the character.
   */
  special?: (ch: number, cp: number, props: RunProps, rPr: string) => string | undefined;
  /** A finished run's XML, wrapped (w:ins, w:del) as its properties say. */
  wrapRun?: (cp: number, end: number, props: RunProps, runXml: string) => string;
  /** Extra package parts (comments.xml, media), with their relationships from document.xml and their content types. */
  parts?: { files: Record<string, Uint8Array>; documentRels: { id: string; type: string; target: string }[]; contentTypes: { partName: string; contentType: string }[] };
}

/** The .docx package's parts for a .doc. */
export function docToDocxFiles(bytes: Uint8Array, hooks: ConvertHooks = {}): Record<string, Uint8Array> {
  const reader = openDoc(bytes);
  return new Converter(reader, hooks).files();
}

/** A .doc as a .docx. */
export function docToDocx(bytes: Uint8Array, hooks: ConvertHooks = {}): Uint8Array {
  return zipSync(docToDocxFiles(bytes, hooks), { mtime: FIXED_ZIP_MTIME });
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

  constructor(private readonly reader: DocReader, private readonly hooks: ConvertHooks) {
    // The whole text once: every story is a slice of it.
    this.text = reader.pieces.text(0, Math.min(reader.pieces.end, reader.parts.end));
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
    const extra = this.hooks.parts;
    if (extra) {
      Object.assign(this.parts, extra.files);
      this.documentRels.push(...extra.documentRels);
      this.contentTypes.push(...extra.contentTypes);
    }
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
        applySectSprms(word.subarray(fcSepx + 2, fcSepx + 2 + cb), props);
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
    const names = parseSttb(tableBlob(fib, streams.table, FC.sttbfBkmk)).strings;
    const starts = parsePlc(tableBlob(fib, streams.table, FC.plcfBkf), 4);
    const ends = parsePlc(tableBlob(fib, streams.table, FC.plcfBkl), 0);
    for (let i = 0; i < starts.data.length && i < names.length; i++) {
      const ibkl = u16(starts.data[i], 0);
      if (ibkl >= ends.cps.length - 1) continue;
      this.addMarker(starts.cps[i], 1, `<w:bookmarkStart w:id="${i}" w:name="${escapeAttr(names[i])}"/>`);
      this.addMarker(ends.cps[ibkl], 0, `<w:bookmarkEnd w:id="${i}"/>`);
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
    for (let cp = start; cp < end; cp++) {
      const ch = this.text.charCodeAt(cp);
      const isMark = ch === 0x0d || ch === 0x07 || (ch === 0x0c && sectionEnds?.has(cp));
      if (!isMark && cp !== end - 1) continue;
      const mark = cp;
      const direct = this.reader.props.paragraphProps(mark);
      const props = newParaProps(direct.istd);
      applyParaSprms(direct.grpprl, props);
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
      const trPr = trPrXml(t, gridBefore, gridAfter);
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
    const markRPr = rPrXml(markProps, ctx, p.props.istd);
    let pPr = ppr.before;
    if (markRPr) pPr += `<w:rPr>${markRPr}</w:rPr>`;
    if (sectPr) pPr += sectPr;
    const out: string[] = [pPr ? `<w:p><w:pPr>${pPr}</w:pPr>` : "<w:p>"];
    const field = { inCode: [] as boolean[] };
    for (const run of runs) {
      const end = Math.min(run.end, p.mark);
      if (run.cp >= end && run.cp !== p.mark) continue;
      if (run.cp >= p.mark) break;
      out.push(this.runsXml(run, end, p.props.istd, field, noteRef));
    }
    out.push(this.markersAt(p.mark));
    return out.join("") + "</w:p>";
  }

  /** The runs for the text of one character run, split at special characters and markers. */
  private runsXml(run: CharRun, end: number, istdPara: number, field: { inCode: boolean[] }, noteRef: NoteRef | null): string {
    const { ctx, fonts } = this.reader;
    const props = newRunProps();
    applyRunSprms(run.grpprl, props);
    const key = `${istdPara}:${latin1(run.grpprl)}`;
    let rPr = this.rPrCache.get(key);
    if (rPr === undefined) {
      rPr = rPrXml(props, ctx, istdPara);
      this.rPrCache.set(key, rPr);
    }
    const rPrXmlWrapped = rPr ? `<w:rPr>${rPr}</w:rPr>` : "";
    const special = toggleValue(props, "fSpec", ctx, istdPara) === true;
    const symbolFont = run.compressed && props.ftc0 !== undefined && fonts[props.ftc0]?.charset === 2;
    const out: string[] = [];
    let text = "";
    let content = "";
    const flushText = (): void => {
      if (!text) return;
      const inCode = field.inCode.length > 0 && field.inCode[field.inCode.length - 1];
      content += inCode ? `<w:instrText xml:space="preserve">${escapeText(text)}</w:instrText>` : `<w:t xml:space="preserve">${escapeText(text)}</w:t>`;
      text = "";
    };
    const flushRun = (cpEnd: number): void => {
      flushText();
      if (!content) return;
      const xml = `<w:r>${rPrXmlWrapped}${content}</w:r>`;
      out.push(this.hooks.wrapRun ? this.hooks.wrapRun(run.cp, cpEnd, props, xml) : xml);
      content = "";
    };
    for (let cp = run.cp; cp < end; cp++) {
      const markers = this.markersAt(cp);
      if (markers) {
        flushRun(cp);
        out.push(markers);
      }
      let ch = this.text.charCodeAt(cp);
      if (ch >= 0x20) {
        if (ch === 0x28 && special && props.symbol) {
          flushText();
          const font = fonts[props.symbol.ftc]?.name ?? "Symbol";
          content += `<w:sym w:font="${escapeAttr(font)}" w:char="${props.symbol.char.toString(16).toUpperCase().padStart(4, "0")}"/>`;
          continue;
        }
        if (special && (ch === 0x3c || ch === 0x3e)) continue;
        if (symbolFont && ch < 0x100) ch += 0xf000;
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
        case 0x13: flushText(); content += '<w:fldChar w:fldCharType="begin"/>'; field.inCode.push(true); break;
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
          const xml = this.hooks.special?.(ch, cp, props, rPr);
          if (xml) out.push(xml);
          break;
        }
        default:
          break;
      }
    }
    flushRun(end);
    return out.join("");
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

  /** The headerReference and footerReference elements of section `index`, registering the parts. */
  private headerReferences(index: number): string {
    const stories = this.headerStories;
    const base = this.reader.parts.headers;
    const kinds = [
      ["header", "even"], ["header", "default"], ["footer", "even"], ["footer", "default"], ["header", "first"], ["footer", "first"],
    ] as const;
    let out = "";
    for (let k = 0; k < 6; k++) {
      const i = 6 + index * 6 + k;
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

function latin1(bytes: Uint8Array): string {
  let s = "";
  for (let i = 0; i < bytes.length; i++) s += String.fromCharCode(bytes[i]);
  return s;
}
