// An OldDocument (model.ts) as the parts of a .docx package, written with
// the .doc reader's property writers (doc/sprm.ts) so the renderer and the
// compare read it as they read a converted Word 97 document.
import { Block, OldDocument, Paragraph, Row, Run, Section, Table } from "./model.js";
import { ParaProps, RunProps, XmlContext, escapeAttr, escapeText, pPrXml, rPrXml, sectPrXml, tblPrXml, tcPrXml, trPrXml } from "../doc/sprm.js";

const W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const DRAWING_NS =
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"';
const REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const XML_HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml";
const PICTURE_TYPES: Record<string, string> = { pict: "image/x-pict", wmf: "image/x-wmf", bmp: "image/bmp" };

/** The .docx package's parts for an old Word document. */
export function oldDocumentFiles(doc: OldDocument): Record<string, Uint8Array> {
  return new PackageWriter(doc).files();
}

class PackageWriter {
  private readonly encoder = new TextEncoder();
  private readonly parts: Record<string, Uint8Array> = {};
  private readonly rels: { id: string; type: string; target: string }[] = [];
  private readonly overrides: { partName: string; contentType: string }[] = [];
  private readonly ctx: XmlContext;
  private readonly noteKind: "footnote" | "endnote";
  private headerCount = 0;
  private footerCount = 0;
  private docPrId = 0;

  constructor(private readonly doc: OldDocument) {
    this.ctx = {
      styleId: (istd) => (istd > 0 && istd < doc.styles.length ? styleId(doc, istd) : undefined),
      fontName: (ftc) => doc.fonts[ftc],
      styleToggle: () => false,
    };
    this.noteKind = doc.endnotes ? "endnote" : "footnote";
  }

  files(): Record<string, Uint8Array> {
    const { doc } = this;
    this.addPart("word/styles.xml", this.stylesXml(), `${CONTENT_TYPE}.styles+xml`, "styles");
    this.addPart("word/settings.xml", this.settingsXml(), `${CONTENT_TYPE}.settings+xml`, "settings");
    if (doc.footnotes.length) {
      const root = `${this.noteKind}s`;
      const notes = doc.footnotes.map((blocks, i) => `<w:${this.noteKind} w:id="${i + 1}">${this.blocksXml(blocks) || "<w:p/>"}</w:${this.noteKind}>`).join("");
      const separators =
        `<w:${this.noteKind} w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:${this.noteKind}>` +
        `<w:${this.noteKind} w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:${this.noteKind}>`;
      this.addPart(`word/${root}.xml`, `${XML_HEAD}<w:${root} ${W_NS} ${DRAWING_NS}>${separators}${notes}</w:${root}>`, `${CONTENT_TYPE}.${root}+xml`, root);
    }
    const body = doc.sections.map((section, i) => this.sectionXml(section, i === doc.sections.length - 1)).join("");
    for (let i = 0; i < doc.pictures.length; i++) {
      const picture = doc.pictures[i];
      const name = `media/image${i + 1}.${picture.ext}`;
      this.parts[`word/${name}`] = picture.data;
      this.rels.push({ id: `rIdImage${i + 1}`, type: `${REL_NS}/image`, target: name });
    }
    this.parts["word/document.xml"] = this.encoder.encode(`${XML_HEAD}<w:document ${W_NS} ${DRAWING_NS}><w:body>${body}</w:body></w:document>`);
    this.parts["word/_rels/document.xml.rels"] = this.encoder.encode(
      `${XML_HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${this.rels.map((rel) => `<Relationship Id="${rel.id}" Type="${rel.type}" Target="${escapeAttr(rel.target)}"/>`).join("")}</Relationships>`,
    );
    this.parts["_rels/.rels"] = this.encoder.encode(
      `${XML_HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL_NS}/officeDocument" Target="word/document.xml"/></Relationships>`,
    );
    const defaults = Object.entries(PICTURE_TYPES).map(([ext, type]) => `<Default Extension="${ext}" ContentType="${type}"/>`).join("");
    const overrides = this.overrides.map((o) => `<Override PartName="${o.partName}" ContentType="${o.contentType}"/>`).join("");
    this.parts["[Content_Types].xml"] = this.encoder.encode(
      `${XML_HEAD}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>${defaults}<Override PartName="/word/document.xml" ContentType="${CONTENT_TYPE}.document.main+xml"/>${overrides}</Types>`,
    );
    return this.parts;
  }

  private addPart(name: string, xml: string, contentType: string, relType: string): string {
    this.parts[name] = this.encoder.encode(xml);
    this.overrides.push({ partName: "/" + name, contentType });
    const id = `rId${this.rels.length + 1}`;
    this.rels.push({ id, type: `${REL_NS}/${relType}`, target: name.slice("word/".length) });
    return id;
  }

  private stylesXml(): string {
    const { doc, ctx } = this;
    const out: string[] = [`${XML_HEAD}<w:styles ${W_NS}>`];
    out.push('<w:docDefaults><w:rPrDefault><w:rPr/></w:rPrDefault><w:pPrDefault><w:pPr><w:widowControl/></w:pPr></w:pPrDefault></w:docDefaults>');
    doc.styles.forEach((style, i) => {
      const pPr = pPrXml({ ...style.para, istd: 0 }, ctx).before;
      const rPr = rPrXml(style.run, ctx, 0);
      const next = style.next !== undefined && style.next !== i && doc.styles[style.next] ? `<w:next w:val="${escapeAttr(styleId(doc, style.next))}"/>` : "";
      out.push(
        `<w:style w:type="paragraph"${i === 0 ? ' w:default="1"' : ""} w:styleId="${escapeAttr(styleId(doc, i))}"><w:name w:val="${escapeAttr(style.name)}"/>${next}<w:qFormat/>` +
          `${pPr ? `<w:pPr>${pPr}</w:pPr>` : ""}${rPr ? `<w:rPr>${rPr}</w:rPr>` : ""}</w:style>`,
      );
    });
    out.push('<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/><w:uiPriority w:val="1"/><w:semiHidden/></w:style>');
    out.push(
      '<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:semiHidden/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>',
    );
    out.push("</w:styles>");
    return out.join("");
  }

  private settingsXml(): string {
    const { doc } = this;
    let out = `${XML_HEAD}<w:settings ${W_NS}>`;
    if (doc.defaultTab !== undefined && doc.defaultTab > 0) out += `<w:defaultTabStop w:val="${doc.defaultTab}"/>`;
    // These documents come from before Word 97: Word lays them out in its oldest compatibility mode.
    out += '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="11"/></w:compat>';
    return out + "</w:settings>";
  }

  private sectionXml(section: Section, last: boolean): string {
    let refs = "";
    for (const header of section.headers) {
      const n = header.kind === "header" ? ++this.headerCount : ++this.footerCount;
      const root = header.kind === "header" ? "hdr" : "ftr";
      const id = this.addPart(`word/${header.kind}${n}.xml`, `${XML_HEAD}<w:${root} ${W_NS} ${DRAWING_NS}>${this.blocksXml(header.blocks) || "<w:p/>"}</w:${root}>`, `${CONTENT_TYPE}.${header.kind}+xml`, header.kind);
      refs += `<w:${header.kind}Reference w:type="${header.type}" r:id="${id}"/>`;
    }
    const sectPr = `<w:sectPr>${sectPrXml(section.props, refs, "", "")}</w:sectPr>`;
    const blocks = section.blocks.length ? section.blocks : [emptyParagraph()];
    if (last) return this.blocksXml(blocks) + sectPr;
    // A section's sectPr is in its last paragraph's pPr; a section that ends with a table gets an empty paragraph for it.
    const lastBlock = blocks[blocks.length - 1];
    const withMark = lastBlock.type === "paragraph" ? blocks : [...blocks, emptyParagraph()];
    return this.blocksXml(withMark.slice(0, -1)) + this.paragraphXml(withMark[withMark.length - 1] as Paragraph, sectPr);
  }

  private blocksXml(blocks: Block[]): string {
    return blocks.map((block) => (block.type === "paragraph" ? this.paragraphXml(block, "") : this.tableXml(block))).join("");
  }

  private paragraphXml(p: Paragraph, sectPr: string): string {
    const { ctx } = this;
    let pPr = pPrXml(p.props, ctx).before;
    const markRPr = rPrXml(p.mark, ctx, p.props.istd);
    if (markRPr) pPr += `<w:rPr>${markRPr}</w:rPr>`;
    pPr += sectPr;
    return `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ""}${p.runs.map((run) => this.runXml(run, p.props.istd)).join("")}</w:p>`;
  }

  private runXml(run: Run, istd: number): string {
    const rPr = rPrXml(run.props, this.ctx, istd);
    const open = `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ""}`;
    let out = "";
    let content = "";
    const flush = (): void => {
      if (content) out += `${open}${content}</w:r>`;
      content = "";
    };
    for (const item of run.content) {
      switch (item.kind) {
        case "text": content += `<w:t xml:space="preserve">${escapeText(item.text)}</w:t>`; break;
        case "tab": content += "<w:tab/>"; break;
        case "break": content += "<w:br/>"; break;
        case "page": content += '<w:br w:type="page"/>'; break;
        case "column": content += '<w:br w:type="column"/>'; break;
        case "noBreakHyphen": content += "<w:noBreakHyphen/>"; break;
        case "softHyphen": content += "<w:softHyphen/>"; break;
        case "footnoteMark": content += `<w:${this.noteKind}Ref/>`; break;
        case "footnote":
          content += item.customMark
            ? `<w:${this.noteKind}Reference w:customMarkFollows="1" w:id="${item.id}"/><w:t xml:space="preserve">${escapeText(item.customMark)}</w:t>`
            : `<w:${this.noteKind}Reference w:id="${item.id}"/>`;
          break;
        case "field":
          flush();
          out +=
            `${open}<w:fldChar w:fldCharType="begin"/></w:r>${open}<w:instrText xml:space="preserve"> ${escapeText(item.code)} </w:instrText></w:r>` +
            `${open}<w:fldChar w:fldCharType="separate"/></w:r>${open}<w:t xml:space="preserve">${escapeText(item.result)}</w:t></w:r>${open}<w:fldChar w:fldCharType="end"/></w:r>`;
          break;
        case "picture":
          content += this.pictureXml(item.index, item.width, item.height);
          break;
      }
    }
    flush();
    return out;
  }

  private pictureXml(index: number, width: number, height: number): string {
    const id = ++this.docPrId;
    // Twips to EMU.
    const cx = Math.max(1, Math.round(width * 635));
    const cy = Math.max(1, Math.round(height * 635));
    return (
      `<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>` +
      `<wp:docPr id="${id}" name="Picture ${id}"/><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>` +
      `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="Picture ${id}"/><pic:cNvPicPr/></pic:nvPicPr>` +
      `<pic:blipFill><a:blip r:embed="rIdImage${index + 1}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
      `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>`
    );
  }

  private tableXml(table: Table): string {
    // The grid: every cell edge of every row, merged within two twips.
    const edges: number[] = [];
    for (const row of table.rows) {
      edges.push(row.props.left ?? 0);
      for (const cell of row.props.cells) edges.push(cell.right);
    }
    edges.sort((a, b) => a - b);
    const grid: number[] = [];
    for (const edge of edges) if (grid.length === 0 || edge - grid[grid.length - 1] > 2) grid.push(edge);
    const column = (edge: number): number => {
      let best = 0;
      for (let k = 1; k < grid.length; k++) if (Math.abs(grid[k] - edge) < Math.abs(grid[best] - edge)) best = k;
      return best;
    };
    const first = table.rows[0].props;
    const tblInd = grid[0] + (first.dxaGapHalf ?? 0);
    let out = `<w:tbl><w:tblPr>${tblPrXml(first, this.ctx, tblInd)}</w:tblPr>`;
    out += `<w:tblGrid>${grid.slice(1).map((edge, k) => `<w:gridCol w:w="${edge - grid[k]}"/>`).join("")}</w:tblGrid>`;
    for (const row of table.rows) out += this.rowXml(row, grid.length - 1, column);
    return out + "</w:tbl>";
  }

  private rowXml(row: Row, columns: number, column: (edge: number) => number): string {
    const t = row.props;
    const left = t.left ?? 0;
    const lastRight = t.cells.length ? t.cells[t.cells.length - 1].right : left;
    const trPr = trPrXml(t, column(left), columns - column(lastRight));
    let out = `<w:tr>${trPr ? `<w:trPr>${trPr}</w:trPr>` : ""}`;
    let prevEdge = left;
    for (let c = 0; c < t.cells.length; c++) {
      const cell = t.cells[c];
      const span = Math.max(1, column(cell.right) - column(prevEdge));
      const content = row.cells[c] && row.cells[c].length ? this.blocksXml(row.cells[c]) : "<w:p/>";
      out += `<w:tc><w:tcPr>${tcPrXml(cell, cell.right - prevEdge, span, false)}</w:tcPr>${content}</w:tc>`;
      prevEdge = cell.right;
    }
    // Cell content with no cell definition still shows.
    for (let c = t.cells.length; c < row.cells.length; c++) out += `<w:tc><w:tcPr><w:tcW w:w="0" w:type="auto"/></w:tcPr>${this.blocksXml(row.cells[c]) || "<w:p/>"}</w:tc>`;
    if (t.cells.length === 0 && row.cells.length === 0) out += "<w:tc><w:p/></w:tc>";
    return out + "</w:tr>";
  }
}

function emptyParagraph(): Paragraph {
  return { type: "paragraph", props: { istd: 0, tabs: [], borders: {}, frame: {} } as ParaProps, mark: { toggles: {} } as RunProps, runs: [] };
}

/** A style's id: its name without what an id cannot hold, unique among the document's styles. */
function styleId(doc: OldDocument, index: number): string {
  if (index === 0) return "Normal";
  const base = doc.styles[index].name.replace(/[^A-Za-z0-9]/g, "") || "Style";
  return `${base}${index}`;
}
