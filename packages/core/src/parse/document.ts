import { XmlElement, attr, child, children, intAttr, localName, onOff, path } from "../xml.js";
import {
  Block,
  Border,
  DrawingContent,
  DrawingImage,
  DrawingLine,
  Hyperlink,
  ImageContent,
  ParaChild,
  Paragraph,
  Run,
  RunContent,
  Section,
  Table,
  TableCell,
  TableCellProps,
  TableProps,
  TableRow,
  TableRowProps,
} from "../model.js";
import { emuToPx, twipsToPx } from "../units.js";
import { ParseContext, parseBorder, parseParaProps, parseRunProps, parseShading } from "./properties.js";
import { parseSectionProps, defaultSectionProps } from "./section.js";
import { Relationships } from "./rels.js";

export interface DocParseContext extends ParseContext {
  rels: Relationships;
}

/** Parse w:body into sections (the body's trailing sectPr closes the last one). */
export function parseBody(body: XmlElement, ctx: DocParseContext): Section[] {
  const sections: Section[] = [];
  let blocks: Block[] = [];

  for (const el of body.children) {
    const ln = localName(el.name);
    if (ln === "p") {
      const para = parseParagraph(el, ctx);
      blocks.push(para);
      if (para.sectionBreak) {
        sections.push({ props: para.sectionBreak, blocks });
        blocks = [];
      }
    } else if (ln === "tbl") {
      blocks.push(parseTable(el, ctx));
    } else if (ln === "sectPr") {
      sections.push({ props: parseSectionProps(el), blocks });
      blocks = [];
    }
    // bookmarkStart/End, proofErr, sdt etc. — sdt handled below, rest ignored
    else if (ln === "sdt") {
      const content = child(el, "sdtContent");
      if (content) {
        for (const inner of content.children) {
          const iln = localName(inner.name);
          if (iln === "p") blocks.push(parseParagraph(inner, ctx));
          else if (iln === "tbl") blocks.push(parseTable(inner, ctx));
        }
      }
    }
  }
  if (blocks.length > 0) {
    sections.push({ props: defaultSectionProps(), blocks });
  }
  return sections;
}

/** Parse block-level content of a header/footer/table-cell container. */
export function parseBlocks(container: XmlElement, ctx: DocParseContext): Block[] {
  const blocks: Block[] = [];
  for (const el of container.children) {
    const ln = localName(el.name);
    if (ln === "p") blocks.push(parseParagraph(el, ctx));
    else if (ln === "tbl") blocks.push(parseTable(el, ctx));
    else if (ln === "sdt") {
      const content = child(el, "sdtContent");
      if (content) blocks.push(...parseBlocks(content, ctx));
    }
  }
  return blocks;
}

// ---------- paragraphs ----------

interface FieldState {
  /** null when not inside a complex field. */
  mode: "instr" | "result" | null;
  instruction: string;
  cachedResult: string;
  /** Run that will carry the resulting field content. */
  carrier: Run | null;
}

export function parseParagraph(p: XmlElement, ctx: DocParseContext): Paragraph {
  const pPr = child(p, "pPr");
  const para: Paragraph = {
    type: "paragraph",
    props: parseParaProps(pPr, ctx),
    children: [],
    src: p,
  };
  const sectPr = child(pPr, "sectPr");
  if (sectPr) para.sectionBreak = parseSectionProps(sectPr);

  const field: FieldState = { mode: null, instruction: "", cachedResult: "", carrier: null };
  parseParaChildren(p, ctx, para.children, field);
  // Unterminated field (shouldn't happen in valid files): flush as text
  flushField(field);
  return para;
}

function parseParaChildren(
  parent: XmlElement,
  ctx: DocParseContext,
  out: ParaChild[],
  field: FieldState,
): void {
  for (const el of parent.children) {
    const ln = localName(el.name);
    if (ln === "r") {
      const run = parseRun(el, ctx, field);
      if (run) run.srcParent = parent;
      // Keep empty runs that carry an open complex field: the field content
      // is appended to the carrier when fldChar end arrives.
      if (run && (run.content.length > 0 || field.carrier === run)) out.push(run);
    } else if (ln === "hyperlink") {
      const rid = attr(el, "id");
      const rel = rid ? ctx.rels.get(rid) : undefined;
      const link: Hyperlink = {
        type: "hyperlink",
        href: rel?.external ? rel.target : undefined,
        anchor: attr(el, "anchor"),
        runs: [],
      };
      const inner: ParaChild[] = [];
      parseParaChildren(el, ctx, inner, field);
      for (const c of inner) {
        if (c.type === "run") link.runs.push(c);
        else link.runs.push(...c.runs);
      }
      if (link.runs.length > 0) out.push(link);
    } else if (ln === "fldSimple") {
      const instruction = attr(el, "instr") ?? "";
      const inner: ParaChild[] = [];
      const innerField: FieldState = { mode: null, instruction: "", cachedResult: "", carrier: null };
      parseParaChildren(el, ctx, inner, innerField);
      let cached = "";
      let props = {};
      for (const c of inner) {
        const runs = c.type === "run" ? [c] : c.runs;
        for (const r of runs) {
          if (Object.keys(props).length === 0) props = r.props;
          for (const rc of r.content) if (rc.kind === "text") cached += rc.text;
        }
      }
      out.push({
        type: "run",
        props,
        content: [{ kind: "field", instruction, cachedResult: cached }],
      });
    } else if (ln === "smartTag" || ln === "sdt") {
      const content = ln === "sdt" ? child(el, "sdtContent") : el;
      if (content) parseParaChildren(content, ctx, out, field);
    }
    // bookmarkStart/bookmarkEnd/proofErr/commentRangeStart... ignored
  }
}

function flushField(field: FieldState): void {
  if (field.mode !== null && field.carrier) {
    field.carrier.content.push({
      kind: "field",
      instruction: field.instruction,
      cachedResult: field.cachedResult,
    });
  }
  field.mode = null;
  field.instruction = "";
  field.cachedResult = "";
  field.carrier = null;
}

/**
 * Parse a run. Complex-field runs (fldChar/instrText) mutate `field` state;
 * the field content is emitted on the run carrying fldChar begin when the
 * field closes.
 */
function parseRun(r: XmlElement, ctx: DocParseContext, field: FieldState): Run | null {
  const run: Run = {
    type: "run",
    props: parseRunProps(child(r, "rPr"), ctx),
    content: [],
    src: r,
  };

  for (const el of r.children) {
    const ln = localName(el.name);
    switch (ln) {
      case "t":
        if (field.mode === "result") field.cachedResult += el.text;
        else if (field.mode !== "instr") run.content.push({ kind: "text", text: el.text, srcT: el });
        break;
      case "instrText":
        if (field.mode === "instr") field.instruction += el.text;
        break;
      case "fldChar": {
        const type = attr(el, "fldCharType");
        if (type === "begin") {
          flushField(field);
          field.mode = "instr";
          field.carrier = run;
        } else if (type === "separate") {
          field.mode = "result";
        } else if (type === "end") {
          const carrier = field.carrier ?? run;
          const content: RunContent = {
            kind: "field",
            instruction: field.instruction,
            cachedResult: field.cachedResult,
          };
          if (carrier === run) run.content.push(content);
          else carrier.content.push(content);
          field.mode = null;
          field.instruction = "";
          field.cachedResult = "";
          field.carrier = null;
        }
        break;
      }
      case "br": {
        const type = attr(el, "type");
        run.content.push({
          kind: "break",
          breakType: type === "page" ? "page" : type === "column" ? "column" : "line",
        });
        break;
      }
      case "tab":
        run.content.push({ kind: "tab" });
        break;
      case "drawing": {
        const img = parseDrawing(el, ctx);
        if (img) {
          if (img.kind === "image") img.srcDrawing = el;
          run.content.push(img);
        }
        break;
      }
      case "pict":
        run.content.push(...parseVmlPict(el, ctx));
        break;
      case "AlternateContent": {
        // Prefer the DrawingML choice when it's a plain picture. For choices
        // we can't render (wpg groups etc.), producers ship a Fallback that
        // is either a rasterized w:drawing (Google Docs) or VML w:pict.
        const choice = child(el, "Choice");
        const choiceDrawing = choice ? child(choice, "drawing") : undefined;
        const img = choiceDrawing ? parseDrawing(choiceDrawing, ctx) : null;
        if (img) {
          if (img.kind === "image") img.srcDrawing = choiceDrawing;
          run.content.push(img);
        } else {
          const fallback = child(el, "Fallback");
          const fbDrawing = fallback ? child(fallback, "drawing") : undefined;
          const fbImg = fbDrawing ? parseDrawing(fbDrawing, ctx) : null;
          if (fbImg) {
            if (fbImg.kind === "image") fbImg.srcDrawing = fbDrawing;
            run.content.push(fbImg);
          } else {
            const pictEl = fallback ? child(fallback, "pict") : undefined;
            if (pictEl) run.content.push(...parseVmlPict(pictEl, ctx));
          }
        }
        break;
      }
      case "noBreakHyphen":
        run.content.push({ kind: "text", text: "‑" });
        break;
      case "softHyphen":
        run.content.push({ kind: "text", text: "­" });
        break;
      case "sym": {
        const charHex = attr(el, "char");
        if (charHex) {
          let code = parseInt(charHex, 16);
          if (code >= 0xf000) code -= 0xf000; // private-use offset Word applies
          run.content.push({ kind: "text", text: String.fromCharCode(code) });
        }
        break;
      }
      case "cr":
        run.content.push({ kind: "break", breakType: "line" });
        break;
    }
  }

  return run;
}

// ---------- drawings ----------

function findDescendant(el: XmlElement, local: string): XmlElement | undefined {
  for (const c of el.children) {
    if (localName(c.name) === local) return c;
    const found = findDescendant(c, local);
    if (found) return found;
  }
  return undefined;
}

const EMU_PER_PT = 12700;

function parseDrawing(
  drawing: XmlElement,
  ctx: DocParseContext,
): ImageContent | DrawingContent | null {
  const inline = child(drawing, "inline");
  const anchor = child(drawing, "anchor");
  const holder = inline ?? anchor;
  if (!holder) return null;
  const extent = child(holder, "extent");
  const cx = intAttr(extent, "cx") ?? 0;
  const cy = intAttr(extent, "cy") ?? 0;

  const images: DrawingImage[] = [];
  const lines: DrawingLine[] = [];

  // Walk the graphic tree, carrying the group coordinate transform:
  // childEmu → boxEmu = (childEmu - chOff) * (ext / chExt) + off, composed.
  const walk = (el: XmlElement, ox: number, oy: number, sx: number, sy: number) => {
    const ln = localName(el.name);
    if (ln === "pic") {
      const xfrm = path(el, "spPr", "xfrm");
      const off = child(xfrm, "off");
      const ext = child(xfrm, "ext");
      const blip = findDescendant(el, "blip");
      const rid = blip ? (attr(blip, "embed") ?? attr(blip, "link")) : undefined;
      const rel = rid ? ctx.rels.get(rid) : undefined;
      if (rel && !rel.external) {
        const x = ox + (intAttr(off, "x") ?? 0) * sx;
        const y = oy + (intAttr(off, "y") ?? 0) * sy;
        const w = (intAttr(ext, "cx") ?? cx) * sx;
        const h = (intAttr(ext, "cy") ?? cy) * sy;
        images.push({ part: rel.target, x: emuToPx(x), y: emuToPx(y), width: emuToPx(w), height: emuToPx(h) });
      }
      return;
    }
    if (ln === "wsp") {
      const spPr = child(el, "spPr");
      const prst = attr(child(spPr, "prstGeom"), "prst") ?? "";
      const lnEl = child(spPr, "ln");
      const isLine = prst === "line" || prst.startsWith("straightConnector") || prst.startsWith("bentConnector");
      if (isLine && lnEl) {
        const xfrm = child(spPr, "xfrm");
        const off = child(xfrm, "off");
        const ext = child(xfrm, "ext");
        const x = ox + (intAttr(off, "x") ?? 0) * sx;
        const y = oy + (intAttr(off, "y") ?? 0) * sy;
        const w = (intAttr(ext, "cx") ?? 0) * sx;
        const h = (intAttr(ext, "cy") ?? 0) * sy;
        const flipH = attr(xfrm, "flipH") === "1";
        const flipV = attr(xfrm, "flipV") === "1";
        const srgb = findDescendant(lnEl, "srgbClr");
        const color = srgb ? "#" + (attr(srgb, "val") ?? "000000") : "#000000";
        const weightEmu = intAttr(lnEl, "w") ?? EMU_PER_PT; // default 1pt
        let [x1, y1, x2, y2] = [x, y, x + w, y + h];
        if (flipH) [x1, x2] = [x2, x1];
        if (flipV) [y1, y2] = [y2, y1];
        lines.push({
          x1: emuToPx(x1),
          y1: emuToPx(y1),
          x2: emuToPx(x2),
          y2: emuToPx(y2),
          color,
          weight: Math.max((weightEmu / EMU_PER_PT) * (4 / 3), 0.75),
        });
      }
      // Textboxes inside shapes are ignored here (rendered via VML fallback
      // when present); recurse for nested pics.
      for (const c of el.children) walk(c, ox, oy, sx, sy);
      return;
    }
    if (ln === "wgp" || ln === "grpSp") {
      const xfrm = path(el, "grpSpPr", "xfrm");
      const off = child(xfrm, "off");
      const chOff = child(xfrm, "chOff");
      const ext = child(xfrm, "ext");
      const chExt = child(xfrm, "chExt");
      const extX = intAttr(ext, "cx") ?? 1;
      const extY = intAttr(ext, "cy") ?? 1;
      const chX = intAttr(chExt, "cx") ?? extX;
      const chY = intAttr(chExt, "cy") ?? extY;
      const nsx = sx * (chX ? extX / chX : 1);
      const nsy = sy * (chY ? extY / chY : 1);
      const nox = ox + (intAttr(off, "x") ?? 0) * sx - (intAttr(chOff, "x") ?? 0) * nsx;
      const noy = oy + (intAttr(off, "y") ?? 0) * sy - (intAttr(chOff, "y") ?? 0) * nsy;
      for (const c of el.children) walk(c, nox, noy, nsx, nsy);
      return;
    }
    for (const c of el.children) walk(c, ox, oy, sx, sy);
  };
  walk(holder, 0, 0, 1, 1);

  // Plain single image covering the extent: keep the simple content kind.
  if (lines.length === 0 && images.length === 1) {
    return {
      kind: "image",
      part: images[0].part,
      width: images[0].width || emuToPx(cx),
      height: images[0].height || emuToPx(cy),
      anchored: !!anchor,
    };
  }
  if (lines.length === 0 && images.length === 0) return null;
  return {
    kind: "drawing",
    width: emuToPx(cx),
    height: emuToPx(cy),
    lines,
    images,
  };
}

// ---------- VML (legacy drawing markup: textboxes, lines, pictures) ----------

/** Parse a VML length ("36pt", "1in", "669.6pt", "12px", bare number = px). */
function vmlLength(raw: string | undefined): number {
  if (!raw) return 0;
  const m = /^(-?[\d.]+)\s*(pt|in|px|cm|mm|pc)?$/.exec(raw.trim());
  if (!m) return 0;
  const v = parseFloat(m[1]);
  switch (m[2]) {
    case "pt": return (v * 4) / 3;
    case "in": return v * 96;
    case "cm": return (v / 2.54) * 96;
    case "mm": return (v / 25.4) * 96;
    case "pc": return v * 16;
    default: return v; // px
  }
}

function parseVmlStyle(style: string | undefined): Map<string, string> {
  const out = new Map<string, string>();
  if (!style) return out;
  for (const decl of style.split(";")) {
    const idx = decl.indexOf(":");
    if (idx > 0) out.set(decl.slice(0, idx).trim(), decl.slice(idx + 1).trim());
  }
  return out;
}

function anchorRel(v: string | undefined): "page" | "margin" | "text" | "column" {
  return v === "page" || v === "margin" || v === "column" ? v : "text";
}

export function parseVmlPict(pict: XmlElement, ctx: DocParseContext): RunContent[] {
  const out: RunContent[] = [];
  const walk = (el: XmlElement) => {
    const ln = localName(el.name);
    if (ln === "shapetype") return; // template definition, not an instance
    if (ln === "line") {
      const style = parseVmlStyle(el.attrs["style"]);
      const from = (el.attrs["from"] ?? "0,0").split(",");
      const to = (el.attrs["to"] ?? "0,0").split(",");
      out.push({
        kind: "anchor",
        shape: {
          type: "line",
          x1: vmlLength(from[0]),
          y1: vmlLength(from[1]),
          x2: vmlLength(to[0]),
          y2: vmlLength(to[1]),
          color: el.attrs["strokecolor"] ?? "#000000",
          weight: vmlLength(el.attrs["strokeweight"]) || 1,
          hRel: anchorRel(style.get("mso-position-horizontal-relative")),
          vRel: anchorRel(style.get("mso-position-vertical-relative")),
        },
      });
      return;
    }
    if (ln === "shape" || ln === "rect") {
      const imagedata = findDescendant(el, "imagedata");
      if (imagedata) {
        const rid = attr(imagedata, "id");
        const rel = rid ? ctx.rels.get(rid) : undefined;
        if (rel && !rel.external) {
          const style = parseVmlStyle(el.attrs["style"]);
          out.push({
            kind: "image",
            part: rel.target,
            width: vmlLength(style.get("width")) || 100,
            height: vmlLength(style.get("height")) || 100,
          });
        }
        return;
      }
      const txbx = findDescendant(el, "txbxContent");
      if (txbx) {
        const style = parseVmlStyle(el.attrs["style"]);
        out.push({
          kind: "anchor",
          shape: {
            type: "textbox",
            x: vmlLength(style.get("margin-left")),
            y: vmlLength(style.get("margin-top")),
            width: vmlLength(style.get("width")),
            height: vmlLength(style.get("height")),
            hRel: anchorRel(style.get("mso-position-horizontal-relative")),
            vRel: anchorRel(style.get("mso-position-vertical-relative")),
            blocks: parseBlocks(txbx, ctx),
          },
        });
        return;
      }
      return;
    }
    for (const c of el.children) walk(c);
  };
  for (const c of pict.children) walk(c);
  return out;
}

// ---------- tables ----------

export function parseTable(tbl: XmlElement, ctx: DocParseContext): Table {
  const tblPr = child(tbl, "tblPr");
  const props: TableProps = {};

  if (tblPr) {
    const styleId = attr(child(tblPr, "tblStyle"), "val");
    if (styleId) props.styleId = styleId;
    const ind = child(tblPr, "tblInd");
    if (ind) {
      const w = intAttr(ind, "w");
      if (w !== undefined && attr(ind, "type") !== "pct") props.indent = twipsToPx(w);
    }
    const jc = attr(child(tblPr, "jc"), "val");
    if (jc === "center" || jc === "right") props.alignment = jc;
    const borders = child(tblPr, "tblBorders");
    if (borders) {
      props.borders = {
        top: parseBorder(child(borders, "top"), ctx),
        bottom: parseBorder(child(borders, "bottom"), ctx),
        left: parseBorder(child(borders, "left"), ctx),
        right: parseBorder(child(borders, "right"), ctx),
        insideH: parseBorder(child(borders, "insideH"), ctx),
        insideV: parseBorder(child(borders, "insideV"), ctx),
      };
    }
    const cellMar = child(tblPr, "tblCellMar");
    if (cellMar) props.cellMargins = parseCellMargins(cellMar);
    const tblW = child(tblPr, "tblW");
    if (tblW) {
      const raw = attr(tblW, "w");
      const type = attr(tblW, "type");
      const num = raw !== undefined ? parseFloat(raw) : NaN;
      if (Number.isFinite(num) && num > 0) {
        if (type === "dxa") props.width = twipsToPx(num);
        // pct: either "NN%" or fiftieths of a percent per ST_MeasurementOrPercent.
        else if (type === "pct") props.widthPct = raw!.trim().endsWith("%") ? num / 100 : num / 5000;
      }
    }
    const layout = attr(child(tblPr, "tblLayout"), "type");
    props.layout = layout === "fixed" ? "fixed" : "autofit";
  }

  const grid: number[] = [];
  const tblGrid = child(tbl, "tblGrid");
  if (tblGrid) {
    for (const col of children(tblGrid, "gridCol")) {
      grid.push(twipsToPx(intAttr(col, "w") ?? 0));
    }
  }

  const rows: TableRow[] = [];
  for (const tr of children(tbl, "tr")) {
    rows.push(parseRow(tr, ctx));
  }

  return { type: "table", props, grid, rows, src: tbl };
}

function parseCellMargins(el: XmlElement): { top?: number; right?: number; bottom?: number; left?: number } {
  const out: { top?: number; right?: number; bottom?: number; left?: number } = {};
  for (const side of ["top", "right", "bottom", "left"] as const) {
    const m = child(el, side) ?? (side === "right" ? child(el, "end") : side === "left" ? child(el, "start") : undefined);
    if (m && attr(m, "type") !== "pct") {
      const w = intAttr(m, "w");
      if (w !== undefined) out[side] = twipsToPx(w);
    }
  }
  return out;
}

function parseRow(tr: XmlElement, ctx: DocParseContext): TableRow {
  const trPr = child(tr, "trPr");
  const props: TableRowProps = {};
  if (trPr) {
    const height = child(trPr, "trHeight");
    if (height) {
      const val = intAttr(height, "val");
      if (val !== undefined) props.height = twipsToPx(val);
      const rule = attr(height, "hRule");
      props.heightRule = rule === "exact" ? "exact" : rule === "auto" ? "auto" : "atLeast";
    }
    props.cantSplit = onOff(child(trPr, "cantSplit"));
    props.tblHeader = onOff(child(trPr, "tblHeader"));
  }
  const cells: TableCell[] = [];
  for (const tc of children(tr, "tc")) {
    cells.push(parseCell(tc, ctx));
  }
  return { props, cells };
}

function parseCell(tc: XmlElement, ctx: DocParseContext): TableCell {
  const tcPr = child(tc, "tcPr");
  const props: TableCellProps = { gridSpan: 1 };
  if (tcPr) {
    const span = intAttr(child(tcPr, "gridSpan"), "val");
    if (span && span > 1) props.gridSpan = span;
    const vMerge = child(tcPr, "vMerge");
    if (vMerge) props.vMerge = attr(vMerge, "val") === "restart" ? "restart" : "continue";
    const tcW = child(tcPr, "tcW");
    if (tcW && attr(tcW, "type") === "dxa") {
      const w = intAttr(tcW, "w");
      if (w) props.width = twipsToPx(w);
    }
    const borders = child(tcPr, "tcBorders");
    if (borders) {
      props.borders = {
        top: parseBorder(child(borders, "top"), ctx),
        bottom: parseBorder(child(borders, "bottom"), ctx),
        left: parseBorder(child(borders, "left"), ctx),
        right: parseBorder(child(borders, "right"), ctx),
      };
    }
    const shd = parseShading(child(tcPr, "shd"), ctx);
    if (shd) props.shading = shd;
    const mar = child(tcPr, "tcMar");
    if (mar) props.margins = parseCellMargins(mar);
    const vAlign = attr(child(tcPr, "vAlign"), "val");
    if (vAlign === "center" || vAlign === "bottom") props.verticalAlign = vAlign;
  }
  return { props, blocks: parseBlocks(tc, ctx) };
}
