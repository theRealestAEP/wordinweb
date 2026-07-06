import { DocxDocument } from "../docx.js";
import { Block, Paragraph } from "../model.js";
import { XmlElement, child, localName } from "../xml.js";
import { pxToTwips } from "../units.js";

/**
 * Block- and document-level edit commands: table insertion, paragraph
 * alignment, section/page layout. Like all commands, these mutate the source
 * XML; callers checkpoint history and refresh/relayout afterwards.
 */

function el(name: string, attrs: Record<string, string> = {}, children: XmlElement[] = [], text = ""): XmlElement {
  return { name, attrs, children, text };
}

function prefixOf(e: XmlElement): string {
  return e.name.includes(":") ? e.name.slice(0, e.name.indexOf(":") + 1) : "";
}

/** Paragraph element containing a given (caret) element, or null. */
export function paragraphOf(doc: DocxDocument, target: XmlElement): XmlElement | null {
  let cur: XmlElement | undefined = doc.findParentOf(target);
  while (cur && localName(cur.name) !== "p") cur = doc.findParentOf(cur);
  return cur ?? null;
}

// ---------- tables ----------

/**
 * Insert a rows×cols bordered table after the paragraph containing `caretT`.
 * Column widths split the section's content width evenly.
 */
export function insertTableAfter(
  doc: DocxDocument,
  caretT: XmlElement,
  rows: number,
  cols: number,
): boolean {
  const pEl = paragraphOf(doc, caretT);
  if (!pEl) return false;
  const parent = doc.findParentOf(pEl);
  if (!parent) return false;
  const w = prefixOf(pEl);

  const sp = doc.sections[0]?.props;
  const contentPx = sp ? sp.pageWidth - sp.marginLeft - sp.marginRight : 624;
  const colTwips = Math.floor(pxToTwips(contentPx) / cols);

  const border = (side: string) =>
    el(`${w}${side}`, { [`${w}val`]: "single", [`${w}sz`]: "4", [`${w}space`]: "0", [`${w}color`]: "auto" });
  const tblPr = el(`${w}tblPr`, {}, [
    el(`${w}tblW`, { [`${w}w`]: "0", [`${w}type`]: "auto" }),
    el(`${w}tblBorders`, {}, ["top", "left", "bottom", "right", "insideH", "insideV"].map(border)),
  ]);
  const grid = el(
    `${w}tblGrid`,
    {},
    Array.from({ length: cols }, () => el(`${w}gridCol`, { [`${w}w`]: String(colTwips) })),
  );
  const makeCell = () =>
    el(`${w}tc`, {}, [
      el(`${w}tcPr`, {}, [el(`${w}tcW`, { [`${w}w`]: String(colTwips), [`${w}type`]: "dxa" })]),
      el(`${w}p`, {}, [el(`${w}r`, {}, [el(`${w}t`, { "xml:space": "preserve" })])]),
    ]);
  const trs = Array.from({ length: rows }, () =>
    el(`${w}tr`, {}, Array.from({ length: cols }, makeCell)),
  );
  const tbl = el(`${w}tbl`, {}, [tblPr, grid, ...trs]);

  // Word requires a paragraph between/after tables; add an empty one.
  const after = el(`${w}p`, {}, [el(`${w}r`, {}, [el(`${w}t`, { "xml:space": "preserve" })])]);

  const idx = parent.children.indexOf(pEl);
  parent.children.splice(idx + 1, 0, tbl, after);
  doc.refresh();
  return true;
}

// ---------- paragraph alignment ----------

export type ParagraphAlignment = "left" | "center" | "right" | "justify";

const JC_VAL: Record<ParagraphAlignment, string> = {
  left: "left",
  center: "center",
  right: "right",
  justify: "both",
};

/** Set w:jc on the paragraphs containing the given (caret/selection) elements. */
export function setParagraphAlignment(
  doc: DocxDocument,
  targets: XmlElement[],
  align: ParagraphAlignment,
): boolean {
  const paragraphs = new Set<XmlElement>();
  for (const t of targets) {
    const p = paragraphOf(doc, t);
    if (p) paragraphs.add(p);
  }
  if (paragraphs.size === 0) return false;
  for (const pEl of paragraphs) {
    const w = prefixOf(pEl);
    let pPr = pEl.children.find((c) => localName(c.name) === "pPr");
    if (!pPr) {
      pPr = el(`${w}pPr`);
      pEl.children.unshift(pPr);
    }
    const existing = pPr.children.findIndex((c) => localName(c.name) === "jc");
    const jc = el(`${w}jc`, { [`${w}val`]: JC_VAL[align] });
    if (existing !== -1) pPr.children[existing] = jc;
    else pPr.children.push(jc);
  }
  doc.refresh();
  return true;
}

// ---------- page layout ----------

export interface PageLayoutPatch {
  /** Margins in inches. */
  margins?: { top?: number; right?: number; bottom?: number; left?: number };
  /** Page size in inches (before orientation). */
  size?: { width: number; height: number };
  orientation?: "portrait" | "landscape";
}

const TWIPS_PER_INCH = 1440;

/**
 * Update every section's pgSz/pgMar in the document. Values in inches.
 */
export function setPageLayout(doc: DocxDocument, patch: PageLayoutPatch): boolean {
  const sectPrs: XmlElement[] = [];
  const walk = (e: XmlElement) => {
    if (localName(e.name) === "sectPr") sectPrs.push(e);
    for (const c of e.children) walk(c);
  };
  walk(doc.editableRoots()[0]);
  if (sectPrs.length === 0) return false;

  for (const sectPr of sectPrs) {
    const w = prefixOf(sectPr);
    let pgSz = child(sectPr, "pgSz");
    if (!pgSz) {
      pgSz = el(`${w}pgSz`, { [`${w}w`]: "12240", [`${w}h`]: "15840" });
      sectPr.children.unshift(pgSz);
    }
    let pgMar = child(sectPr, "pgMar");
    if (!pgMar) {
      pgMar = el(`${w}pgMar`, {
        [`${w}top`]: "1440", [`${w}right`]: "1440", [`${w}bottom`]: "1440", [`${w}left`]: "1440",
        [`${w}header`]: "720", [`${w}footer`]: "720", [`${w}gutter`]: "0",
      });
      const idx = sectPr.children.indexOf(pgSz);
      sectPr.children.splice(idx + 1, 0, pgMar);
    }

    const setAttr = (elm: XmlElement, local: string, v: string) => {
      const key = Object.keys(elm.attrs).find((k) => localName(k) === local) ?? `${w}${local}`;
      elm.attrs[key] = v;
    };
    const getAttr = (elm: XmlElement, local: string): string | undefined => {
      const key = Object.keys(elm.attrs).find((k) => localName(k) === local);
      return key ? elm.attrs[key] : undefined;
    };

    if (patch.size) {
      setAttr(pgSz, "w", String(Math.round(patch.size.width * TWIPS_PER_INCH)));
      setAttr(pgSz, "h", String(Math.round(patch.size.height * TWIPS_PER_INCH)));
    }
    if (patch.orientation) {
      const cw = parseInt(getAttr(pgSz, "w") ?? "12240", 10);
      const ch = parseInt(getAttr(pgSz, "h") ?? "15840", 10);
      const wantLandscape = patch.orientation === "landscape";
      const isLandscape = cw > ch;
      if (wantLandscape !== isLandscape) {
        setAttr(pgSz, "w", String(ch));
        setAttr(pgSz, "h", String(cw));
      }
      setAttr(pgSz, "orient", patch.orientation);
    }
    if (patch.margins) {
      for (const side of ["top", "right", "bottom", "left"] as const) {
        const v = patch.margins[side];
        if (v !== undefined) setAttr(pgMar, side, String(Math.round(v * TWIPS_PER_INCH)));
      }
    }
  }
  doc.refresh();
  return true;
}

/**
 * Merge a paragraph into the one before it (Backspace at paragraph start).
 * The previous paragraph's pPr wins, matching Word. Returns false when
 * there is no preceding paragraph sibling (start of document/cell, or a
 * table sits in between).
 */
export function mergeParagraphBackward(doc: DocxDocument, pEl: XmlElement): boolean {
  const parent = doc.findParentOf(pEl);
  if (!parent) return false;
  const idx = parent.children.indexOf(pEl);
  if (idx <= 0) return false;
  const prev = parent.children[idx - 1];
  if (localName(prev.name) !== "p") return false;
  const moved = pEl.children.filter((c) => localName(c.name) !== "pPr");
  prev.children.push(...moved);
  parent.children.splice(idx, 1);
  doc.refresh();
  return true;
}

/** First w:t element inside a paragraph (document order), if any. */
export function firstTextOf(pEl: XmlElement): XmlElement | null {
  const walk = (el: XmlElement): XmlElement | null => {
    for (const c of el.children) {
      if (localName(c.name) === "pPr") continue;
      if (localName(c.name) === "t") return c;
      const found = walk(c);
      if (found) return found;
    }
    return null;
  };
  return walk(pEl);
}

/** Last w:t element inside a paragraph (document order), if any. */
export function lastTextOf(pEl: XmlElement): XmlElement | null {
  let last: XmlElement | null = null;
  const walk = (el: XmlElement) => {
    for (const c of el.children) {
      if (localName(c.name) === "t") last = c;
      walk(c);
    }
  };
  walk(pEl);
  return last;
}

/** Next/previous sibling paragraph of pEl within its container, if any. */
export function siblingParagraph(doc: DocxDocument, pEl: XmlElement, dir: -1 | 1): XmlElement | null {
  const parent = doc.findParentOf(pEl);
  if (!parent) return null;
  const idx = parent.children.indexOf(pEl);
  const sib = parent.children[idx + dir];
  return sib && localName(sib.name) === "p" ? sib : null;
}

const EMU_PER_PX = 9525;
const NS_WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const NS_PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture";
const NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

/**
 * Insert an inline image (already registered via doc.addImageResource) after
 * the run containing the caret. Namespaces are declared on the drawing
 * subtree so the document root needs no changes.
 */
export function insertImageAt(
  doc: DocxDocument,
  caretT: XmlElement,
  relId: string,
  widthPx: number,
  heightPx: number,
): XmlElement | null {
  const rEl = doc.findParentOf(caretT);
  if (!rEl || localName(rEl.name) !== "r") return null;
  const parent = doc.findParentOf(rEl);
  if (!parent) return null;
  const w = prefixOf(rEl);
  const cx = String(Math.round(widthPx * EMU_PER_PX));
  const cy = String(Math.round(heightPx * EMU_PER_PX));
  const id = String(doc.nextDrawingId());

  const drawing = el(`${w}drawing`, {}, [
    el("wp:inline", { "xmlns:wp": NS_WP, distT: "0", distB: "0", distL: "0", distR: "0" }, [
      el("wp:extent", { cx, cy }),
      el("wp:docPr", { id, name: `Picture ${id}` }),
      el("a:graphic", { "xmlns:a": NS_A }, [
        el("a:graphicData", { uri: NS_PIC }, [
          el("pic:pic", { "xmlns:pic": NS_PIC }, [
            el("pic:nvPicPr", {}, [
              el("pic:cNvPr", { id, name: `Picture ${id}` }),
              el("pic:cNvPicPr"),
            ]),
            el("pic:blipFill", {}, [
              el("a:blip", { "xmlns:r": NS_R, "r:embed": relId }),
              el("a:stretch", {}, [el("a:fillRect")]),
            ]),
            el("pic:spPr", {}, [
              el("a:xfrm", {}, [el("a:off", { x: "0", y: "0" }), el("a:ext", { cx, cy })]),
              el("a:prstGeom", { prst: "rect" }, [el("a:avLst")]),
            ]),
          ]),
        ]),
      ]),
    ]),
  ]);
  const run = el(`${w}r`, {}, [drawing]);
  parent.children.splice(parent.children.indexOf(rEl) + 1, 0, run);
  doc.refresh();
  return drawing;
}

/** Model paragraph for a w:p element (searched across sections and tables). */
function modelParagraphOf(doc: DocxDocument, pEl: XmlElement): Paragraph | null {
  const walk = (blocks: Block[]): Paragraph | null => {
    for (const b of blocks) {
      if (b.type === "paragraph") {
        if (b.src === pEl) return b;
      } else if (b.type === "table") {
        for (const row of b.rows) {
          for (const cell of row.cells) {
            const hit = walk(cell.blocks);
            if (hit) return hit;
          }
        }
      }
    }
    return null;
  };
  for (const s of doc.sections) {
    const hit = walk(s.blocks);
    if (hit) return hit;
  }
  return null;
}

/**
 * Fixed line height (px) of the paragraph containing target, when its
 * effective line spacing uses the "exact" rule — content taller than this
 * cannot grow the line (Word clips it). Null for auto/atLeast spacing.
 */
export function exactLineHeightAt(doc: DocxDocument, target: XmlElement): number | null {
  const pEl = paragraphOf(doc, target);
  if (!pEl) return null;
  const para = modelParagraphOf(doc, pEl);
  if (!para) return null;
  const ls = doc.effectiveParaProps(para).lineSpacing;
  return ls?.rule === "exact" ? ls.value : null;
}

/** The pStyle id of the paragraph containing target (null = Normal/none). */
export function paragraphStyleIdOf(doc: DocxDocument, target: XmlElement): string | null {
  const pEl = paragraphOf(doc, target);
  if (!pEl) return null;
  const pPr = pEl.children.find((c) => localName(c.name) === "pPr");
  const pStyle = pPr?.children.find((c) => localName(c.name) === "pStyle");
  if (!pStyle) return null;
  const key = Object.keys(pStyle.attrs).find((k) => localName(k) === "val");
  return key ? pStyle.attrs[key] : null;
}

/** Apply (or clear, with null) a paragraph style to the target paragraphs. */
export function setParagraphStyle(
  doc: DocxDocument,
  targets: XmlElement[],
  styleId: string | null,
): boolean {
  // Word's built-in styles work without a declaration; inject one if needed.
  if (styleId !== null) doc.ensureParagraphStyle(styleId);
  const paragraphs = new Set<XmlElement>();
  for (const t of targets) {
    const p = paragraphOf(doc, t);
    if (p) paragraphs.add(p);
  }
  if (paragraphs.size === 0) return false;
  for (const pEl of paragraphs) {
    const w = prefixOf(pEl);
    let pPr = pEl.children.find((c) => localName(c.name) === "pPr");
    if (!pPr) {
      pPr = el(`${w}pPr`);
      pEl.children.unshift(pPr);
    }
    const idx = pPr.children.findIndex((c) => localName(c.name) === "pStyle");
    if (styleId === null) {
      if (idx !== -1) pPr.children.splice(idx, 1);
    } else {
      const st = el(`${w}pStyle`, { [`${w}val`]: styleId });
      if (idx !== -1) pPr.children[idx] = st;
      else pPr.children.unshift(st); // pStyle must lead pPr
    }
  }
  doc.refresh();
  return true;
}
