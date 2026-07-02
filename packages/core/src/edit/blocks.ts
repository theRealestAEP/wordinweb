import { DocxDocument } from "../docx.js";
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
