import { DocxDocument } from "../docx.js";
import { XmlElement, attr, localName } from "../xml.js";
import { setTabStops, type TabStopSpec } from "./paragraph.js";

/**
 * One-click header/footer content: the page-number position gallery (Word's
 * Insert > Page Number > Top/Bottom of Page) and the Header & Footer preset
 * gallery. Both compose machinery that already exists — ensureHfPart, the
 * PAGE/DATE field vocabulary fields.ts already allows, and the wave-2
 * paragraph tab-stop op — rather than inventing new wire vocabulary: a
 * gallery pick is just a specific, pre-authored arrangement of those parts.
 *
 * Like insertWatermark, a pick REPLACES the target part's content — that is
 * Word's own gallery behavior (picking "Plain Number 2" throws away whatever
 * was in the header before), so there is no equality-based no-op check here;
 * the operation always writes when it can resolve a part to write into.
 */

function el(name: string, attrs: Record<string, string> = {}, children: XmlElement[] = [], text = ""): XmlElement {
  return { name, attrs, children, text };
}

function prefixOf(e: XmlElement): string {
  return e.name.includes(":") ? e.name.slice(0, e.name.indexOf(":") + 1) : "w:";
}

// ---------------------------------------------------------------------------
// Page-number position gallery
// ---------------------------------------------------------------------------

export const PAGE_NUMBER_POSITIONS = ["top", "bottom"] as const;
export type PageNumberGalleryPosition = (typeof PAGE_NUMBER_POSITIONS)[number];

export const PAGE_NUMBER_ALIGNMENTS = ["left", "center", "right"] as const;
export type PageNumberGalleryAlign = (typeof PAGE_NUMBER_ALIGNMENTS)[number];

function pageFieldParagraph(w: string, align: PageNumberGalleryAlign): XmlElement {
  return el(`${w}p`, {}, [
    el(`${w}pPr`, {}, [el(`${w}jc`, { [`${w}val`]: align })]),
    el(`${w}fldSimple`, { [`${w}instr`]: " PAGE \\* MERGEFORMAT " }, [
      el(`${w}r`, {}, [el(`${w}t`, { "xml:space": "preserve" }, [], "1")]),
    ]),
  ]);
}

/**
 * Insert a single live PAGE field into the header ("top") or footer
 * ("bottom") part, aligned left/center/right — Word's "Plain Number 1/2/3"
 * gallery entries. Creates the part on demand (ensureHfPart, the same
 * machinery `ensureHeaderFooter` uses) and replaces its ENTIRE content with
 * the one gallery paragraph.
 */
export function insertPageNumberPosition(
  doc: DocxDocument,
  position: PageNumberGalleryPosition,
  align: PageNumberGalleryAlign,
): boolean {
  const root = doc.ensureHfPart(position === "top" ? "header" : "footer");
  const w = prefixOf(root);
  root.children = [pageFieldParagraph(w, align)];
  doc.refresh();
  return true;
}

function runText(run: XmlElement): string {
  let s = "";
  for (const c of run.children) if (localName(c.name) === "t") s += c.text;
  return s;
}

function fldCharType(run: XmlElement): string | undefined {
  const f = run.children.find((c) => localName(c.name) === "fldChar");
  return f ? attr(f, "fldCharType") : undefined;
}

function instrTextOf(run: XmlElement): string | undefined {
  const it = run.children.find((c) => localName(c.name) === "instrText");
  return it ? it.text : undefined;
}

function fieldKeyword(instruction: string): string {
  return instruction.trim().split(/\s+/)[0]?.toUpperCase() ?? "";
}

/** Remove PAGE/NUMPAGES field content from one header/footer paragraph.
 * Handles both `w:fldSimple` (what this engine's own field inserts write)
 * and a same-paragraph complex-field span (`w:fldChar` begin/…/end +
 * `w:instrText`, what Word's own Page Number command writes) — true for
 * every PAGE field a real Word document round-trips through this engine.
 * The literal "Page "/" of " text the "Page X of Y" gallery/insert wraps
 * around the fields goes with them, but only when it sits DIRECTLY next to
 * a removed field, so unrelated text reading "Page " is never touched. */
function removePageNumberFieldsFromParagraph(p: XmlElement): boolean {
  const kids = p.children;
  const drop = new Set<XmlElement>();
  let i = 0;
  while (i < kids.length) {
    const node = kids[i];
    const ln = localName(node.name);
    if (ln === "fldSimple") {
      if (["PAGE", "NUMPAGES"].includes(fieldKeyword(attr(node, "instr") ?? ""))) drop.add(node);
      i++;
      continue;
    }
    if (ln === "r" && fldCharType(node) === "begin") {
      let j = i + 1;
      let instruction = "";
      while (j < kids.length && localName(kids[j].name) === "r" && instrTextOf(kids[j]) !== undefined) {
        instruction += instrTextOf(kids[j]);
        j++;
      }
      let end = j;
      while (end < kids.length && !(localName(kids[end].name) === "r" && fldCharType(kids[end]) === "end")) end++;
      if (["PAGE", "NUMPAGES"].includes(fieldKeyword(instruction)) && end < kids.length) {
        for (let k = i; k <= end; k++) drop.add(kids[k]);
        i = end + 1;
        continue;
      }
    }
    i++;
  }
  if (drop.size === 0) return false;
  for (let idx = 0; idx < kids.length; idx++) {
    const node = kids[idx];
    if (localName(node.name) !== "r" || drop.has(node)) continue;
    const text = runText(node);
    if (text !== "Page " && text !== " of ") continue;
    const prevDropped = idx > 0 && drop.has(kids[idx - 1]);
    const nextDropped = idx + 1 < kids.length && drop.has(kids[idx + 1]);
    if (prevDropped || nextDropped) drop.add(node);
  }
  p.children = kids.filter((c) => !drop.has(c));
  return true;
}

/**
 * Word's "Remove Page Numbers": strip PAGE/NUMPAGES field content from every
 * header and footer part (default plus the first-page/even-page variants).
 * The header/footer story itself stays, possibly now empty — Word does not
 * delete the band, only its page-number content. False when none was found.
 */
export function removePageNumberFields(doc: DocxDocument): boolean {
  let changed = false;
  for (const root of [...doc.headerRoots(), ...doc.footerRoots()]) {
    for (const p of root.children) {
      if (localName(p.name) === "p" && removePageNumberFieldsFromParagraph(p)) changed = true;
    }
  }
  if (changed) doc.refresh();
  return changed;
}

// ---------------------------------------------------------------------------
// Header & Footer preset gallery
// ---------------------------------------------------------------------------

export const HEADER_FOOTER_PRESETS = ["blank", "centeredTitle", "titleAndDate", "threeColumn"] as const;
export type HeaderFooterPreset = (typeof HEADER_FOOTER_PRESETS)[number];

function emptyRun(w: string): XmlElement {
  return el(`${w}r`, {}, [el(`${w}t`, { "xml:space": "preserve" }, [], "")]);
}

function styledTextRun(
  w: string,
  text: string,
  opts: { bold?: boolean; sizePt?: number; color?: string } = {},
): XmlElement {
  const sz = opts.sizePt !== undefined ? String(Math.round(opts.sizePt * 2)) : null;
  return el(`${w}r`, {}, [
    el(`${w}rPr`, {}, [
      ...(opts.bold ? [el(`${w}b`)] : []),
      ...(opts.color ? [el(`${w}color`, { [`${w}val`]: opts.color })] : []),
      ...(sz ? [el(`${w}sz`, { [`${w}val`]: sz }), el(`${w}szCs`, { [`${w}val`]: sz })] : []),
    ]),
    el(`${w}t`, { "xml:space": "preserve" }, [], text),
  ]);
}

function tabRun(w: string): XmlElement {
  return el(`${w}r`, {}, [el(`${w}tab`)]);
}

function dateFieldRun(w: string): XmlElement {
  return el(`${w}fldSimple`, { [`${w}instr`]: " DATE \\* MERGEFORMAT " }, [
    el(`${w}r`, {}, [el(`${w}t`, { "xml:space": "preserve" }, [], "")]),
  ]);
}

function pageFieldOnlyRun(w: string): XmlElement {
  return el(`${w}fldSimple`, { [`${w}instr`]: " PAGE \\* MERGEFORMAT " }, [
    el(`${w}r`, {}, [el(`${w}t`, { "xml:space": "preserve" }, [], "1")]),
  ]);
}

function galleryParagraph(w: string, align: "left" | "center" | "right" | null, content: XmlElement[]): XmlElement {
  return el(`${w}p`, {}, [
    ...(align ? [el(`${w}pPr`, {}, [el(`${w}jc`, { [`${w}val`]: align })])] : []),
    ...content,
  ]);
}

const TITLE_STYLE = { bold: true, sizePt: 16, color: "2F5597" } as const;
const DATE_STYLE_COLOR = "5B6573";

function presetParagraphs(w: string, preset: HeaderFooterPreset): XmlElement[] {
  switch (preset) {
    case "blank":
      return [galleryParagraph(w, null, [emptyRun(w)])];
    case "centeredTitle":
      return [galleryParagraph(w, "center", [styledTextRun(w, "[Document Title]", TITLE_STYLE)])];
    case "titleAndDate":
      return [
        galleryParagraph(w, "center", [styledTextRun(w, "[Document Title]", TITLE_STYLE)]),
        galleryParagraph(w, "center", [dateFieldRun(w)]),
      ];
    case "threeColumn":
      return [
        galleryParagraph(w, null, [
          styledTextRun(w, "[Company Name]", { sizePt: 9, color: DATE_STYLE_COLOR }),
          tabRun(w),
          styledTextRun(w, "[Document Title]", { sizePt: 9, color: DATE_STYLE_COLOR }),
          tabRun(w),
          pageFieldOnlyRun(w),
        ]),
      ];
  }
}

/** Tracked-node budget per preset: exactly the paragraphs and runs
 * presetParagraphs() builds (a run inside a w:fldSimple still counts — only
 * w:p/w:tbl/w:r are id-tracked). Read by the registered op's `nodeIds`. */
export const HEADER_FOOTER_PRESET_NODE_BUDGET: Record<HeaderFooterPreset, number> = {
  blank: 2, // 1 paragraph + 1 run
  centeredTitle: 2, // 1 paragraph + 1 run
  titleAndDate: 4, // 2 paragraphs + 1 title run + 1 date field's result run
  threeColumn: 6, // 1 paragraph + 2 text runs + 2 tab runs + 1 field result run
};

/** The section governing the document's DEFAULT (last, body-level) sectPr —
 * good enough for a document-scoped preset: every registered document-scope
 * op in this file already treats the document as having one page geometry. */
function defaultSectPr(doc: DocxDocument): XmlElement | undefined {
  const found: XmlElement[] = [];
  const walk = (e: XmlElement): void => {
    if (localName(e.name) === "sectPr") found.push(e);
    else for (const c of e.children) walk(c);
  };
  walk(doc.docRoot);
  return found[found.length - 1];
}

/** The text column width in points, read from the default section's pgSz/
 * pgMar (Word's Normal.dotm defaults — 8.5in page, 1in margins — when the
 * document declares neither). Used to place the three-column preset's
 * center/right tab stops at the same spots Word's own default footer tabs
 * sit (half and full text width). */
function bodyTextWidthPt(doc: DocxDocument): number {
  const sectPr = defaultSectPr(doc);
  const pgSz = sectPr?.children.find((c) => localName(c.name) === "pgSz");
  const pgMar = sectPr?.children.find((c) => localName(c.name) === "pgMar");
  const width = parseInt((pgSz && attr(pgSz, "w")) ?? "12240", 10) || 12240;
  const left = parseInt((pgMar && attr(pgMar, "left")) ?? "1440", 10) || 1440;
  const right = parseInt((pgMar && attr(pgMar, "right")) ?? "1440", 10) || 1440;
  const twips = Math.max(720, width - left - right);
  return twips / 20;
}

/**
 * Replace a header or footer part's content with a preset layout: blank,
 * centered title, title + date, or a three-column line (left/center/right,
 * the right slot a live PAGE field). Three-column sets its tab stops through
 * `setTabStops` — the wave-2 op — rather than reimplementing tab placement.
 * Creates the part on demand, exactly like insertPageNumberPosition.
 */
export function insertHeaderFooterPreset(
  doc: DocxDocument,
  kind: "header" | "footer",
  preset: HeaderFooterPreset,
): boolean {
  const root = doc.ensureHfPart(kind);
  const w = prefixOf(root);
  const paragraphs = presetParagraphs(w, preset);
  root.children = paragraphs;
  if (preset === "threeColumn") {
    const width = bodyTextWidthPt(doc);
    const stops: TabStopSpec[] = [
      { posPt: width / 2, align: "center", leader: "none" },
      { posPt: width, align: "right", leader: "none" },
    ];
    setTabStops(doc, [paragraphs[0]], stops);
  } else {
    doc.refresh();
  }
  return true;
}
