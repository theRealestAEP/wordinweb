import { DocxDocument } from "../docx.js";
import { XmlElement, localName } from "../xml.js";
import { paragraphOf } from "./blocks.js";

/**
 * Paragraph-level formatting commands: indent steps and line/paragraph
 * spacing. Like all commands, these mutate the source XML; callers
 * checkpoint history and refresh/relayout afterwards.
 */

const INDENT_STEP_TWIPS = 720; // Word's 0.5in indent step

function prefixOf(e: XmlElement): string {
  return e.name.includes(":") ? e.name.slice(0, e.name.indexOf(":") + 1) : "";
}

function ensurePPr(pEl: XmlElement): XmlElement {
  let pPr = pEl.children.find((c) => localName(c.name) === "pPr");
  if (!pPr) {
    pPr = { name: `${prefixOf(pEl)}pPr`, attrs: {}, children: [], text: "" };
    pEl.children.unshift(pPr);
  }
  return pPr;
}

function attrKey(el: XmlElement, name: string, w: string): string {
  return Object.keys(el.attrs).find((k) => localName(k) === name) ?? `${w}${name}`;
}

function paragraphsOf(doc: DocxDocument, targets: XmlElement[]): Set<XmlElement> {
  const out = new Set<XmlElement>();
  for (const t of targets) {
    const p = paragraphOf(doc, t);
    if (p) out.add(p);
  }
  return out;
}

/** Step the left indent by ±0.5in like Word's indent buttons (floor 0). */
export function adjustIndent(doc: DocxDocument, targets: XmlElement[], direction: 1 | -1): boolean {
  const paragraphs = paragraphsOf(doc, targets);
  if (paragraphs.size === 0) return false;
  let touched = false;
  for (const pEl of paragraphs) {
    const w = prefixOf(pEl);
    const pPr = ensurePPr(pEl);
    let ind = pPr.children.find((c) => localName(c.name) === "ind");
    if (!ind) {
      ind = { name: `${w}ind`, attrs: {}, children: [], text: "" };
      pPr.children.push(ind);
    }
    const key = attrKey(ind, "left", w);
    const cur = parseInt(ind.attrs[key] ?? "0", 10) || 0;
    const next = Math.max(0, cur + direction * INDENT_STEP_TWIPS);
    if (next === cur) continue;
    ind.attrs[key] = String(next);
    touched = true;
  }
  if (touched) doc.refresh();
  return touched;
}

export interface ParagraphSpacingPatch {
  /** Line spacing multiple (1, 1.15, 1.5, 2 …) — w:line auto rule. */
  lineMultiple?: number;
  /** Space before/after in points; null removes the attribute. */
  beforePt?: number | null;
  afterPt?: number | null;
}

/** Set line spacing and/or space before/after on the target paragraphs. */
export function setParagraphSpacing(
  doc: DocxDocument,
  targets: XmlElement[],
  patch: ParagraphSpacingPatch,
): boolean {
  const paragraphs = paragraphsOf(doc, targets);
  if (paragraphs.size === 0) return false;
  for (const pEl of paragraphs) {
    const w = prefixOf(pEl);
    const pPr = ensurePPr(pEl);
    let sp = pPr.children.find((c) => localName(c.name) === "spacing");
    if (!sp) {
      sp = { name: `${w}spacing`, attrs: {}, children: [], text: "" };
      pPr.children.push(sp);
    }
    if (patch.lineMultiple !== undefined) {
      sp.attrs[attrKey(sp, "line", w)] = String(Math.round(patch.lineMultiple * 240));
      sp.attrs[attrKey(sp, "lineRule", w)] = "auto";
    }
    const setPt = (name: "before" | "after", v: number | null | undefined) => {
      if (v === undefined) return;
      const key = attrKey(sp!, name, w);
      if (v === null) delete sp!.attrs[key];
      else sp!.attrs[key] = String(Math.round(v * 20));
    };
    setPt("before", patch.beforePt);
    setPt("after", patch.afterPt);
  }
  doc.refresh();
  return true;
}
