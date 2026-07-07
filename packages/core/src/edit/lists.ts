import { DocxDocument } from "../docx.js";
import { XmlElement, attr, localName } from "../xml.js";
import { paragraphOf } from "./blocks.js";

/**
 * List edit commands: turn paragraphs into bulleted/numbered list items and
 * back. Definitions are added to numbering.xml on demand (the part itself is
 * created when the document has none). Like all commands, these mutate the
 * source XML; callers checkpoint history and refresh/relayout afterwards.
 */

export type ListKind = "bullet" | "number";

function el(name: string, attrs: Record<string, string> = {}, children: XmlElement[] = [], text = ""): XmlElement {
  return { name, attrs, children, text };
}

function prefixOf(e: XmlElement): string {
  return e.name.includes(":") ? e.name.slice(0, e.name.indexOf(":") + 1) : "";
}

/** The list kind of the paragraph containing `target` (direct numPr only). */
export function listTypeAt(doc: DocxDocument, target: XmlElement): ListKind | null {
  const pEl = paragraphOf(doc, target);
  if (!pEl) return null;
  const pPr = pEl.children.find((c) => localName(c.name) === "pPr");
  const numPr = pPr?.children.find((c) => localName(c.name) === "numPr");
  if (!numPr) return null;
  const numId = numPr.children.find((c) => localName(c.name) === "numId");
  const ilvl = numPr.children.find((c) => localName(c.name) === "ilvl");
  const id = parseInt(attr(numId ?? { name: "", attrs: {}, children: [], text: "" }, "val") ?? "", 10);
  if (!Number.isFinite(id)) return null;
  const lvl = doc.numberingLevel(id, parseInt(attr(ilvl ?? { name: "", attrs: {}, children: [], text: "" }, "val") ?? "0", 10) || 0);
  return lvl?.format === "bullet" ? "bullet" : "number";
}

/**
 * Find (or create) a single-level numbering definition of the given kind and
 * return its numId. Reuses a matching definition when the file already has
 * one so repeated toggles don't pile up abstractNums.
 */
function ensureListNum(doc: DocxDocument, kind: ListKind): number | null {
  for (const [numId, inst] of doc.numbering.instances) {
    const lvl = doc.numberingLevel(numId, 0);
    if (!lvl) continue;
    if (kind === "bullet" ? lvl.format === "bullet" : lvl.format === "decimal") return numId;
    void inst;
  }
  const root = doc.numberingTree(true);
  if (!root) return null;
  const w = prefixOf(root);
  let maxAbs = -1;
  let maxNum = 0;
  for (const c of root.children) {
    if (localName(c.name) === "abstractNum") {
      maxAbs = Math.max(maxAbs, parseInt(attr(c, "abstractNumId") ?? "-1", 10));
    } else if (localName(c.name) === "num") {
      maxNum = Math.max(maxNum, parseInt(attr(c, "numId") ?? "0", 10));
    }
  }
  const absId = String(maxAbs + 1);
  const numId = maxNum + 1;
  // Nine levels like Word's defaults, so Tab/Shift-Tab demotion has labels.
  const BULLETS = ["•", "○", "■"];
  const levels = Array.from({ length: 9 }, (_, i) =>
    el(`${w}lvl`, { [`${w}ilvl`]: String(i) }, [
      el(`${w}start`, { [`${w}val`]: "1" }),
      el(`${w}numFmt`, { [`${w}val`]: kind === "bullet" ? "bullet" : i % 3 === 1 ? "lowerLetter" : i % 3 === 2 ? "lowerRoman" : "decimal" }),
      el(`${w}lvlText`, { [`${w}val`]: kind === "bullet" ? BULLETS[i % 3] : `%${i + 1}.` }),
      el(`${w}lvlJc`, { [`${w}val`]: "left" }),
      el(`${w}pPr`, {}, [el(`${w}ind`, { [`${w}left`]: String(720 * (i + 1)), [`${w}hanging`]: "360" })]),
    ]),
  );
  // abstractNum elements must precede num elements in the part.
  let insertAt = root.children.length;
  for (let i = 0; i < root.children.length; i++) {
    if (localName(root.children[i].name) === "num") {
      insertAt = i;
      break;
    }
  }
  root.children.splice(insertAt, 0, el(`${w}abstractNum`, { [`${w}abstractNumId`]: absId }, levels));
  root.children.push(
    el(`${w}num`, { [`${w}numId`]: String(numId) }, [el(`${w}abstractNumId`, { [`${w}val`]: absId })]),
  );
  doc.markNumberingChanged();
  return numId;
}

/** Step the list level of the target paragraphs (Tab / Shift-Tab), 0..8. */
export function setListLevel(doc: DocxDocument, targets: XmlElement[], delta: 1 | -1): boolean {
  const paragraphs = new Set<XmlElement>();
  for (const t of targets) {
    const p = paragraphOf(doc, t);
    if (p) paragraphs.add(p);
  }
  let touched = false;
  for (const pEl of paragraphs) {
    const pPr = pEl.children.find((c) => localName(c.name) === "pPr");
    const numPr = pPr?.children.find((c) => localName(c.name) === "numPr");
    if (!numPr) continue;
    const w = prefixOf(pEl);
    let ilvl = numPr.children.find((c) => localName(c.name) === "ilvl");
    if (!ilvl) {
      ilvl = el(`${w}ilvl`, { [`${w}val`]: "0" });
      numPr.children.unshift(ilvl);
    }
    const key = Object.keys(ilvl.attrs).find((k) => localName(k) === "val") ?? `${w}val`;
    const cur = parseInt(ilvl.attrs[key] ?? "0", 10) || 0;
    const next = Math.min(8, Math.max(0, cur + delta));
    if (next === cur) continue;
    ilvl.attrs[key] = String(next);
    touched = true;
  }
  if (touched) doc.refresh();
  return touched;
}

/**
 * Make the paragraphs containing `targets` list items of the given kind, or
 * plain paragraphs again with kind=null. Returns false when nothing changed.
 */
export function setListType(doc: DocxDocument, targets: XmlElement[], kind: ListKind | null): boolean {
  const paragraphs = new Set<XmlElement>();
  for (const t of targets) {
    const p = paragraphOf(doc, t);
    if (p) paragraphs.add(p);
  }
  if (paragraphs.size === 0) return false;

  const numId = kind === null ? null : ensureListNum(doc, kind);
  if (kind !== null && numId === null) return false;

  let touched = false;
  for (const pEl of paragraphs) {
    const w = prefixOf(pEl);
    let pPr = pEl.children.find((c) => localName(c.name) === "pPr");
    if (kind === null) {
      if (!pPr) continue;
      const before = pPr.children.length;
      pPr.children = pPr.children.filter((c) => localName(c.name) !== "numPr");
      touched = touched || pPr.children.length !== before;
      continue;
    }
    if (!pPr) {
      pPr = el(`${w}pPr`);
      pEl.children.unshift(pPr);
    }
    const numPr = el(`${w}numPr`, {}, [
      el(`${w}ilvl`, { [`${w}val`]: "0" }),
      el(`${w}numId`, { [`${w}val`]: String(numId) }),
    ]);
    const existing = pPr.children.findIndex((c) => localName(c.name) === "numPr");
    if (existing !== -1) {
      pPr.children[existing] = numPr;
    } else {
      // Schema order: numPr follows pStyle (and the keep/pageBreak flags).
      const styleIdx = pPr.children.findIndex((c) => localName(c.name) === "pStyle");
      pPr.children.splice(styleIdx + 1, 0, numPr);
    }
    touched = true;
  }
  if (touched) doc.refresh();
  return touched;
}
