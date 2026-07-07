import { DocxDocument } from "../docx.js";
import { XmlElement, localName } from "../xml.js";

/**
 * Section editing: resolve the sectPr governing a caret position, insert
 * page/column/section breaks. A section's sectPr lives either embedded in
 * the pPr of its LAST paragraph or (for the final section) at body level.
 */

function el(name: string, attrs: Record<string, string> = {}, children: XmlElement[] = [], text = ""): XmlElement {
  return { name, attrs, children, text };
}

function cloneDeep(e: XmlElement): XmlElement {
  return { name: e.name, attrs: { ...e.attrs }, children: e.children.map(cloneDeep), text: e.text };
}

function bodyOf(doc: DocxDocument): XmlElement | null {
  const root = doc.editableRoots()[0];
  const find = (e: XmlElement): XmlElement | null => {
    if (localName(e.name) === "body") return e;
    for (const c of e.children) {
      const f = find(c);
      if (f) return f;
    }
    return null;
  };
  return root ? find(root) : null;
}

function containsEl(root: XmlElement, target: XmlElement): boolean {
  if (root === target) return true;
  for (const c of root.children) if (containsEl(c, target)) return true;
  return false;
}

/**
 * The sectPr that governs the section containing `t` (or the body's final
 * sectPr when the caret sits in the last section).
 */
export function sectPrAt(doc: DocxDocument, t: XmlElement): XmlElement | null {
  const body = bodyOf(doc);
  if (!body) return null;
  let seen = false;
  let bodySectPr: XmlElement | null = null;
  for (const child of body.children) {
    const ln = localName(child.name);
    if (ln === "sectPr") {
      bodySectPr = child;
      continue;
    }
    if (!seen && containsEl(child, t)) seen = true;
    if (seen && ln === "p") {
      const pPr = child.children.find((c) => localName(c.name) === "pPr");
      const sp = pPr?.children.find((c) => localName(c.name) === "sectPr");
      if (sp) return sp;
    }
  }
  return seen ? bodySectPr : bodySectPr;
}

/** Insert a page or column break at the caret (splits the run). */
export function insertBreakAt(doc: DocxDocument, t: XmlElement, offset: number, kind: "page" | "column"): boolean {
  const rEl = doc.findParentOf(t);
  const pEl = rEl && doc.findParentOf(rEl);
  if (!rEl || !pEl || localName(rEl.name) !== "r") return false;
  const rw = rEl.name.includes(":") ? rEl.name.slice(0, rEl.name.indexOf(":") + 1) : "";
  const rPr = rEl.children.find((c) => localName(c.name) === "rPr");
  const brRun = el(`${rw}r`, {}, [
    ...(rPr ? [cloneDeep(rPr)] : []),
    el(`${rw}br`, { [`${rw}type`]: kind }),
  ]);
  const rIdx = pEl.children.indexOf(rEl);
  if (offset >= t.text.length) {
    pEl.children.splice(rIdx + 1, 0, brRun);
  } else if (offset <= 0) {
    pEl.children.splice(rIdx, 0, brRun);
  } else {
    const tailT = el(`${rw}t`, { "xml:space": "preserve" }, [], t.text.slice(offset));
    t.text = t.text.slice(0, offset);
    const tail = el(`${rw}r`, {}, [...(rPr ? [cloneDeep(rPr)] : []), tailT]);
    pEl.children.splice(rIdx + 1, 0, brRun, tail);
  }
  doc.refresh();
  return true;
}

/**
 * Insert a section break after the caret's paragraph: the paragraphs before
 * the break become their own section (an embedded sectPr cloned from the
 * governing one, sans type), and the FOLLOWING section's sectPr carries the
 * chosen break type - OOXML puts "how do I start" on the section after the
 * break.
 */
export function insertSectionBreak(
  doc: DocxDocument,
  t: XmlElement,
  type: "nextPage" | "continuous",
): boolean {
  const body = bodyOf(doc);
  const governing = sectPrAt(doc, t);
  if (!body || !governing) return false;
  const w = governing.name.includes(":") ? governing.name.slice(0, governing.name.indexOf(":") + 1) : "w:";

  // The paragraph (direct body child) containing the caret.
  let host: XmlElement | null = null;
  for (const child of body.children) {
    if (containsEl(child, t)) {
      host = child;
      break;
    }
  }
  if (!host) return false;

  const closing = cloneDeep(governing);
  closing.children = closing.children.filter((c) => !["headerReference", "footerReference"].includes(localName(c.name)));
  closing.children = closing.children.filter((c) => localName(c.name) !== "type");

  // Set the break type on the governing (following) sectPr.
  let typeEl = governing.children.find((c) => localName(c.name) === "type");
  if (!typeEl) {
    typeEl = el(`${w}type`, {}, []);
    governing.children.unshift(typeEl);
  }
  const typeKey = Object.keys(typeEl.attrs).find((k) => localName(k) === "val") ?? `${w}val`;
  typeEl.attrs[typeKey] = type;

  const breakPara = el(`${w}p`, {}, [el(`${w}pPr`, {}, [closing])]);
  body.children.splice(body.children.indexOf(host) + 1, 0, breakPara);
  doc.refresh();
  return true;
}
