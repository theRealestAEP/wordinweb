import { DocxDocument } from "../docx.js";
import { XmlElement, attr, localName } from "../xml.js";

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

// ---------- line numbering (w:lnNumType) ----------

function prefixOf(e: XmlElement): string {
  return e.name.includes(":") ? e.name.slice(0, e.name.indexOf(":") + 1) : "";
}

/** Canonical CT_SectPr child order (subset we touch), for schema-correct
 * insertion — Word repairs a sectPr whose children are out of order. */
const SECTPR_ORDER = [
  "headerReference", "footerReference", "footnotePr", "endnotePr", "type", "pgSz", "pgMar",
  "paperSrc", "pgBorders", "lnNumType", "pgNumType", "cols", "formProt", "vAlign", "noEndnote",
  "titlePg", "textDirection", "bidi", "rtlGutter", "docGrid", "printerSettings", "sectPrChange",
];

/** Insert `child` into `sectPr` at its schema-ordered position. */
function insertInOrder(sectPr: XmlElement, childEl: XmlElement): void {
  const rank = (e: XmlElement) => {
    const i = SECTPR_ORDER.indexOf(localName(e.name));
    return i === -1 ? SECTPR_ORDER.length : i;
  };
  const r = rank(childEl);
  const at = sectPr.children.findIndex((c) => rank(c) > r);
  if (at === -1) sectPr.children.push(childEl);
  else sectPr.children.splice(at, 0, childEl);
}

function allSectPrs(doc: DocxDocument): XmlElement[] {
  const out: XmlElement[] = [];
  const root = doc.editableRoots()[0];
  const walk = (e: XmlElement) => {
    if (localName(e.name) === "sectPr") out.push(e);
    for (const c of e.children) walk(c);
  };
  if (root) walk(root);
  return out;
}

export interface LineNumberingPatch {
  /** Turn margin line numbering on/off for the target section(s). */
  enabled: boolean;
  /** Number every Nth line (1/5/10). Default 1 when first enabled. */
  countBy?: number;
  /** When the count resets. Default newPage when first enabled. */
  restart?: "continuous" | "newPage" | "newSection";
  /** First line number. Default 1. */
  start?: number;
}

/** Current line-numbering settings for the section governing `t`, or null when
 * the section has none (line numbering off). */
export function lineNumberingAt(
  doc: DocxDocument,
  t: XmlElement,
): { countBy: number; restart: "continuous" | "newPage" | "newSection"; start: number } | null {
  const sectPr = sectPrAt(doc, t);
  const ln = sectPr?.children.find((c) => localName(c.name) === "lnNumType");
  if (!ln) return null;
  const restart = attr(ln, "restart");
  return {
    countBy: parseInt(attr(ln, "countBy") ?? "1", 10) || 1,
    restart: restart === "continuous" || restart === "newSection" ? restart : "newPage",
    start: parseInt(attr(ln, "start") ?? "1", 10) || 1,
  };
}

/**
 * Toggle/configure margin line numbering. With `target` (a sectPr), only that
 * section changes; otherwise every section in the document updates. Disabling
 * removes the w:lnNumType element entirely (Word's "None"). Enabling creates
 * or updates it in schema-correct position and relayouts.
 */
export function setLineNumbering(doc: DocxDocument, patch: LineNumberingPatch, target?: XmlElement): boolean {
  const sectPrs = target ? [target] : allSectPrs(doc);
  if (sectPrs.length === 0) return false;
  for (const sectPr of sectPrs) {
    const w = prefixOf(sectPr) || "w:";
    if (!patch.enabled) {
      sectPr.children = sectPr.children.filter((c) => localName(c.name) !== "lnNumType");
      continue;
    }
    let ln = sectPr.children.find((c) => localName(c.name) === "lnNumType");
    const creating = !ln;
    if (!ln) {
      ln = el(`${w}lnNumType`, {});
      insertInOrder(sectPr, ln);
    }
    const setA = (local: string, v: string) => {
      const key = Object.keys(ln!.attrs).find((k) => localName(k) === local) ?? `${w}${local}`;
      ln!.attrs[key] = v;
    };
    const delA = (local: string) => {
      const key = Object.keys(ln!.attrs).find((k) => localName(k) === local);
      if (key) delete ln!.attrs[key];
    };
    // countBy always present (a bare lnNumType already means "every line").
    if (patch.countBy !== undefined) setA("countBy", String(patch.countBy));
    else if (creating) setA("countBy", "1");
    if (patch.restart !== undefined) {
      // newPage is the OOXML default; keep the XML minimal by omitting it.
      if (patch.restart === "newPage") delA("restart");
      else setA("restart", patch.restart);
    }
    if (patch.start !== undefined) {
      if (patch.start === 1) delA("start");
      else setA("start", String(patch.start));
    }
  }
  doc.refresh();
  return true;
}
