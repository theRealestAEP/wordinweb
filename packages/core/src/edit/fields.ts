import { DocxDocument } from "../docx.js";
import { XmlElement, localName } from "../xml.js";

/**
 * Field insertion: dynamic page numbers as `w:fldSimple` elements, which the
 * parser already resolves at layout time (PAGE / NUMPAGES with per-section
 * number formats). Works anywhere a caret can sit — body, header or footer.
 */

function el(name: string, attrs: Record<string, string> = {}, children: XmlElement[] = [], text = ""): XmlElement {
  return { name, attrs, children, text };
}

function cloneDeep(e: XmlElement): XmlElement {
  return { name: e.name, attrs: { ...e.attrs }, children: e.children.map(cloneDeep), text: e.text };
}

/**
 * Insert a dynamic page-number at a text position. `kind` "page" inserts a
 * single PAGE field; "pageOfTotal" inserts the literal runs
 * `Page {PAGE} of {NUMPAGES}`. The destination run's rPr is copied onto the
 * new runs so the field matches the surrounding text.
 */
export function insertPageField(
  doc: DocxDocument,
  t: XmlElement,
  offset: number,
  kind: "page" | "pageOfTotal" = "page",
): boolean {
  const rEl = doc.findParentOf(t);
  const pEl = rEl && doc.findParentOf(rEl);
  if (!rEl || !pEl || localName(rEl.name) !== "r") return false;

  const rw = rEl.name.includes(":") ? rEl.name.slice(0, rEl.name.indexOf(":") + 1) : "";
  const rPr = rEl.children.find((c) => localName(c.name) === "rPr");
  const run = (content: XmlElement) => el(`${rw}r`, {}, [...(rPr ? [cloneDeep(rPr)] : []), content]);
  const textRun = (s: string) => run(el(`${rw}t`, { "xml:space": "preserve" }, [], s));
  const fld = (instr: string) =>
    el(`${rw}fldSimple`, { [`${rw}instr`]: ` ${instr} \\* MERGEFORMAT ` }, [textRun("1")]);

  const inserted: XmlElement[] =
    kind === "page" ? [fld("PAGE")] : [textRun("Page "), fld("PAGE"), textRun(" of "), fld("NUMPAGES")];

  const rIdx = pEl.children.indexOf(rEl);
  if (offset >= t.text.length) {
    pEl.children.splice(rIdx + 1, 0, ...inserted);
  } else if (offset <= 0) {
    pEl.children.splice(rIdx, 0, ...inserted);
  } else {
    const tailT = el(`${rw}t`, { "xml:space": "preserve" }, [], t.text.slice(offset));
    t.text = t.text.slice(0, offset);
    const tail = el(`${rw}r`, {}, [...(rPr ? [cloneDeep(rPr)] : []), tailT]);
    pEl.children.splice(rIdx + 1, 0, ...inserted, tail);
  }
  doc.refresh();
  return true;
}
