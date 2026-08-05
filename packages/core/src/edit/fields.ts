import { DocxDocument } from "../docx.js";
import { XmlElement, localName } from "../xml.js";

/** Field insertion as `w:fldSimple`, which the parser resolves at layout time. */

function el(name: string, attrs: Record<string, string> = {}, children: XmlElement[] = [], text = ""): XmlElement {
  return { name, attrs, children, text };
}

function cloneDeep(e: XmlElement): XmlElement {
  return { name: e.name, attrs: { ...e.attrs }, children: e.children.map(cloneDeep), text: e.text };
}

function insertElementsAt(
  doc: DocxDocument,
  t: XmlElement,
  offset: number,
  inserted: XmlElement[],
): boolean {
  const rEl = doc.findParentOf(t);
  const parent = rEl && doc.findParentOf(rEl);
  if (!rEl || !parent || localName(rEl.name) !== "r") return false;

  const w = rEl.name.includes(":") ? rEl.name.slice(0, rEl.name.indexOf(":") + 1) : "";
  const rPr = rEl.children.find((c) => localName(c.name) === "rPr");
  const rIdx = parent.children.indexOf(rEl);
  const tIdx = rEl.children.indexOf(t);
  if (rIdx < 0 || tIdx < 0) return false;
  const at = Math.max(0, Math.min(offset, t.text.length));
  const makeText = (text: string): XmlElement =>
    el(t.name, { ...t.attrs, "xml:space": "preserve" }, [], text);
  const makeRun = (content: XmlElement[]): XmlElement =>
    el(rEl.name, { ...rEl.attrs }, [...(rPr ? [cloneDeep(rPr)] : []), ...content]);
  const before = rEl.children.slice(0, tIdx).filter((child) => localName(child.name) !== "rPr");
  const after = rEl.children.slice(tIdx + 1).filter((child) => localName(child.name) !== "rPr");
  let beforeRun: XmlElement | null;
  let afterRun: XmlElement | null;
  if (at === 0) {
    beforeRun = before.length > 0 ? makeRun(before) : null;
    rEl.children = [...(rPr ? [rPr] : []), t, ...after];
    afterRun = rEl;
  } else {
    const tail = at < t.text.length ? makeText(t.text.slice(at)) : null;
    t.text = t.text.slice(0, at);
    rEl.children = [...(rPr ? [rPr] : []), ...before, t];
    beforeRun = rEl;
    // A field is atomic display content, so the editor cannot place a caret
    // inside it. Keep a real text anchor after a field inserted at line end.
    afterRun = makeRun([...(tail ? [tail] : [makeText("")]), ...after]);
  }
  parent.children.splice(
    rIdx,
    1,
    ...(beforeRun ? [beforeRun] : []),
    ...inserted,
    ...(afterRun ? [afterRun] : []),
  );
  doc.refresh();
  return true;
}

/**
 * Field instructions this engine will INSERT, by keyword.
 *
 * An arbitrary instruction is an injection surface, not a feature: Word's
 * INCLUDETEXT and INCLUDEPICTURE read a path or URL the document names,
 * DDE/DDEAUTO/LINK open a channel to another application, and MACROBUTTON
 * names code to run. A document that ARRIVES holding one still renders — the
 * parser keeps every instruction and unsupported ones paint their cached
 * result — but nothing this editor writes may create one. Same posture
 * validatePastedOoxml takes toward pasted markup.
 *
 * The list is the field types that read nothing outside the document: what the
 * engine evaluates itself (layout/inline.ts resolveField, edit/update-fields.ts)
 * plus the document-property fields, which are inert text.
 */
const INSERTABLE_FIELD_KEYWORDS = new Set([
  "PAGE", "NUMPAGES", "SECTIONPAGES", "SECTION", "DATE", "TIME",
  "CREATEDATE", "SAVEDATE", "PRINTDATE", "AUTHOR", "TITLE", "SUBJECT",
  "KEYWORDS", "COMMENTS", "FILENAME", "NUMWORDS", "NUMCHARS", "PAGEREF",
  "REF", "SEQ", "STYLEREF", "TOC", "INDEX", "LISTNUM", "QUOTE",
]);

/** Longest instruction this engine writes. Real ones are far shorter; the cap
 * bounds what a remote intent can push through the wire into a w:instrText. */
export const MAX_FIELD_INSTRUCTION = 200;

/**
 * Whether `instruction` may be inserted.
 *
 * The collaborative sequencer validates inbound intents with this same
 * predicate, so the local path and the wire path allow exactly the same set —
 * two lists that drifted would mean an instruction one host writes and another
 * refuses.
 */
export function isInsertableFieldInstruction(instruction: string): boolean {
  if (typeof instruction !== "string") return false;
  if (instruction.length === 0 || instruction.length > MAX_FIELD_INSTRUCTION) return false;
  // Printable ASCII only. A control or format character inside a w:instrText
  // can split one instruction into two when Word re-reads it, smuggling a
  // second keyword past the check above.
  if (!/^[\x20-\x7e]+$/.test(instruction)) return false;
  const keyword = instruction.trim().split(/\s+/)[0]?.toUpperCase();
  return !!keyword && INSERTABLE_FIELD_KEYWORDS.has(keyword);
}

/** Insert a field instruction at a text position, preserving surrounding run formatting. */
export function insertField(
  doc: DocxDocument,
  t: XmlElement,
  offset: number,
  instruction: string,
  cachedResult = "",
): boolean {
  if (!isInsertableFieldInstruction(instruction)) return false;
  const clean = instruction.trim();
  const rEl = doc.findParentOf(t);
  if (!rEl || localName(rEl.name) !== "r") return false;
  const w = rEl.name.includes(":") ? rEl.name.slice(0, rEl.name.indexOf(":") + 1) : "";
  const rPr = rEl.children.find((c) => localName(c.name) === "rPr");
  const resultRun = el(`${w}r`, {}, [
    ...(rPr ? [cloneDeep(rPr)] : []),
    el(`${w}t`, { "xml:space": "preserve" }, [], cachedResult),
  ]);
  const field = el(`${w}fldSimple`, { [`${w}instr`]: ` ${clean} ` }, [resultRun]);
  return insertElementsAt(doc, t, offset, [field]);
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
  if (!rEl || localName(rEl.name) !== "r") return false;

  const rw = rEl.name.includes(":") ? rEl.name.slice(0, rEl.name.indexOf(":") + 1) : "";
  const rPr = rEl.children.find((c) => localName(c.name) === "rPr");
  const run = (content: XmlElement) => el(`${rw}r`, {}, [...(rPr ? [cloneDeep(rPr)] : []), content]);
  const textRun = (s: string) => run(el(`${rw}t`, { "xml:space": "preserve" }, [], s));
  const fld = (instr: string) =>
    el(`${rw}fldSimple`, { [`${rw}instr`]: ` ${instr} \\* MERGEFORMAT ` }, [textRun("1")]);

  const inserted: XmlElement[] =
    kind === "page" ? [fld("PAGE")] : [textRun("Page "), fld("PAGE"), textRun(" of "), fld("NUMPAGES")];
  return insertElementsAt(doc, t, offset, inserted);
}

/** Insert a live DATE or TIME field using Word's date-picture syntax. */
export function insertDateTimeField(
  doc: DocxDocument,
  t: XmlElement,
  offset: number,
  kind: "date" | "time",
  picture: string,
): boolean {
  const safePicture = picture.replace(/"/g, "");
  return insertField(doc, t, offset, `${kind.toUpperCase()} \\@ "${safePicture}" \\* MERGEFORMAT`);
}
