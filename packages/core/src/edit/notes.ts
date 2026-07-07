import { DocxDocument } from "../docx.js";
import { XmlElement, attr, localName } from "../xml.js";

/**
 * Footnote insertion. The reference run goes at the caret; the note body
 * goes into footnotes.xml (created on demand). Note numbering is
 * document-order derived at layout, so no renumbering is needed here.
 */

function el(name: string, attrs: Record<string, string> = {}, children: XmlElement[] = [], text = ""): XmlElement {
  return { name, attrs, children, text };
}

/**
 * Insert a footnote whose reference sits at `offset` inside the w:t `t`.
 * Returns the new footnote id, or null.
 */
export function insertFootnote(doc: DocxDocument, t: XmlElement, offset: number, text: string): number | null {
  if (!text.trim()) return null;
  const notesRoot = doc.footnotesTree(true);
  if (!notesRoot) return null;
  const w = notesRoot.name.includes(":") ? notesRoot.name.slice(0, notesRoot.name.indexOf(":") + 1) : "";

  const rEl = doc.findParentOf(t);
  const pEl = rEl && doc.findParentOf(rEl);
  if (!rEl || !pEl || localName(rEl.name) !== "r") return null;

  // Fresh id: separators use -1/0, so real notes start at 1.
  let maxId = 0;
  for (const c of notesRoot.children) {
    if (localName(c.name) !== "footnote") continue;
    const id = parseInt(attr(c, "id") ?? "0", 10);
    if (Number.isFinite(id)) maxId = Math.max(maxId, id);
  }
  const newId = maxId + 1;

  // Split the run at the caret so the reference lands mid-text correctly.
  const rw = rEl.name.includes(":") ? rEl.name.slice(0, rEl.name.indexOf(":") + 1) : "";
  const rPr = rEl.children.find((c) => localName(c.name) === "rPr");
  const refRun = el(`${rw}r`, {}, [
    el(`${rw}rPr`, {}, [el(`${rw}rStyle`, { [`${rw}val`]: "FootnoteReference" })]),
    el(`${rw}footnoteReference`, { [`${rw}id`]: String(newId) }),
  ]);
  const rIdx = pEl.children.indexOf(rEl);
  if (offset >= t.text.length) {
    pEl.children.splice(rIdx + 1, 0, refRun);
  } else if (offset <= 0) {
    pEl.children.splice(rIdx, 0, refRun);
  } else {
    const tailT = el(`${rw}t`, { "xml:space": "preserve" }, [], t.text.slice(offset));
    t.text = t.text.slice(0, offset);
    const tail = el(`${rw}r`, {}, [...(rPr ? [cloneShallow(rPr)] : []), tailT]);
    pEl.children.splice(rIdx + 1, 0, refRun, tail);
  }

  // Note body: FootnoteText paragraph, own mark, then the text.
  notesRoot.children.push(
    el(`${w}footnote`, { [`${w}id`]: String(newId) }, [
      el(`${w}p`, {}, [
        el(`${w}pPr`, {}, [el(`${w}pStyle`, { [`${w}val`]: "FootnoteText" })]),
        el(`${w}r`, {}, [
          el(`${w}rPr`, {}, [el(`${w}rStyle`, { [`${w}val`]: "FootnoteReference" })]),
          el(`${w}footnoteRef`),
        ]),
        el(`${w}r`, {}, [el(`${w}t`, { "xml:space": "preserve" }, [], " " + text.trim())]),
      ]),
    ]),
  );

  doc.markFootnotesChanged();
  doc.refresh();
  return newId;
}

function cloneShallow(e: XmlElement): XmlElement {
  return { name: e.name, attrs: { ...e.attrs }, children: e.children.map(cloneShallow), text: e.text };
}
