import { XmlElement, attr, intAttr, localName } from "../xml.js";
import { Block } from "../model.js";
import { DocParseContext, parseBlocks } from "./document.js";

/**
 * Parse a footnotes.xml / endnotes.xml part into id → blocks. Separator and
 * continuation pseudo-notes are skipped — the engine draws its own rule.
 *
 * Source references to the part's XML are KEPT so the caret can bind to note
 * text and edits route back into the retained tree (save() re-serializes the
 * part when dirty).
 */
export function parseNotesPart(root: XmlElement, ctx: DocParseContext): Map<number, Block[]> {
  const notes = new Map<number, Block[]>();
  for (const el of root.children) {
    const ln = localName(el.name);
    if (ln !== "footnote" && ln !== "endnote") continue;
    const type = attr(el, "type");
    if (type === "separator" || type === "continuationSeparator" || type === "continuationNotice") continue;
    const id = intAttr(el, "id");
    if (id === undefined) continue;
    notes.set(id, parseBlocks(el, ctx));
  }
  return notes;
}
