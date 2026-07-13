import { XmlElement, attr, intAttr, localName } from "../xml.js";
import { Block } from "../model.js";
import { DocParseContext, parseBlocks } from "./document.js";

/**
 * Parse a footnotes.xml / endnotes.xml part into id → blocks. Separator and
 * continuation pseudo-notes are skipped — the engine draws its own rule.
 *
 * With `editable` (footnotes), source references to the part's XML are KEPT so
 * the caret can bind to footnote text and edits route back into the retained
 * footnotes.xml tree (save() re-serializes it when dirty). Without it
 * (endnotes, still render-only), the sources are stripped so the editor never
 * routes edits into a part save() doesn't yet re-serialize.
 */
export function parseNotesPart(
  root: XmlElement,
  ctx: DocParseContext,
  editable = false,
): Map<number, Block[]> {
  const notes = new Map<number, Block[]>();
  for (const el of root.children) {
    const ln = localName(el.name);
    if (ln !== "footnote" && ln !== "endnote") continue;
    const type = attr(el, "type");
    if (type === "separator" || type === "continuationSeparator" || type === "continuationNotice") continue;
    const id = intAttr(el, "id");
    if (id === undefined) continue;
    const blocks = parseBlocks(el, ctx);
    if (!editable) stripSources(blocks);
    notes.set(id, blocks);
  }
  return notes;
}

function stripSources(blocks: Block[]): void {
  for (const block of blocks) {
    if (block.type === "paragraph") {
      block.src = undefined;
      for (const c of block.children) {
        const runs = c.type === "run" ? [c] : c.runs;
        for (const r of runs) {
          r.src = undefined;
          r.srcParent = undefined;
          for (const rc of r.content) {
            if (rc.kind === "text") rc.srcT = undefined;
            else if (rc.kind === "image") rc.srcDrawing = undefined;
            else if (rc.kind === "anchor" && rc.shape.type === "image") rc.shape.srcDrawing = undefined;
          }
        }
      }
    } else {
      block.src = undefined;
      for (const row of block.rows) for (const cell of row.cells) stripSources(cell.blocks);
    }
  }
}
