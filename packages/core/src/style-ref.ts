import { DocxDocument } from "./docx.js";
import { Block, FieldContent, Paragraph } from "./model.js";

/**
 * Body STYLEREF resolution: the nearest paragraph of the named style at or
 * before the field, in document order.
 *
 * This is Word's rule for a STYLEREF OUTSIDE a header or footer, where there
 * is no page for the page-relative rule to work from. A header field instead
 * shows the first paragraph of the style that starts on the field's own page
 * (or the last, with \l), which the layout engine resolves from recorded page
 * occurrences — see Engine.prepareStyleRef.
 *
 * It lives here, above both consumers, because the layout engine and the field
 * update pass MUST agree: layout paints what the reader sees, the update pass
 * writes what the saved file carries, and a document whose screen and whose
 * cache disagreed about the same field would be a bug with no right answer.
 * The engine cannot import the update pass (that one already imports layout),
 * so the shared rule is a module of its own rather than a second copy.
 */

/** A paragraph's plain text, for a STYLEREF that names its style. */
function paragraphText(para: Paragraph): string {
  let out = "";
  for (const child of para.children) {
    for (const run of child.type === "run" ? [child] : child.runs) {
      for (const content of run.content) {
        if (content.kind === "text") out += content.text;
        else if (content.kind === "tab") out += "\t";
      }
    }
  }
  return out;
}

/**
 * Resolve every body STYLEREF in the document, keyed by the field occurrence
 * itself — the same `FieldContent` object layout passes as its field key, so
 * a caller looks its answer up by identity and re-breaking a paragraph cannot
 * change it.
 *
 * The style is named by its display name ("Heading 1") or its styleId
 * ("Heading1"), case-insensitively, matching the header resolver in the
 * layout engine. A field with no matching paragraph before it is absent from
 * the map, which every caller reads as "keep the cached result".
 */
export function bodyStyleRefText(doc: DocxDocument): Map<FieldContent, string> {
  const resolved = new Map<FieldContent, string>();
  /** Lower-cased style name AND id → the last paragraph text seen for it. */
  const latest = new Map<string, string>();

  const record = (para: Paragraph): void => {
    const styleId = para.props.styleId ?? doc.styles.defaultParagraphStyle;
    if (!styleId) return;
    const text = paragraphText(para);
    if (!text) return;
    latest.set(styleId.toLowerCase(), text);
    const name = doc.styles.byId.get(styleId)?.name;
    if (name) latest.set(name.toLowerCase(), text);
  };

  const visitBlocks = (blocks: Block[]): void => {
    for (const block of blocks) {
      if (block.type !== "paragraph") {
        for (const row of block.rows) for (const cell of row.cells) visitBlocks(cell.blocks);
        continue;
      }
      // A STYLEREF inside the paragraph resolves against what came BEFORE it,
      // so read the fields first and only then let this paragraph contribute.
      for (const child of block.children) {
        for (const run of child.type === "run" ? [child] : child.runs) {
          for (const content of run.content) {
            if (content.kind !== "field") continue;
            const m = /^STYLEREF\s+(?:"([^"]*)"|([^\s\\]+))/i.exec(content.instruction.trim());
            if (!m) continue;
            const name = (m[1] ?? m[2] ?? "").toLowerCase();
            const text = name ? latest.get(name) : undefined;
            if (text !== undefined) resolved.set(content, text);
          }
        }
      }
      record(block);
    }
  };
  for (const section of doc.sections) visitBlocks(section.blocks);
  return resolved;
}
