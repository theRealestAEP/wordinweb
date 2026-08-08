import { DocxDocument } from "./docx.js";
import { Block, Run } from "./model.js";

/**
 * Word's document statistics (the Word Count dialog / NUMWORDS / NUMCHARS),
 * computed from the parsed body model. One shared rule, consumed by:
 *
 *  - the host API (DocxViewApi.wordCount),
 *  - the field update pass (NUMWORDS / NUMCHARS recompute), and
 *  - layout-time field resolution (the painted text of those fields),
 *
 * so the dialog, the painted field and the written cache cannot disagree.
 *
 * Scope follows Word's defaults: the body only (tables included), not
 * headers, footers, footnotes or comments. A field contributes its cached
 * result text. Tabs and soft line breaks separate words; like spaces they
 * count only in charactersWithSpaces, never in characters. Paragraphs counts
 * paragraphs with any text, Word's dialog rule.
 */
export interface TextStatistics {
  words: number;
  /** Characters excluding whitespace (Word's "characters (no spaces)", the
   * NUMCHARS field value). */
  characters: number;
  /** Characters including spaces and tabs. */
  charactersWithSpaces: number;
  paragraphs: number;
}

export function documentTextStatistics(doc: DocxDocument): TextStatistics {
  const stats: TextStatistics = { words: 0, characters: 0, charactersWithSpaces: 0, paragraphs: 0 };
  const runText = (run: Run): string => {
    let out = "";
    for (const c of run.content) {
      if (c.kind === "text") out += c.text;
      else if (c.kind === "field") out += c.cachedResult;
      else if (c.kind === "tab" || c.kind === "break") out += "\t";
    }
    return out;
  };
  const visit = (blocks: Block[]): void => {
    for (const block of blocks) {
      if (block.type === "paragraph") {
        let text = "";
        for (const child of block.children) {
          for (const run of child.type === "run" ? [child] : child.runs) text += runText(run);
        }
        if (!/\S/.test(text)) continue;
        stats.paragraphs += 1;
        stats.words += text.match(/\S+/g)?.length ?? 0;
        stats.characters += text.replace(/\s/g, "").length;
        stats.charactersWithSpaces += text.length;
      } else {
        for (const row of block.rows) for (const cell of row.cells) visit(cell.blocks);
      }
    }
  };
  for (const section of doc.sections) visit(section.blocks);
  return stats;
}
