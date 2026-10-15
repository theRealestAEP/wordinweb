import bidiFactory from "bidi-js";

const bidi = bidiFactory();

/**
 * A right-to-left run's characters in the order they are drawn from left
 * to right (Unicode bidi algorithm, base direction RTL), brackets mirrored.
 * The layout places the run's box; the browser reorders inside it with
 * direction: rtl, and this does the same for the PDF.
 */
export function visualOrder(text: string): string {
  const levels = bidi.getEmbeddingLevels(text, "rtl");
  const mirrored = bidi.getMirroredCharactersMap(text, levels);
  let logical = text;
  if (mirrored.size > 0) {
    const chars = text.split("");
    for (const [index, char] of mirrored) chars[index] = char;
    logical = chars.join("");
  }
  return bidi.getReorderedString(logical, levels);
}
