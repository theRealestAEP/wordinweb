import bidiFactory from "bidi-js";

const bidi = bidiFactory();

/**
 * Each UTF-16 unit's embedding level (Unicode bidi algorithm) in a run
 * whose base direction is `rtl` or not. Odd levels are drawn right to
 * left. The layout places the run's box and resolves which spans are RTL
 * (TextItem.rtl); the renderer orders the glyphs inside the box.
 */
export function embeddingLevels(text: string, rtl: boolean): Uint8Array {
  return bidi.getEmbeddingLevels(text, rtl ? "rtl" : "ltr").levels;
}

/** The mirrored form of a character at an odd level (a bracket), or null. */
export function mirrored(char: string): string | null {
  return bidi.getMirroredCharacter(char);
}
