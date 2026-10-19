declare module "bidi-js" {
  interface EmbeddingLevels {
    levels: Uint8Array;
    paragraphs: { start: number; end: number; level: number }[];
  }

  interface Bidi {
    getEmbeddingLevels(text: string, direction?: "ltr" | "rtl"): EmbeddingLevels;
    getReorderedString(text: string, levels: EmbeddingLevels, start?: number, end?: number): string;
    getMirroredCharactersMap(text: string, levels: EmbeddingLevels, start?: number, end?: number): Map<number, string>;
    /** The mirrored form of a character (a bracket), or null when it has none. */
    getMirroredCharacter(char: string): string | null;
  }

  function bidiFactory(): Bidi;

  export default bidiFactory;
}
