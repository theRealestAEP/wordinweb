declare module "bidi-js" {
  interface EmbeddingLevels {
    levels: Uint8Array;
    paragraphs: { start: number; end: number; level: number }[];
  }

  interface Bidi {
    getEmbeddingLevels(text: string, direction?: "ltr" | "rtl"): EmbeddingLevels;
    getReorderedString(text: string, levels: EmbeddingLevels, start?: number, end?: number): string;
    getMirroredCharactersMap(text: string, levels: EmbeddingLevels, start?: number, end?: number): Map<number, string>;
  }

  function bidiFactory(): Bidi;

  export default bidiFactory;
}
