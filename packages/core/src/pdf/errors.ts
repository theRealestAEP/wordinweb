/** Why a render failed, for a caller to act on without reading the message. */
export type PdfRenderErrorCode =
  /** Text in a script the renderer cannot shape (Arabic, Indic, Thai, ...);
   * the message names the script and where it first appears. */
  | "unsupported_script"
  /** The input is not a .docx: not a zip, or a zip with no Word document. */
  | "not_docx"
  /** The input is a .docx that cannot be read: a corrupt entry, or XML that is not well formed. */
  | "damaged"
  /** Anything else. The message and `cause` say what. */
  | "failed";

export class PdfRenderError extends Error {
  readonly cause?: unknown;

  constructor(
    readonly code: PdfRenderErrorCode,
    message: string,
    options?: { cause?: unknown },
  ) {
    super(message);
    this.name = "PdfRenderError";
    this.cause = options?.cause;
  }
}

/** The PdfRenderError for an error thrown while opening or laying out a document. */
export function asPdfRenderError(error: unknown): PdfRenderError {
  if (error instanceof PdfRenderError) return error;
  const message = error instanceof Error ? error.message : String(error);
  // stream/zip.ts and stream/body.ts say these for a file that is no .docx ...
  if (/not a \.docx file|^missing |encrypted entries/.test(message)) return new PdfRenderError("not_docx", message, { cause: error });
  // ... and these (or fflate's inflate errors, or a read past the end) for one that is but cannot be read.
  if (/is corrupt|bad (central directory|local header|ZIP64)|compression method|XML|unterminated|unexpected|invalid |outside the bounds/i.test(message)) return new PdfRenderError("damaged", message, { cause: error });
  return new PdfRenderError("failed", message, { cause: error });
}

/**
 * Scripts whose text needs a shaper (joining forms, reordering, cluster
 * composition) that this renderer does not have; each with the Unicode
 * blocks that hold it. Hebrew is not here: its letters do not join, and
 * its points are zero-width marks a font places beside the letter before.
 */
const UNSUPPORTED_SCRIPTS: [name: string, from: number, to: number][] = [
  ["Arabic", 0x0600, 0x06ff],
  ["Syriac", 0x0700, 0x074f],
  ["Arabic", 0x0750, 0x077f],
  ["Thaana", 0x0780, 0x07bf],
  ["N'Ko", 0x07c0, 0x07ff],
  ["Arabic", 0x08a0, 0x08ff],
  ["Devanagari", 0x0900, 0x097f],
  ["Bengali", 0x0980, 0x09ff],
  ["Gurmukhi", 0x0a00, 0x0a7f],
  ["Gujarati", 0x0a80, 0x0aff],
  ["Oriya", 0x0b00, 0x0b7f],
  ["Tamil", 0x0b80, 0x0bff],
  ["Telugu", 0x0c00, 0x0c7f],
  ["Kannada", 0x0c80, 0x0cff],
  ["Malayalam", 0x0d00, 0x0d7f],
  ["Sinhala", 0x0d80, 0x0dff],
  ["Thai", 0x0e00, 0x0e7f],
  ["Lao", 0x0e80, 0x0eff],
  ["Tibetan", 0x0f00, 0x0fff],
  ["Myanmar", 0x1000, 0x109f],
  ["Khmer", 0x1780, 0x17ff],
  ["Mongolian", 0x1800, 0x18af],
  ["Limbu", 0x1900, 0x194f],
  ["Tai Le", 0x1950, 0x197f],
  ["New Tai Lue", 0x1980, 0x19df],
  ["Khmer", 0x19e0, 0x19ff],
  ["Buginese", 0x1a00, 0x1a1f],
  ["Tai Tham", 0x1a20, 0x1aaf],
  ["Balinese", 0x1b00, 0x1b7f],
  ["Sundanese", 0x1b80, 0x1bbf],
  ["Batak", 0x1bc0, 0x1bff],
  ["Lepcha", 0x1c00, 0x1c4f],
  ["Devanagari", 0xa8e0, 0xa8ff],
  ["Myanmar", 0xa9e0, 0xa9ff],
  ["Javanese", 0xa980, 0xa9df],
  ["Myanmar", 0xaa60, 0xaa7f],
  ["Tai Viet", 0xaa80, 0xaadf],
  ["Arabic", 0xfb50, 0xfdff],
  ["Arabic", 0xfe70, 0xfeff],
];

/** The name of the first unsupported script `text` holds, if any. */
export function unsupportedScript(text: string): string | undefined {
  for (const ch of text) {
    const cp = ch.codePointAt(0) as number;
    if (cp < 0x0600) continue;
    for (const [name, from, to] of UNSUPPORTED_SCRIPTS) if (cp >= from && cp <= to) return name;
  }
  return undefined;
}
