import { shapedScriptOf } from "./scripts.js";
import { shaperLoaded } from "./shaper.js";

/** Why a render failed, for a caller to act on without reading the message. */
export type PdfRenderErrorCode =
  /** Text in a script the renderer cannot shape: Mongolian; or, without
   * the harfbuzzjs package, Arabic, Indic, Thai and the other scripts
   * that need a shaper. The message names the script and where it first
   * appears. */
  | "unsupported_script"
  /** The input is not a .docx: not a zip, or a zip with no Word document. */
  | "not_docx"
  /** The input is another format this renderer does not read, which the
   * message names: a password-protected document, an RTF or PDF file, an
   * OpenDocument file, a Word 6.0/95 file. Nothing of it was read past its
   * first bytes (a Word 97-2003 .doc is read: doc/convert.ts). */
  | "unsupported_format"
  /** The input is a .docx or .doc that cannot be read: a corrupt entry, XML
   * that is not well formed, a .doc whose tables are corrupt. */
  | "damaged"
  /** The document is over the size limits (DOCUMENT_LIMITS and
   * DEFAULT_ZIP_LIMITS in zip.ts), which the message names: found from the
   * zip's directory before anything is inflated, or counted from the XML
   * before it is parsed or as the body streams. */
  | "too_large"
  /** The document holds content Word prints that this renderer does not
   * read, which the message names: a w:altChunk (a part in another
   * format, such as HTML, that Word imports on open) or a w:subDoc (a
   * subdocument Word reads from another file). The document is refused
   * rather than printed without it. */
  | "unsupported_content"
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
  // stream/body.ts says this for a file of another format it knows ...
  if (/^unsupported format: /.test(message)) return new PdfRenderError("unsupported_format", message.slice("unsupported format: ".length), { cause: error });
  // ... and parse/document.ts says this for content it does not read ...
  if (/^unsupported content: /.test(message)) return new PdfRenderError("unsupported_content", message.slice("unsupported content: ".length), { cause: error });
  // ... and stream/zip.ts and stream/body.ts say these for a file that is no .docx ...
  if (/not a \.docx file|^missing |encrypted entries/.test(message)) return new PdfRenderError("not_docx", message, { cause: error });
  // ... and these (or fflate's inflate errors, or a read past the end) for one that is but cannot be read.
  // ... and doc/ says "damaged: ..." for a .doc whose structures cannot be read.
  // ... and zip.ts, stream/zip.ts and stream/scan.ts say "too large: ..." for a document over the limits.
  if (/^too large: /.test(message)) return new PdfRenderError("too_large", message.slice("too large: ".length), { cause: error });
  if (/^damaged: |is corrupt|too many entries|bad (central directory|local header|ZIP64)|compression method|XML|unterminated|unexpected|invalid |outside the bounds/i.test(message)) return new PdfRenderError("damaged", message.replace(/^damaged: /, ""), { cause: error });
  return new PdfRenderError("failed", message, { cause: error });
}

/**
 * The name of the first script in `text` this renderer cannot shape, if
 * any: Mongolian, whose lines run top to bottom; and, when HarfBuzz is not
 * loaded (the harfbuzzjs package is absent; see shaper.ts), every script
 * that needs a shaper (scripts.ts). Hebrew letters never need one.
 */
export function unsupportedScript(text: string): string | undefined {
  for (const ch of text) {
    const cp = ch.codePointAt(0) as number;
    if (cp < 0x0600) continue;
    if (cp >= 0x1800 && cp <= 0x18af) return "Mongolian";
    if (shaperLoaded()) continue;
    const script = shapedScriptOf(cp);
    if (script && script.tag !== "Hebr") return script.name;
  }
  return undefined;
}
