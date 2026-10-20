/** Why a compare failed, for a caller to act on without reading the message. */
export type CompareErrorCode =
  /** An input does not exist, or is not a file. */
  | "missing"
  /** An input is not a .docx: not a zip, or a zip with no Word document. */
  | "not_docx"
  /** An input is another format the compare does not read, which the message
   * names: an RTF, PDF or OpenDocument file, a password-protected
   * document. (A Word 6.0-2003 .doc is read: it is converted to a .docx
   * working file first.) */
  | "unsupported_format"
  /** An input is a .docx or .doc that cannot be read: a corrupt zip entry, XML that is not well formed, a .doc whose tables are corrupt. */
  | "damaged"
  /** An input holds content Word prints that the compare does not read,
   * which the message names: a w:altChunk (a part in another format, such
   * as HTML, that Word imports on open) or a w:subDoc (a subdocument Word
   * reads from another file). */
  | "unsupported_content"
  /** An input holds more outside its body than the compare keeps in memory
   * (MAX_HELD_BYTES), or a block of its body or notes over the XML limit. */
  | "too_large"
  /** The temp folder or the result's folder ran out of space. */
  | "disk_full"
  /** The folder for the result (or the printable copy) does not exist. */
  | "result_folder_missing"
  /** The result would overwrite an input. */
  | "result_is_input"
  /** Two of the result, the printable copy and the PDF would be written to one path; the message says which. */
  | "output_collision"
  /** The PDF (`pdf` option) holds text in a script the PDF renderer cannot shape; the message names it. */
  | "unsupported_script"
  /** Anything else. The message and `cause` say what. */
  | "failed";

export class CompareError extends Error {
  /** The error this one wraps, when it wraps one. */
  readonly cause?: unknown;

  constructor(
    readonly code: CompareErrorCode,
    message: string,
    /** The input at fault, when one is. */
    readonly file?: "original" | "revised",
    options?: { cause?: unknown },
  ) {
    super(message);
    this.name = "CompareError";
    this.cause = options?.cause;
  }
}

/** Out of space, as the file system reports it. */
export function isDiskFull(error: unknown): boolean {
  const code = (error as { code?: unknown } | null)?.code;
  return code === "ENOSPC" || code === "EDQUOT";
}
