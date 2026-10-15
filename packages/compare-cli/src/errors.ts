/** Why a compare failed, for a caller to act on without reading the message. */
export type CompareErrorCode =
  /** An input does not exist, or is not a file. */
  | "missing"
  /** An input is not a .docx: not a zip, a .doc or password-protected file, or a zip with no Word document. */
  | "not_docx"
  /** An input is a .docx that cannot be read: a corrupt zip entry, or XML that is not well formed. */
  | "damaged"
  /** An input holds more outside its body than the compare keeps in memory (MAX_HELD_BYTES). */
  | "too_large"
  /** The temp folder or the result's folder ran out of space. */
  | "disk_full"
  /** The folder for the result (or the printable copy) does not exist. */
  | "result_folder_missing"
  /** The result would overwrite an input. */
  | "result_is_input"
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
