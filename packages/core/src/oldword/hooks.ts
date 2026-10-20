// The one place a document in one of the oldest Word formats (before Word
// 6.0, not an OLE compound file) becomes a .docx for the renderer and the
// compare, as doc/hooks.ts is for a Word 97-2003 .doc.
import { zipSync } from "fflate";
import { FIXED_ZIP_MTIME } from "../zip.js";
import { isDosWordFile, readDosWord } from "./dos.js";
import { isMacEarlyFile, readMacEarly } from "./mac-early.js";
import { isMacWordFile, readMacWord } from "./mac.js";
import { OldDocument } from "./model.js";
import { oldDocumentFiles } from "./package.js";
import { isWinWordFile, readWinWord } from "./winword.js";

/** Whether the file whose first bytes are `head` is one of the old Word formats read here. */
export function isOldWordFile(head: Uint8Array): boolean {
  return isMacEarlyFile(head) || isMacWordFile(head) || isWinWordFile(head) || isDosWordFile(head);
}

function readOldWord(bytes: Uint8Array): OldDocument {
  if (isMacEarlyFile(bytes)) return readMacEarly(bytes);
  if (isDosWordFile(bytes)) return readDosWord(bytes);
  return isMacWordFile(bytes) ? readMacWord(bytes) : readWinWord(bytes);
}

/** An old Word document as a .docx package's parts; "damaged: ..." for one whose structures cannot be read. */
export function oldWordPackageFiles(bytes: Uint8Array): Record<string, Uint8Array> {
  try {
    return oldDocumentFiles(readOldWord(bytes));
  } catch (error) {
    throw new Error(`damaged: ${error instanceof Error ? error.message : String(error)}`);
  }
}

/** An old Word document as a .docx. */
export function oldWordToDocx(bytes: Uint8Array): Uint8Array {
  return zipSync(oldWordPackageFiles(bytes), { mtime: FIXED_ZIP_MTIME });
}
