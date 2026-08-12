import { CompareNote, DocxDocument, compareDocuments } from "@wordinweb/core";

/**
 * The host affordance for Compare Documents — Word's "legal blackline".
 *
 * The engine call (`compareDocuments` in @wordinweb/core) is pure: two
 * documents in, a third out. Everything a browser has to do around it lives
 * here — pick a file, read it, hand the result somewhere — so the toolbar
 * entry is one button and the engine stays free of the DOM.
 *
 * The document on screen is the ORIGINAL and the picked file is the REVISED
 * one, which is the direction Word's dialog defaults to. The result is a new
 * document whose differences are tracked changes attributed to `author`;
 * accepting them all gives the picked file back, rejecting them all gives what
 * is on screen back.
 */

export interface CompareResult {
  /** The compared document, as .docx bytes. */
  bytes: Uint8Array;
  /** What the comparison could not express as a tracked change. */
  notes: CompareNote[];
  /** The picked file's name, for whatever the host calls the result. */
  revisedName: string;
}

export interface CompareWithFileOptions {
  /** Who the tracked changes are attributed to. */
  author?: string;
  /** Timestamp for every revision. Defaults to now. */
  date?: string;
  /** Compare formatting as well as text. Default true, as in Word. */
  formatting?: boolean;
  /** Report whitespace-only differences. Default true, as in Word. */
  whitespace?: boolean;
}

/**
 * Compare `originalBytes` with a .docx `file`. Returns the result and the list
 * of differences that could not be tracked, which a caller should show rather
 * than swallow — a comparison that quietly skipped something is worse than one
 * that says what it skipped.
 */
export async function compareWithFile(
  originalBytes: Uint8Array,
  file: File,
  options: CompareWithFileOptions = {},
): Promise<CompareResult> {
  const revisedBytes = new Uint8Array(await file.arrayBuffer());
  const notes: CompareNote[] = [];
  const merged = compareDocuments(DocxDocument.load(originalBytes), DocxDocument.load(revisedBytes), {
    author: options.author ?? "Comparison",
    date: options.date ?? new Date().toISOString().replace(/\.\d{3}Z$/, "Z"),
    formatting: options.formatting,
    whitespace: options.whitespace,
    onNote: (note) => notes.push(note),
  });
  return { bytes: merged.save(), notes, revisedName: file.name };
}

/** Ask for one .docx. Resolves to null when the picker is dismissed. */
export function pickDocx(): Promise<File | null> {
  return new Promise((resolve) => {
    const input = document.createElement("input");
    input.type = "file";
    input.accept = ".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document";
    input.style.display = "none";
    // Chrome fires no event when the dialog is cancelled, so the input is torn
    // down on the first change OR when the window regains focus, whichever
    // comes first — otherwise a cancelled pick leaks a node per attempt.
    const done = (value: File | null): void => {
      window.removeEventListener("focus", onFocus);
      input.remove();
      resolve(value);
    };
    const onFocus = (): void => {
      setTimeout(() => done(input.files?.[0] ?? null), 300);
    };
    input.addEventListener("change", () => done(input.files?.[0] ?? null), { once: true });
    window.addEventListener("focus", onFocus, { once: true });
    document.body.append(input);
    input.click();
  });
}

/** Save bytes to the user's downloads under `name`. */
export function downloadDocx(bytes: Uint8Array, name: string): void {
  const blob = new Blob([bytes as unknown as BlobPart], {
    type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = name;
  link.click();
  // Revoking immediately can beat the download in Safari; one turn is enough.
  setTimeout(() => URL.revokeObjectURL(url), 0);
}

/**
 * The whole affordance behind one click: pick a file, compare, hand the result
 * to `onResult` (or download it), and report anything untrackable.
 *
 * Returns false when no file was picked, so a caller can tell "cancelled" from
 * "compared with nothing to say".
 */
export async function pickAndCompare(
  originalBytes: Uint8Array,
  options: CompareWithFileOptions & {
    onResult?: (result: CompareResult) => void;
    onNotes?: (notes: CompareNote[]) => void;
  } = {},
): Promise<boolean> {
  const file = await pickDocx();
  if (!file) return false;
  const result = await compareWithFile(originalBytes, file, options);
  if (options.onResult) options.onResult(result);
  else downloadDocx(result.bytes, comparedName(file.name));
  if (result.notes.length > 0) options.onNotes?.(result.notes);
  return true;
}

/** "report.docx" → "report (compared).docx". */
export function comparedName(revisedName: string): string {
  const dot = revisedName.lastIndexOf(".");
  return dot <= 0
    ? `${revisedName} (compared).docx`
    : `${revisedName.slice(0, dot)} (compared)${revisedName.slice(dot)}`;
}
