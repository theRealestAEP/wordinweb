// The streams of a Word 97-2003 .doc: an OLE compound file ([MS-CFB]) read
// with the cfb package, holding the WordDocument stream (the FIB, the text,
// the formatting pages), a Table stream (0Table or 1Table; the FIB says
// which) and an optional Data stream (pictures, large property sets).
import * as CFB from "cfb";
import { u16 } from "./bytes.js";

export interface DocStreams {
  wordDocument: Uint8Array;
  table: Uint8Array;
  /** Absent in a document with no pictures or large properties. */
  data: Uint8Array | null;
}

/** The first bytes of every OLE compound file. */
const OLE_SIGNATURE = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];

export function isOleFile(head: Uint8Array): boolean {
  return head.length >= 8 && OLE_SIGNATURE.every((byte, i) => head[i] === byte);
}

/**
 * What an OLE compound file holds, from its directory: a Word document (a
 * .doc), an encrypted package (a password-protected .docx), or something
 * else. "none" for a file that is no compound file at all, or one whose
 * directory cannot be read.
 */
export function oleKind(bytes: Uint8Array): "doc" | "encrypted" | "other" | "none" {
  if (!isOleFile(bytes)) return "none";
  let names: string[];
  try {
    names = rootStreamNames(CFB.read(bytes, { type: "array" }));
  } catch {
    return "none";
  }
  if (names.includes("EncryptionInfo") || names.includes("EncryptedPackage")) return "encrypted";
  return names.includes("WordDocument") ? "doc" : "other";
}

/** Whether `bytes` are a Word 97-2003 .doc: an OLE compound file with a WordDocument stream. */
export function isWordDocumentFile(bytes: Uint8Array): boolean {
  return oleKind(bytes) === "doc";
}

/**
 * The three streams. Throws "unsupported format: ..." for a file that is
 * another kind of document, and "damaged: ..." for a .doc that cannot be
 * read.
 */
export function openDocStreams(bytes: Uint8Array): DocStreams {
  if (!isOleFile(bytes)) throw new Error("unsupported format: not an OLE compound file (a Word 97-2003 .doc is one)");
  let cfb: CFB.CFB$Container;
  try {
    cfb = CFB.read(bytes, { type: "array" });
  } catch (error) {
    throw new Error(`damaged: the compound file cannot be read (${error instanceof Error ? error.message : String(error)})`);
  }
  const wordDocument = streamBytes(cfb, "WordDocument");
  if (!wordDocument) {
    const names = rootStreamNames(cfb);
    const what = names.includes("EncryptionInfo") || names.includes("EncryptedPackage") ? "a password-protected document (its package is encrypted)" : "an OLE compound file with no Word document";
    throw new Error(`unsupported format: ${what}`);
  }
  if (wordDocument.length < 0x0c) throw new Error("damaged: the WordDocument stream is shorter than the FIB");
  // A Word 6.0/95 file keeps everything in the WordDocument stream and has no Table stream.
  const wIdent = u16(wordDocument, 0);
  const nFib = u16(wordDocument, 2);
  if (wIdent === 0xa5dc || (wIdent === 0xa5ec && nFib < 0x00c1)) throw new Error("unsupported format: a Word 6.0/95 document (only Word 97-2003 .doc files are read)");
  if (wIdent !== 0xa5ec) throw new Error("unsupported format: not a Word document (the WordDocument stream does not start with a FIB)");
  const flags = u16(wordDocument, 0x0a);
  const tableName = flags & 0x0200 ? "1Table" : "0Table";
  const table = streamBytes(cfb, tableName);
  if (!table) throw new Error(`damaged: the compound file has no ${tableName} stream`);
  return { wordDocument, table, data: streamBytes(cfb, "Data") };
}

/**
 * A stream of the root storage by name. An embedded Word object keeps
 * streams of the same names deeper down (ObjectPool/_123/1Table), so the
 * path is exact.
 */
function streamBytes(cfb: CFB.CFB$Container, name: string): Uint8Array | null {
  const entry = CFB.find(cfb, "/" + name);
  if (!entry || entry.type !== 2) return null;
  const content = entry.content;
  return content instanceof Uint8Array ? content : Uint8Array.from(content);
}

/** The names of the streams in the root storage. */
function rootStreamNames(cfb: CFB.CFB$Container): string[] {
  const root = cfb.FullPaths[0];
  return cfb.FullPaths.filter((path, i) => path.startsWith(root) && !path.slice(root.length).includes("/") && cfb.FileIndex[i].type === 2).map((path) => path.slice(root.length));
}
