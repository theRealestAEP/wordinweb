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
  /** The stream `name` of the embedded object storage `ObjectPool/<storage>`, where OLE objects keep their pictures. */
  objectStream?: (storage: string, name: string) => Uint8Array | null;
}

/** The first bytes of every OLE compound file. */
const OLE_SIGNATURE = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];

export function isOleFile(head: Uint8Array): boolean {
  return head.length >= 8 && OLE_SIGNATURE.every((byte, i) => head[i] === byte);
}

/** Whether `bytes` are a Word .doc: an OLE compound file with a WordDocument stream. */
export function isWordDocumentFile(bytes: Uint8Array): boolean {
  if (!isOleFile(bytes)) return false;
  try {
    return streamBytes(CFB.read(bytes, { type: "array" }), "WordDocument") !== null;
  } catch {
    return false;
  }
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
    throw new Error(`unsupported format: ${compoundFileKind(rootStreamNames(cfb))}`);
  }
  if (wordDocument.length < 0x0c) throw new Error("damaged: the WordDocument stream is shorter than the FIB");
  // A Word 6.0/95 file keeps everything, pictures too, in the WordDocument stream and has no Table stream.
  const wIdent = u16(wordDocument, 0);
  const nFib = u16(wordDocument, 2);
  if (wIdent === 0xa5dc) return { wordDocument, table: wordDocument, data: wordDocument, objectStream: (storage, name) => streamBytes(cfb, `ObjectPool/${storage}/${name}`) };
  if (wIdent === 0xa5ec && nFib < 0x00c1) throw new Error("unsupported format: a Word 6.0/95 document (only Word 97-2003 .doc files are read)");
  if (wIdent !== 0xa5ec) throw new Error("unsupported format: not a Word document (the WordDocument stream does not start with a FIB)");
  const flags = u16(wordDocument, 0x0a);
  const tableName = flags & 0x0200 ? "1Table" : "0Table";
  const table = streamBytes(cfb, tableName);
  if (!table) throw new Error(`damaged: the compound file has no ${tableName} stream`);
  return { wordDocument, table, data: streamBytes(cfb, "Data"), objectStream: (storage, name) => streamBytes(cfb, `ObjectPool/${storage}/${name}`) };
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

/** What a compound file with no WordDocument stream holds, by its streams' names. */
function compoundFileKind(names: string[]): string {
  if (names.includes("EncryptionInfo") || names.includes("EncryptedPackage")) return "a password-protected document (its package is encrypted)";
  if (names.includes("PowerPoint Document")) return "a PowerPoint 97-2003 presentation";
  if (names.includes("Workbook") || names.includes("Book")) return "an Excel 97-2003 workbook";
  return "an OLE compound file with no Word document";
}

/** The names of the streams in the root storage. */
function rootStreamNames(cfb: CFB.CFB$Container): string[] {
  const root = cfb.FullPaths[0];
  return cfb.FullPaths.filter((path, i) => path.startsWith(root) && !path.slice(root.length).includes("/") && cfb.FileIndex[i].type === 2).map((path) => path.slice(root.length));
}
