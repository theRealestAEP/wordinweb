// The streams of a Word 97-2003 .doc: an OLE compound file ([MS-CFB]) read
// with the cfb package, holding the WordDocument stream (the FIB, the text,
// the formatting pages), a Table stream (0Table or 1Table; the FIB says
// which) and an optional Data stream (pictures, large property sets).
import * as CFB from "cfb";

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
 * Whether the front of an OLE compound file names a WordDocument stream in
 * its directory (a .doc), rather than an EncryptionInfo stream (a
 * password-protected .docx). The directory sits within the first sectors of
 * any file Word writes; 64 KB is plenty.
 */
export function isWordDocumentFile(front: Uint8Array): boolean {
  if (!isOleFile(front)) return false;
  const names = utf16Names(front);
  return names.includes("WordDocument") && !names.includes("EncryptionInfo") && !names.includes("EncryptedPackage");
}

function utf16Names(bytes: Uint8Array): string {
  let out = "";
  for (let i = 0; i + 1 < bytes.length; i += 2) {
    const code = bytes[i] | (bytes[i + 1] << 8);
    out += code >= 0x20 && code < 0x7f ? String.fromCharCode(code) : " ";
  }
  return out;
}

/** The three streams, or an Error naming what the file is instead. */
export function openDocStreams(bytes: Uint8Array): DocStreams {
  if (!isOleFile(bytes)) throw new Error("not a Word document: not an OLE compound file");
  const cfb = CFB.read(bytes, { type: "array" });
  const wordDocument = streamBytes(cfb, "WordDocument");
  if (!wordDocument) throw new Error("not a Word document: the compound file has no WordDocument stream");
  if (wordDocument.length < 0x0c) throw new Error("damaged: the WordDocument stream is shorter than the FIB");
  const flags = wordDocument[0x0a] | (wordDocument[0x0b] << 8);
  const tableName = flags & 0x0200 ? "1Table" : "0Table";
  const table = streamBytes(cfb, tableName);
  if (!table) throw new Error(`damaged: the compound file has no ${tableName} stream`);
  return { wordDocument, table, data: streamBytes(cfb, "Data") };
}

function streamBytes(cfb: CFB.CFB$Container, name: string): Uint8Array | null {
  const entry = CFB.find(cfb, "/" + name) ?? CFB.find(cfb, name);
  if (!entry || entry.type !== 2) return null;
  const content = entry.content;
  return content instanceof Uint8Array ? content : Uint8Array.from(content);
}
