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
    return streamBytes(readCompoundFile(bytes), "WordDocument") !== null;
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
    cfb = readCompoundFile(bytes);
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

function readCompoundFile(bytes: Uint8Array): CFB.CFB$Container {
  bytes = checkCompoundFile(bytes);
  // CFB keeps streams as Buffers with this input. Array input creates JS
  // number arrays, which can use many times the file's size in heap memory.
  return typeof Buffer !== "undefined"
    ? CFB.read(Buffer.from(bytes.buffer, bytes.byteOffset, bytes.byteLength), { type: "buffer" })
    : CFB.read(bytes, { type: "array" });
}

/** Check the first directory name and repair absent links ([MS-CFB] 2.2, 2.6). */
function checkCompoundFile(bytes: Uint8Array): Uint8Array {
  if (bytes.length < 512) return bytes;
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const int = (offset: number): number => view.getInt32(offset, true);
  const sectorSize = view.getUint16(26, true) === 4 ? 4096 : 512;
  const at = (sector: number): number => (sector + 1) * sectorSize;
  const directory = int(48);
  // 062193.doc points into garbage with a 6,801-byte root name. CFB treats
  // it as directory entries and exhausts memory. A directory name fits in
  // 64 bytes ([MS-CFB] 2.6.1); check only the first entry before allocation.
  if (directory >= 0 && at(directory) + 128 <= bytes.length && view.getUint16(at(directory) + 64, true) > 64) {
    throw new Error("the compound file's first directory name exceeds 64 bytes");
  }
  const fat: number[] = [];
  for (let offset = 76; offset < 512 && int(offset) >= 0; offset += 4) fat.push(int(offset));
  const difatSeen = new Set<number>();
  for (let sector = int(68), remaining = int(72); sector >= 0 && at(sector) + sectorSize <= bytes.length && !difatSeen.has(sector); sector = int(at(sector) + sectorSize - 4)) {
    difatSeen.add(sector);
    for (let offset = at(sector); offset < at(sector) + sectorSize - 4 && int(offset) !== -2; offset += 4) fat.push(int(offset));
    if (remaining-- <= 0) break;
  }
  let result = bytes;
  let repaired: DataView | undefined;
  const seen = new Set<number>();
  for (let sector = directory; sector >= 0 && at(sector) + 128 <= bytes.length && !seen.has(sector);) {
    seen.add(sector);
    for (let offset = at(sector); offset + 128 <= Math.min(at(sector) + sectorSize, bytes.length); offset += 128) {
      // Some writers use ENDOFCHAIN (-2) for absent directory links.
      // CFB needs NOSTREAM (-1). Copy only when this repair is needed.
      for (const field of [68, 72, 76]) {
        if (int(offset + field) === -2) {
          if (!repaired) {
            result = new Uint8Array(bytes);
            repaired = new DataView(result.buffer);
          }
          repaired.setInt32(offset + field, -1, true);
        }
      }
    }
    const fatSector = fat[Math.floor(sector / (sectorSize / 4))];
    if (fatSector === undefined || fatSector < 0) break;
    const next = at(fatSector) + (sector % (sectorSize / 4)) * 4;
    if (next + 4 > bytes.length) break;
    sector = int(next);
  }
  return result;
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
  if (["WordDocument", "0Table", "1Table", "Data"].includes(name) && content.length < entry.size) {
    throw new Error(`damaged: the ${name} stream is shorter than its declared size`);
  }
  return content instanceof Uint8Array
    ? new Uint8Array(content.buffer, content.byteOffset, content.length)
    : Uint8Array.from(content);
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
