import { crc32, deflateRawSync } from "node:zlib";
import { unzipSync } from "fflate";
import { makeDocx, p, wrapDocument } from "./helpers.js";

/**
 * Crafted .docx files for the size limits: each small on disk, each a part
 * that would take hundreds of megabytes if it were inflated. Kept in one
 * place for the core and compare-cli tests.
 */

const MB = 1024 * 1024;

/** One zip entry as it is stored: deflated bytes, with the CRC and size the directory gives. */
interface Entry {
  name: string;
  deflated: Uint8Array;
  crc: number;
  size: number;
}

/** About `mb` megabytes of XML, deflated: a real deflate bomb, about 1000 to 1. By default an element of spaces; else `unit` repeated, between `head` and `tail`. */
function bomb(name: string, mb: number, head = Buffer.from("<a>"), unit = " ", tail = Buffer.from("</a>")): Entry {
  // Each megabyte deflates to the same bytes, so the stream is built from
  // one deflated megabyte between a head and a tail: no 400 MB buffer.
  const block = Buffer.from(unit.repeat(Math.floor(MB / unit.length)));
  const pieces = [head, ...Array.from({ length: mb }, () => block), tail];
  const deflated = deflatePieces(pieces);
  let crc = 0;
  for (const piece of pieces) crc = crc32(piece, crc);
  return { name, deflated, crc, size: head.length + mb * block.length + tail.length };
}

/**
 * Deflate the concatenated pieces without holding them together: each
 * piece is a stored-free run of non-final deflate blocks (a sync flush
 * between them), and the last block is final. A repeated piece deflates
 * to the same bytes, which are made once.
 */
function deflatePieces(pieces: Buffer[]): Uint8Array {
  const made = new Map<Buffer, Uint8Array>();
  const out: Uint8Array[] = [];
  pieces.forEach((piece, i) => {
    const last = i === pieces.length - 1;
    if (last) {
      out.push(deflateRawSync(piece, { level: 9 }));
      return;
    }
    let bytes = made.get(piece);
    if (!bytes) {
      // A full flush ends on a byte boundary and keeps no history: the next
      // piece's blocks follow it as they are.
      bytes = deflateRawSync(piece, { level: 9, finishFlush: 3 /* Z_FULL_FLUSH */ });
      made.set(piece, bytes);
    }
    out.push(bytes);
  });
  return Buffer.concat(out);
}

/** A zip of `entries`, each deflated, with the sizes and CRCs they carry. */
function zip(entries: Entry[]): Uint8Array {
  const local: Buffer[] = [];
  const central: Buffer[] = [];
  let offset = 0;
  for (const entry of entries) {
    const name = Buffer.from(entry.name);
    const header = Buffer.alloc(30);
    header.writeUInt32LE(0x04034b50, 0);
    header.writeUInt16LE(20, 4);
    header.writeUInt16LE(8, 8);
    header.writeUInt32LE(entry.crc >>> 0, 14);
    header.writeUInt32LE(entry.deflated.length, 18);
    header.writeUInt32LE(entry.size, 22);
    header.writeUInt16LE(name.length, 26);
    const record = Buffer.alloc(46);
    record.writeUInt32LE(0x02014b50, 0);
    record.writeUInt16LE(20, 4);
    record.writeUInt16LE(20, 6);
    record.writeUInt16LE(8, 10);
    record.writeUInt32LE(entry.crc >>> 0, 16);
    record.writeUInt32LE(entry.deflated.length, 20);
    record.writeUInt32LE(entry.size, 24);
    record.writeUInt16LE(name.length, 28);
    record.writeUInt32LE(offset, 42);
    local.push(header, name, Buffer.from(entry.deflated));
    central.push(record, name);
    offset += 30 + name.length + entry.deflated.length;
  }
  const cd = Buffer.concat(central);
  const end = Buffer.alloc(22);
  end.writeUInt32LE(0x06054b50, 0);
  end.writeUInt16LE(entries.length, 8);
  end.writeUInt16LE(entries.length, 10);
  end.writeUInt32LE(cd.length, 12);
  end.writeUInt32LE(offset, 16);
  return Buffer.concat([...local, cd, end]);
}

/** The parts of a small, good .docx, as entries. */
function goodEntries(): Entry[] {
  const parts = unzipSync(makeDocx({ "word/document.xml": wrapDocument(p("Text.")) }));
  return Object.entries(parts).map(([name, data]) => ({ name, deflated: deflateRawSync(data), crc: crc32(data), size: data.length }));
}

/** A .docx whose styles.xml is a deflate bomb of `mb` megabytes, its directory honest. */
export function stylesBomb(mb = 400): Uint8Array {
  return zip([...goodEntries(), bomb("word/styles.xml", mb)]);
}

/** A .docx whose styles.xml inflates to `mb` megabytes while its directory says 1000 bytes. */
export function lyingStyles(mb = 400): Uint8Array {
  return zip([...goodEntries(), { ...bomb("word/styles.xml", mb), size: 1000 }]);
}

/** A .docx whose `count` headers are each an `mb` megabyte bomb: each under the XML limit, together over the package's. */
export function manyParts(count = 9, mb = 60): Uint8Array {
  const part = bomb("", mb);
  return zip([...goodEntries(), ...Array.from({ length: count }, (_, i) => ({ ...part, name: `word/header${i + 1}.xml` }))]);
}

/** A .docx whose document.xml is a deflate bomb of one paragraph of `mb` megabytes of spaces, its directory honest. */
export function giantParagraph(mb = 400): Uint8Array {
  const head = Buffer.from(wrapDocument("").replace("</w:body></w:document>", "<w:p><w:r><w:t>"));
  const body = bomb("word/document.xml", mb, head, " ", Buffer.from("</w:t></w:r></w:p></w:body></w:document>"));
  return zip([...goodEntries().filter((e) => e.name !== "word/document.xml"), body]);
}

/** A .docx whose document.xml is a deflate bomb of `mb` megabytes of empty paragraphs, its directory honest. */
export function paragraphsBody(mb = 400): Uint8Array {
  const head = Buffer.from(wrapDocument("").replace("</w:body></w:document>", ""));
  const body = bomb("word/document.xml", mb, head, "<w:p/>", Buffer.from("</w:body></w:document>"));
  return zip([...goodEntries().filter((e) => e.name !== "word/document.xml"), body]);
}

/** A .docx whose document.xml inflates to `mb` megabytes while its directory says its real size less 1000. */
export function lyingBody(mb = 400): Uint8Array {
  const body = bomb("word/document.xml", mb);
  return zip([...goodEntries().filter((e) => e.name !== "word/document.xml"), { ...body, size: body.size - 1000 }]);
}
