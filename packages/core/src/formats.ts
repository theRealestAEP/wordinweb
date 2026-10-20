// The formats a .docx reader is handed instead of a .docx, told from their
// first bytes, so the refusal names what the file is. A Word 97-2003 .doc
// (an OLE compound file) is not among them: it is read (doc/). Nor is a
// Word for Macintosh 4.0/5.x document: oldword/ reads it, and the callers
// ask oldword/ first, so the entry below names only a variant it does not
// read.

/** Enough of a file's start for every check here. */
export const FORMAT_HEAD_BYTES = 1024;

const SIGNATURES: [number[], string][] = [
  [[0x7b, 0x5c, 0x72, 0x74, 0x66], "an RTF file"], // {\rtf
  [[0x25, 0x50, 0x44, 0x46], "a PDF file"], // %PDF
  [[0xff, 0x57, 0x50, 0x43], "a WordPerfect document"], // \xFFWPC
  [[0xfe, 0x37, 0x00], "a Word for Macintosh 4.0/5.0 document"],
  [[0x31, 0xbe, 0x00, 0x00], "a Word for DOS or Windows Write document"],
  [[0x32, 0xbe, 0x00, 0x00], "a Word for DOS or Windows Write document"],
  [[0x9b, 0xa5], "a Word for Windows 1.x/2.0 document"],
  [[0xdb, 0xa5], "a Word for Windows 1.x/2.0 document"],
];

/** What the file whose first bytes are `head` is, when it is one of the formats above, a web page or Word 2003 XML. */
export function otherFormat(head: Uint8Array): string | undefined {
  for (const [signature, name] of SIGNATURES) if (signature.every((byte, i) => head[i] === byte)) return name;
  // Text formats: after a byte order mark and white space.
  const text = new TextDecoder("latin1").decode(head).replace(/^(\xef\xbb\xbf|\xff\xfe|\xfe\xff)/, "").replace(/\0/g, "").trimStart().toLowerCase();
  if (text.startsWith("mime-version:")) return "a web archive (MHTML)";
  if (text.startsWith("<?xml") && /<w:worddocument|progid="word\.document"/.test(text)) return "a Word 2003 XML document";
  if (/^<(!doctype html|html)[\s>]/.test(text)) return "an HTML file";
  return undefined;
}
