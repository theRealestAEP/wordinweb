import { zipSync, strToU8 } from "fflate";

/** Build a minimal in-memory .docx from part name → XML string. */
export function makeDocx(parts: Record<string, string>): Uint8Array {
  const files: Record<string, Uint8Array> = {};
  const defaults: Record<string, string> = {
    "[Content_Types].xml": `<?xml version="1.0"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`,
    "_rels/.rels": `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`,
  };
  for (const [name, content] of Object.entries({ ...defaults, ...parts })) {
    files[name] = strToU8(content);
  }
  return zipSync(files);
}

/** makeDocx plus binary media parts. */
export function makeDocxWithMedia(parts: Record<string, string>, media: Record<string, Uint8Array>): Uint8Array {
  const files: Record<string, Uint8Array> = {};
  const base = makeDocx(parts);
  // unzip-free approach: rebuild with media included
  void base;
  const defaults: Record<string, string> = {};
  void defaults;
  const textFiles: Record<string, Uint8Array> = {};
  for (const [name, content] of Object.entries(parts)) textFiles[name] = strToU8(content);
  if (!parts["[Content_Types].xml"]) {
    textFiles["[Content_Types].xml"] = strToU8(`<?xml version="1.0"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`);
  }
  if (!parts["_rels/.rels"]) {
    textFiles["_rels/.rels"] = strToU8(`<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`);
  }
  for (const [k, v] of Object.entries(textFiles)) files[k] = v;
  for (const [k, v] of Object.entries(media)) files[k] = v;
  return zipSync(files);
}

export const W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

export function wrapDocument(bodyXml: string): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${W_NS}><w:body>${bodyXml}</w:body></w:document>`;
}

/** Simple paragraph XML. */
export function p(text: string, extra = ""): string {
  return `<w:p>${extra}<w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
}
