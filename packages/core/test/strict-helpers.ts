import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";

/** Convert the namespace, direction and DrawingML percentage forms used by the PDF reductions to Strict. */
export function toStrictPackage(docx: Uint8Array): Uint8Array {
  const directions: Record<string, string> = { lrTb: "tb", tbRl: "rl", btLr: "lr", lrTbV: "tbV", tbRlV: "rlV", tbLrV: "lrV" };
  const files = unzipSync(docx);
  for (const name in files) {
    if (!/\.(xml|rels)$/i.test(name)) continue;
    const xml = strFromU8(files[name])
      .replace(/http:\/\/schemas\.openxmlformats\.org\/(wordprocessingml|officeDocument|drawingml)\/2006\//g, "http://purl.oclc.org/ooxml/$1/")
      .replace(/http:\/\/purl\.oclc\.org\/ooxml\/officeDocument\/(relationships\/)?(extended|custom)-properties/g, "http://purl.oclc.org/ooxml/officeDocument/$1$2Properties")
      .replace(/(<w:textDirection\b[^>]*\bw:val=")([^"]+)(")/g, (_whole, before: string, direction: string, after: string) => before + (directions[direction] ?? direction) + after)
      .replace(/(<wp14:(?:pctPosHOffset|pctPosVOffset|pctWidth|pctHeight)>)(-?\d+)(<\/wp14:[^>]+>)/g, (_whole, before: string, value: string, after: string) => `${before}${Number(value) / 1000}%${after}`)
      .replace(/<(\/?)wp[sg]:/g, "<$1wp:")
      .replace(/<(\/?)w:txbxContent/g, "<$1wne:txbxContent")
      .replace("<w:document ", '<w:document xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" ');
    files[name] = strToU8(xml);
  }
  return zipSync(files);
}

/** A Strict Open XML document part: unit-carrying measures, Strict names and values, a Strict border side. */
export const STRICT_DOCUMENT = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://purl.oclc.org/ooxml/wordprocessingml/main" xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships" xmlns:wp="http://purl.oclc.org/ooxml/drawingml/wordprocessingDrawing" xmlns:a="http://purl.oclc.org/ooxml/drawingml/main" xmlns:m="http://purl.oclc.org/ooxml/officeDocument/math" w:conformance="strict"><w:body><w:p><w:pPr><w:jc w:val="end"/><w:spacing w:before="6pt" w:after="0.5in" w:line="12.60pt" w:lineRule="auto"/><w:ind w:start="36pt" w:hanging="18pt"/><w:pBdr><w:start w:val="single" w:sz="0.5pt" w:space="4pt" w:color="auto"/></w:pBdr><w:tabs><w:tab w:val="end" w:pos="6in"/></w:tabs></w:pPr><w:r><w:rPr><w:sz w:val="12pt"/><w:position w:val="-3pt"/><w:spacing w:val="1pt"/><w:w w:val="150%"/><w:bdr w:val="single" w:sz="2.25pt" w:space="1pt" w:color="auto"/></w:rPr><w:t xml:space="preserve">Hello 72pt wide &amp; w:val="1in"</w:t></w:r><w:r><w:drawing><wp:inline distT="0" distB="0.5in" distL="0" distR="0"><wp:extent cx="1in" cy="2.54cm"/><a:graphic><a:graphicData uri="http://purl.oclc.org/ooxml/drawingml/picture"><a:solidFill><a:schemeClr val="accent1"><a:lumMod val="75%"/><a:alpha val="62.5%"/></a:schemeClr></a:solidFill></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p><w:tbl><w:tblPr><w:tblW w:w="90.0%" w:type="pct"/><w:jc w:val="start"/><w:tblInd w:w="5.40pt" w:type="dxa"/><w:tblBorders><w:end w:val="single" w:sz="1pt" w:space="0pt" w:color="000000"/></w:tblBorders><w:tblCellMar><w:start w:w="5.40pt" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid><w:gridCol w:w="2.54cm"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="1in" w:type="dxa"/><w:textDirection w:val="rl"/></w:tcPr><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sectPr><w:pgSz w:w="595.30pt" w:h="841.90pt"/><w:pgMar w:top="2.54cm" w:right="1in" w:bottom="72pt" w:left="6pc" w:header="35.40pt" w:footer="3pi" w:gutter="0mm"/><w:cols w:space="36pt"/></w:sectPr></w:body></w:document>`;

/** The same document as Word writes it in Transitional. */
export const TRANSITIONAL_DOCUMENT = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><w:body><w:p><w:pPr><w:jc w:val="right"/><w:spacing w:before="120" w:after="720" w:line="252" w:lineRule="auto"/><w:ind w:start="720" w:hanging="360"/><w:pBdr><w:left w:val="single" w:sz="4" w:space="4" w:color="auto"/></w:pBdr><w:tabs><w:tab w:val="right" w:pos="8640"/></w:tabs></w:pPr><w:r><w:rPr><w:sz w:val="24"/><w:position w:val="-6"/><w:spacing w:val="20"/><w:w w:val="150"/><w:bdr w:val="single" w:sz="18" w:space="1" w:color="auto"/></w:rPr><w:t xml:space="preserve">Hello 72pt wide &amp; w:val="1in"</w:t></w:r><w:r><w:drawing><wp:inline distT="0" distB="457200" distL="0" distR="0"><wp:extent cx="914400" cy="914400"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><a:solidFill><a:schemeClr val="accent1"><a:lumMod val="75000"/><a:alpha val="62500"/></a:schemeClr></a:solidFill></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p><w:tbl><w:tblPr><w:tblW w:w="4500" w:type="pct"/><w:jc w:val="left"/><w:tblInd w:w="108" w:type="dxa"/><w:tblBorders><w:right w:val="single" w:sz="8" w:space="0" w:color="000000"/></w:tblBorders><w:tblCellMar><w:left w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid><w:gridCol w:w="1440"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="1440" w:type="dxa"/><w:textDirection w:val="tbRl"/></w:tcPr><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="720" w:gutter="0"/><w:cols w:space="720"/></w:sectPr></w:body></w:document>`;

export const STRICT_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="http://purl.oclc.org/ooxml/officeDocument/relationships/extendedProperties" Target="docProps/app.xml"/></Relationships>`;

export const TRANSITIONAL_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>`;

const CONTENT_TYPES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`;

/** A package of the document and the rels, Strict or Transitional, with `text` in place of the first paragraph's text. */
export function strictPackage(strict: boolean, text = "Hello 72pt wide &amp; w:val=\"1in\""): Uint8Array {
  const document = (strict ? STRICT_DOCUMENT : TRANSITIONAL_DOCUMENT).replace("Hello 72pt wide &amp; w:val=\"1in\"", text);
  return zipSync({
    "[Content_Types].xml": strToU8(CONTENT_TYPES),
    "_rels/.rels": strToU8(strict ? STRICT_RELS : TRANSITIONAL_RELS),
    "word/document.xml": strToU8(document),
  });
}
