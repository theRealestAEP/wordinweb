import { W_NS, makeDocx, p, wrapDocument } from "./helpers.js";

/** XML failures shared by the PDF and file-compare tests. */
export const xmlErrorCases: [string, Uint8Array][] = [
  ["DOCTYPE in document.xml", makeDocx({
    "word/document.xml": '<!DOCTYPE w:document>' + wrapDocument(p("Text")),
  })],
  ["DOCTYPE with entities in document.xml", makeDocx({
    "word/document.xml": '<!DOCTYPE w:document [<!ENTITY a "expanded">]>' + wrapDocument(p("&a;")),
  })],
  ["entity declaration in styles.xml", makeDocx({
    "word/document.xml": wrapDocument(p("Text")),
    "word/styles.xml": `<!ENTITY a "expanded"><w:styles ${W_NS}/>`,
  })],
  ["malformed XML in a header", makeDocx({
    "word/document.xml": wrapDocument(p("Text") + '<w:sectPr><w:headerReference xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" w:type="default" r:id="rIdH"/></w:sectPr>'),
    "word/_rels/document.xml.rels": '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdH" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/></Relationships>',
    "word/header1.xml": `<w:hdr ${W_NS}><w:p attr="unterminated></w:p></w:hdr>`,
  })],
  ["nesting deeper than 512 levels", makeDocx({
    "word/document.xml": wrapDocument('<w:p>' + '<w:r>'.repeat(513) + '<w:t>Text</w:t>' + '</w:r>'.repeat(513) + '</w:p>'),
  })],
  ["nesting deeper than 512 levels in styles.xml", makeDocx({
    "word/document.xml": wrapDocument(p("Text")),
    "word/styles.xml": `<w:styles ${W_NS}>` + '<w:style>'.repeat(513) + '</w:style>'.repeat(513) + '</w:styles>',
  })],
  ["DOCTYPE in an unused part", makeDocx({
    "word/document.xml": wrapDocument(p("Text")),
    "customXml/item1.xml": '<!DOCTYPE x SYSTEM "file:///unused"><x/>',
  })],
  ["invalid numeric character reference", makeDocx({
    "word/document.xml": wrapDocument(p("&#x110000;")),
  })],
];
