import { describe, expect, it } from "vitest";
import { DocxDocument } from "../src/docx.js";
import { Block } from "../src/model.js";
import { openStreamingDocument } from "../src/stream/body.js";
import { bytesSource } from "../src/stream/source.js";
import { W_NS, makeDocx, p, wrapDocument } from "./helpers.js";

// Elements that only wrap content Word prints: the content is read as if
// the wrapper were not there. Each case's text must come out of both the
// whole load (DocxDocument.load) and the streaming one (renderDocxToPdf's).

const MC = 'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"';
const run = (text: string): string => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const tbl = (rows: string): string => `<w:tbl><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>${rows}</w:tbl>`;
const tc = (content: string): string => `<w:tc>${content}</w:tc>`;

/** Every run's text in the blocks, a paragraph a line. */
function text(blocks: Block[]): string[] {
  const lines: string[] = [];
  for (const block of blocks) {
    if (block.type === "table") {
      for (const row of block.rows) for (const cell of row.cells) lines.push(...text(cell.blocks));
      continue;
    }
    let line = "";
    for (const c of block.children) {
      for (const r of c.type === "run" ? [c] : c.runs) {
        for (const rc of r.content) {
          if (rc.kind === "text") line += rc.text;
          if (rc.kind === "ruby") line += `${rc.base.content.map((x) => (x.kind === "text" ? x.text : "")).join("")}(${rc.rt.content.map((x) => (x.kind === "text" ? x.text : "")).join("")})`;
        }
      }
    }
    lines.push(line);
  }
  return lines;
}

async function bodyText(body: string): Promise<{ whole: string[]; streamed: string[] }> {
  const docx = makeDocx({ "word/document.xml": wrapDocument(body).replace("<w:document ", `<w:document ${MC} `) });
  const whole = text(DocxDocument.load(docx).sections.flatMap((s) => s.blocks));
  const opened = await openStreamingDocument(bytesSource(docx));
  const streamed = text(opened.doc.sections.flatMap((s) => [...s.blocks]));
  return { whole, streamed };
}

describe("wrapped content", () => {
  const cases: Record<string, [string, string[]]> = {
    "a custom XML element around paragraphs": [`<w:customXml w:element="a"><w:customXmlPr/>${p("one")}${p("two")}</w:customXml>${p("three")}`, ["one", "two", "three"]],
    "a custom XML element in a content control": [`<w:sdt><w:sdtContent><w:customXml w:element="a">${p("one")}</w:customXml></w:sdtContent></w:sdt>${p("two")}`, ["one", "two"]],
    "a custom XML element around a run": [`<w:p>${run("one ")}<w:customXml w:element="a">${run("two")}</w:customXml></w:p>`, ["one two"]],
    "a custom XML element around a row": [tbl(`<w:customXml w:element="r"><w:tr>${tc(p("cell"))}</w:tr></w:customXml>`), ["cell"]],
    "a custom XML element around a cell": [tbl(`<w:tr><w:customXml w:element="c">${tc(p("cell"))}</w:customXml></w:tr>`), ["cell"]],
    "a content control around a cell": [tbl(`<w:tr><w:sdt><w:sdtPr/><w:sdtContent>${tc(p("cell"))}</w:sdtContent></w:sdt></w:tr>`), ["cell"]],
    "w:dir and w:bdo around runs": [`<w:p>${run("a ")}<w:dir w:val="rtl">${run("b ")}</w:dir><w:bdo w:val="ltr">${run("c")}</w:bdo></w:p>`, ["a b c"]],
    "an mc:AlternateContent around paragraphs: its Choice": [`<mc:AlternateContent><mc:Choice Requires="w14">${p("choice")}</mc:Choice><mc:Fallback>${p("fallback")}</mc:Fallback></mc:AlternateContent>${p("after")}`, ["choice", "after"]],
    "an mc:AlternateContent around paragraphs with only a Fallback": [`<mc:AlternateContent><mc:Fallback>${p("fallback")}</mc:Fallback></mc:AlternateContent>`, ["fallback"]],
    "an mc:AlternateContent around runs": [`<w:p>${run("a ")}<mc:AlternateContent><mc:Choice Requires="w14">${run("choice")}</mc:Choice><mc:Fallback>${run("fallback")}</mc:Fallback></mc:AlternateContent></w:p>`, ["a choice"]],
    "an mc:AlternateContent in a run": [`<w:p><w:r><w:t xml:space="preserve">a </w:t><mc:AlternateContent><mc:Choice Requires="w14"><w:t>choice</w:t></mc:Choice></mc:AlternateContent></w:r></w:p>`, ["a choice"]],
    "a ruby of several runs": [`<w:p><w:r><w:ruby><w:rubyPr/><w:rt>${run("r1")}${run("r2")}</w:rt><w:rubyBase>${run("b1")}${run("b2")}</w:rubyBase></w:ruby></w:r></w:p>`, ["b1b2(r1r2)"]],
  };
  for (const [name, [body, expected]] of Object.entries(cases)) {
    it(`reads ${name}`, async () => {
      const { whole, streamed } = await bodyText(body);
      expect(whole).toEqual(expected);
      expect(streamed).toEqual(expected);
    });
  }

  it("ends a section at a paragraph inside a custom XML element", async () => {
    const body = `<w:customXml w:element="a"><w:p><w:pPr><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:pPr>${run("one")}</w:p></w:customXml>${p("two")}`;
    const docx = makeDocx({ "word/document.xml": wrapDocument(body) });
    expect(DocxDocument.load(docx).sections.map((s) => text(s.blocks))).toEqual([["one"], ["two"]]);
    const opened = await openStreamingDocument(bytesSource(docx));
    expect(opened.doc.sections.map((s) => text([...s.blocks]))).toEqual([["one"], ["two"]]);
  });

  it("reads a custom XML element in a header", () => {
    const header = `<?xml version="1.0"?>\n<w:hdr ${W_NS}><w:customXml w:element="h">${p("header")}</w:customXml></w:hdr>`;
    const rels = `<?xml version="1.0"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdH" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/></Relationships>`;
    const body = p("body") + `<w:sectPr><w:headerReference xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" w:type="default" r:id="rIdH"/></w:sectPr>`;
    const doc = DocxDocument.load(makeDocx({ "word/document.xml": wrapDocument(body), "word/header1.xml": header, "word/_rels/document.xml.rels": rels }));
    expect(text(doc.headers.get("rIdH")!.blocks)).toEqual(["header"]);
  });

  it("refuses a w:subDoc, whose content is in another file", async () => {
    const rels = `<?xml version="1.0"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdS" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/subDocument" Target="file:///C:/part.docx" TargetMode="External"/></Relationships>`;
    const body = `${p("one")}<w:p><w:subDoc xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rIdS"/></w:p>`;
    const docx = makeDocx({ "word/document.xml": wrapDocument(body), "word/_rels/document.xml.rels": rels });
    const message = "unsupported content: the document includes a subdocument (w:subDoc, file:///C:/part.docx) that Word reads from another file";
    expect(() => DocxDocument.load(docx)).toThrow(message);
    await expect(openStreamingDocument(bytesSource(docx))).rejects.toThrow(message);
  });
});
