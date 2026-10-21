import { Writable } from "node:stream";
import { beforeAll, describe, expect, it } from "vitest";
import { DocxDocument } from "../src/docx.js";
import { PdfRenderError, unsupportedScript } from "../src/pdf/errors.js";
import { FontSet, defaultFontDirs } from "../src/pdf/fonts.js";
import { renderDocxToPdf } from "../src/pdf/render.js";
import { loadShaper, shaperLoaded } from "../src/pdf/shaper.js";
import { bytesSource } from "../src/stream/source.js";
import { xmlErrorCases } from "./xml-error-fixtures.js";
import { W_NS, makeDocx, p, wrapDocument } from "./helpers.js";

// Mongolian is the one script still refused with the shaper present: its
// lines run top to bottom. Without the harfbuzzjs package, the scripts
// the shaper is for are refused too (pdf-shaping.test.ts covers them).
const MONGOLIAN = "ᠮᠣᠩᠭᠣᠯ ᠬᠡᠯᠡ";

beforeAll(() => loadShaper());

class Collect extends Writable {
  bytes = 0;
  _write(chunk: Buffer, _encoding: string, callback: () => void): void {
    this.bytes += chunk.length;
    callback();
  }
}

async function failure(docx: Uint8Array, options: { comments?: boolean; markup?: boolean } = {}): Promise<{ error: PdfRenderError; bytes: number }> {
  const out = new Collect();
  try {
    await renderDocxToPdf(bytesSource(docx), out, options);
  } catch (error) {
    expect(error).toBeInstanceOf(PdfRenderError);
    return { error: error as PdfRenderError, bytes: out.bytes };
  }
  throw new Error("rendered without an error");
}

describe("unsupportedScript", () => {
  it("names Mongolian, and lets every other script through once the shaper is loaded", () => {
    expect(unsupportedScript("Hello, world")).toBeUndefined();
    expect(unsupportedScript("שלום עליכם פֿונעם")).toBeUndefined();
    expect(unsupportedScript("Привет мир")).toBeUndefined();
    expect(unsupportedScript("漢字かな")).toBeUndefined();
    expect(unsupportedScript("ᠮᠣᠩᠭᠣᠯ")).toBe("Mongolian");
    const shaped = ["مرحبا", "ﻣﺮﺣﺒﺎ", "नमस्ते", "สวัสดี", "ສະບາຍດີ", "မင်္ဂလာပါ", "ߒߞߏ"];
    if (shaperLoaded()) for (const text of shaped) expect(unsupportedScript(text)).toBeUndefined();
    else expect(shaped.map(unsupportedScript)).toEqual(["Arabic", "Arabic", "Devanagari", "Thai", "Lao", "Myanmar", "N'Ko"]);
  });
});

describe("renderDocxToPdf errors", () => {
  it.each(xmlErrorCases)("refuses %s as damaged", async (_name, docx) => {
    const { error, bytes } = await failure(docx);
    expect(error.code).toBe("damaged");
    expect(error.cause).toBeInstanceOf(Error);
    expect(bytes).toBe(0);
  });

  it("refuses Mongolian text with unsupported_script before any byte is written, naming the paragraph", async () => {
    const docx = makeDocx({ "word/document.xml": wrapDocument(p("Latin first page") + `<w:p><w:r><w:br w:type="page"/></w:r></w:p>` + p(`then ${MONGOLIAN}`)) });
    const { error, bytes } = await failure(docx);
    expect(error.code).toBe("unsupported_script");
    expect(error.message).toContain("Mongolian");
    expect(error.message).toContain("paragraph 3");
    expect(error.message).toContain(MONGOLIAN);
    expect(bytes).toBe(0);
  });

  it("refuses Mongolian text in a footer", async () => {
    const footer = `<?xml version="1.0"?>\n<w:ftr ${W_NS}><w:p><w:r><w:t>${MONGOLIAN}</w:t></w:r></w:p></w:ftr>`;
    const rels = `<?xml version="1.0"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdF" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/></Relationships>`;
    const body = p("Latin") + `<w:sectPr><w:footerReference xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" w:type="default" r:id="rIdF"/></w:sectPr>`;
    const docx = makeDocx({ "word/document.xml": wrapDocument(body), "word/footer1.xml": footer, "word/_rels/document.xml.rels": rels });
    const { error } = await failure(docx);
    expect(error.code).toBe("unsupported_script");
    expect(error.message).toContain("Mongolian");
    expect(error.message).toContain("footer");
  });

  it("refuses a Mongolian comment only when comments are printed", async () => {
    const comments = `<?xml version="1.0"?>\n<w:comments ${W_NS}><w:comment w:id="0" w:author="A"><w:p><w:r><w:t>${MONGOLIAN}</w:t></w:r></w:p></w:comment></w:comments>`;
    const body = `<w:p><w:commentRangeStart w:id="0"/><w:r><w:t>Latin</w:t></w:r><w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r></w:p>`;
    const docx = makeDocx({ "word/document.xml": wrapDocument(body), "word/comments.xml": comments });
    const { error, bytes } = await failure(docx, { comments: true });
    expect(error.code).toBe("unsupported_script");
    expect(error.message).toContain("Mongolian");
    expect(bytes).toBe(0);
    const out = new Collect();
    await renderDocxToPdf(bytesSource(docx), out);
    expect(out.bytes).toBeGreaterThan(0);
  });

  it("refuses Mongolian deleted text in a footer only when the markup prints it", async () => {
    const footer = `<?xml version="1.0"?>\n<w:ftr ${W_NS}><w:p><w:r><w:t>Latin</w:t></w:r><w:del w:id="1" w:author="A" w:date="2026-01-01T00:00:00Z"><w:r><w:delText>${MONGOLIAN}</w:delText></w:r></w:del></w:p></w:ftr>`;
    const rels = `<?xml version="1.0"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdF" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/></Relationships>`;
    const body = p("Latin") + `<w:sectPr><w:footerReference xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" w:type="default" r:id="rIdF"/></w:sectPr>`;
    const docx = makeDocx({ "word/document.xml": wrapDocument(body), "word/footer1.xml": footer, "word/_rels/document.xml.rels": rels });
    const { error, bytes } = await failure(docx, { markup: true });
    expect(error.code).toBe("unsupported_script");
    expect(error.message).toContain("Mongolian");
    expect(error.message).toContain("footer");
    expect(bytes).toBe(0);
    const out = new Collect();
    await renderDocxToPdf(bytesSource(docx), out);
    expect(out.bytes).toBeGreaterThan(0);
  });

  it("refuses a Mongolian change author, in the body and in a header, when the markup names them", async () => {
    const change = (author: string): string => `<w:p><w:ins w:id="1" w:author="${author}" w:date="2026-01-01T00:00:00Z"><w:r><w:t>Latin</w:t></w:r></w:ins></w:p>`;
    const inBody = makeDocx({ "word/document.xml": wrapDocument(change(MONGOLIAN)) });
    const { error, bytes } = await failure(inBody, { markup: true });
    expect(error.code).toBe("unsupported_script");
    expect(error.message).toContain("paragraph 1");
    expect(error.message).toContain(MONGOLIAN);
    expect(bytes).toBe(0);
    const out = new Collect();
    await renderDocxToPdf(bytesSource(inBody), out);
    expect(out.bytes).toBeGreaterThan(0);

    const header = `<?xml version="1.0"?>\n<w:hdr ${W_NS}>${change(MONGOLIAN)}</w:hdr>`;
    const rels = `<?xml version="1.0"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdH" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/></Relationships>`;
    const body = p("Latin") + `<w:sectPr><w:headerReference xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" w:type="default" r:id="rIdH"/></w:sectPr>`;
    const inHeader = makeDocx({ "word/document.xml": wrapDocument(body), "word/header1.xml": header, "word/_rels/document.xml.rels": rels });
    const found = await failure(inHeader, { markup: true });
    expect(found.error.code).toBe("unsupported_script");
    expect(found.error.message).toContain("Mongolian");
    expect(found.error.message).toContain("header");
  });

  it("refuses a Mongolian comment author or initials when the comments are printed", async () => {
    const body = `<w:p><w:commentRangeStart w:id="0"/><w:r><w:t>Latin</w:t></w:r><w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r></w:p>`;
    for (const who of [`w:author="${MONGOLIAN}" w:initials="F"`, 'w:author="Fatima" w:initials="ᠮ"']) {
      const comments = `<?xml version="1.0"?>\n<w:comments ${W_NS}><w:comment w:id="0" ${who}><w:p><w:r><w:t>Latin note</w:t></w:r></w:p></w:comment></w:comments>`;
      const docx = makeDocx({ "word/document.xml": wrapDocument(body), "word/comments.xml": comments });
      const { error, bytes } = await failure(docx, { comments: true });
      expect(error.code).toBe("unsupported_script");
      expect(error.message).toContain("Mongolian");
      expect(error.message).toContain("comment");
      expect(bytes).toBe(0);
    }
  });

  it("refuses, in the shaper, any string that reaches the page in an unsupported script", () => {
    const fonts = new FontSet(defaultFontDirs());
    const font = { family: "Arial", size: 12, bold: false, italic: false };
    expect(fonts.shape("Latin only", font).length).toBeGreaterThan(0);
    expect(() => fonts.shape(`Label ${MONGOLIAN}`, font)).toThrow(PdfRenderError);
    try {
      fonts.shape(MONGOLIAN, font);
    } catch (error) {
      expect((error as PdfRenderError).code).toBe("unsupported_script");
      expect((error as PdfRenderError).message).toContain("Mongolian");
    }
  });

  it("refuses a w:altChunk with unsupported_content wherever it would print, and so does DocxDocument.load", async () => {
    const R_NS = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
    const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    const chunk = `<w:altChunk ${R_NS} r:id="rIdAlt"/>`;
    const rels = (...more: string[]): string =>
      `<?xml version="1.0"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdAlt" Type="${REL}/aFChunk" Target="chunk.html"/>${more.join("")}</Relationships>`;
    const tbl = (cell: string): string => `<w:tbl><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc>${cell}</w:tc></w:tr></w:tbl>`;
    const box = `<w:p><w:r><w:pict><v:shape xmlns:v="urn:schemas-microsoft-com:vml" style="width:100pt;height:50pt"><v:textbox><w:txbxContent>${p("box")}${chunk}</w:txbxContent></v:textbox></v:shape></w:pict></w:r></w:p>`;
    const cases: Record<string, Record<string, string>> = {
      body: { "word/document.xml": wrapDocument(p("before") + chunk + p("after")), "word/_rels/document.xml.rels": rels() },
      "table cell": { "word/document.xml": wrapDocument(tbl(p("cell") + chunk)), "word/_rels/document.xml.rels": rels() },
      "content control": { "word/document.xml": wrapDocument(`<w:sdt><w:sdtContent>${p("in")}${chunk}</w:sdtContent></w:sdt>`), "word/_rels/document.xml.rels": rels() },
      "text box": { "word/document.xml": wrapDocument(box), "word/_rels/document.xml.rels": rels() },
      // The placements no parse path walks into: each one's raw XML is checked.
      "tracked insertion": { "word/document.xml": wrapDocument(p("before") + `<w:ins w:id="1" w:author="A">${chunk}</w:ins>`), "word/_rels/document.xml.rels": rels() },
      "tracked deletion": { "word/document.xml": wrapDocument(p("before") + `<w:del w:id="1" w:author="A">${chunk}</w:del>`), "word/_rels/document.xml.rels": rels() },
      "move destination": { "word/document.xml": wrapDocument(p("before") + `<w:moveTo w:id="1" w:author="A">${chunk}</w:moveTo>`), "word/_rels/document.xml.rels": rels() },
      "inside a paragraph": { "word/document.xml": wrapDocument(`<w:p><w:r><w:t>run</w:t></w:r>${chunk}</w:p>`), "word/_rels/document.xml.rels": rels() },
      "mc:Fallback": {
        "word/document.xml": wrapDocument(
          `<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="w14">${p("choice")}</mc:Choice><mc:Fallback>${chunk}</mc:Fallback></mc:AlternateContent>`,
        ),
        "word/_rels/document.xml.rels": rels(),
      },
      "another prefix": { "word/document.xml": wrapDocument(p("before") + `<x:altChunk xmlns:x="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ${R_NS} r:id="rIdAlt"/>`), "word/_rels/document.xml.rels": rels() },
      header: {
        "word/document.xml": wrapDocument(p("body") + `<w:sectPr><w:headerReference ${R_NS} w:type="default" r:id="rIdH"/></w:sectPr>`),
        "word/_rels/document.xml.rels": rels(`<Relationship Id="rIdH" Type="${REL}/header" Target="header1.xml"/>`),
        "word/header1.xml": `<?xml version="1.0"?>\n<w:hdr ${W_NS}>${p("header")}${chunk}</w:hdr>`,
        "word/_rels/header1.xml.rels": rels(),
      },
      footnote: {
        "word/document.xml": wrapDocument(`<w:p><w:r><w:t>body</w:t></w:r><w:r><w:footnoteReference w:id="1"/></w:r></w:p>`),
        "word/_rels/document.xml.rels": rels(`<Relationship Id="rIdF" Type="${REL}/footnotes" Target="footnotes.xml"/>`),
        "word/footnotes.xml": `<?xml version="1.0"?>\n<w:footnotes ${W_NS}><w:footnote w:id="1">${p("note")}${chunk}</w:footnote></w:footnotes>`,
        "word/_rels/footnotes.xml.rels": rels(),
      },
      comment: {
        "word/document.xml": wrapDocument(`<w:p><w:commentRangeStart w:id="0"/><w:r><w:t>body</w:t></w:r><w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r></w:p>`),
        "word/comments.xml": `<?xml version="1.0"?>\n<w:comments ${W_NS}><w:comment w:id="0" w:author="A">${p("comment")}${chunk}</w:comment></w:comments>`,
        "word/_rels/comments.xml.rels": rels(),
      },
    };
    for (const [where, parts] of Object.entries(cases)) {
      const docx = makeDocx({ ...parts, "word/chunk.html": "<html><body><p>chunk</p></body></html>" });
      const { error, bytes } = await failure(docx);
      expect(error.code, where).toBe("unsupported_content");
      expect(error.message, where).toBe("the document embeds an alternative-format chunk (w:altChunk, word/chunk.html) that Word imports on open");
      expect(bytes, where).toBe(0);
      expect(() => DocxDocument.load(docx), where).toThrow("unsupported content: the document embeds an alternative-format chunk");
    }
  });

  it("says not_docx for bytes that are no zip, and damaged for a zip entry that is corrupt", async () => {
    expect((await failure(new Uint8Array(100))).error.code).toBe("not_docx");
    const docx = makeDocx({ "word/document.xml": wrapDocument(p("x")) });
    // Flip bytes inside a deflated part, and in the central directory.
    for (const [from, to] of [[60, 80], [200, 220], [docx.length - 120, docx.length - 100]]) {
      const broken = docx.slice();
      for (let i = from; i < to; i++) broken[i] ^= 0xff;
      const { error } = await failure(broken);
      expect(["damaged", "not_docx"]).toContain(error.code);
      expect(error.cause).toBeInstanceOf(Error);
    }
  });
});
