import { Writable } from "node:stream";
import { describe, expect, it } from "vitest";
import { PdfRenderError, unsupportedScript } from "../src/pdf/errors.js";
import { FontSet, defaultFontDirs } from "../src/pdf/fonts.js";
import { renderDocxToPdf } from "../src/pdf/render.js";
import { bytesSource } from "../src/stream/source.js";
import { W_NS, makeDocx, p, wrapDocument } from "./helpers.js";

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
  it("names the scripts a shaper is needed for, and lets the others through", () => {
    expect(unsupportedScript("Hello, world")).toBeUndefined();
    expect(unsupportedScript("שלום עליכם פֿונעם")).toBeUndefined();
    expect(unsupportedScript("Привет мир")).toBeUndefined();
    expect(unsupportedScript("漢字かな")).toBeUndefined();
    expect(unsupportedScript("مرحبا")).toBe("Arabic");
    expect(unsupportedScript("ﻣﺮﺣﺒﺎ")).toBe("Arabic");
    expect(unsupportedScript("नमस्ते")).toBe("Devanagari");
    expect(unsupportedScript("สวัสดี")).toBe("Thai");
    expect(unsupportedScript("ສະບາຍດີ")).toBe("Lao");
    expect(unsupportedScript("မင်္ဂလာပါ")).toBe("Myanmar");
    expect(unsupportedScript("ᠮᠣᠩᠭᠣᠯ")).toBe("Mongolian");
    expect(unsupportedScript("ߒߞߏ")).toBe("N'Ko");
  });
});

describe("renderDocxToPdf errors", () => {
  it("refuses Arabic text with unsupported_script before any byte is written, naming the paragraph", async () => {
    const docx = makeDocx({ "word/document.xml": wrapDocument(p("Latin first page") + `<w:p><w:r><w:br w:type="page"/></w:r></w:p>` + p("ثم العربية هنا")) });
    const { error, bytes } = await failure(docx);
    expect(error.code).toBe("unsupported_script");
    expect(error.message).toContain("Arabic");
    expect(error.message).toContain("paragraph 3");
    expect(error.message).toContain("العربية");
    expect(bytes).toBe(0);
  });

  it("refuses Thai text in a footer", async () => {
    const footer = `<?xml version="1.0"?>\n<w:ftr ${W_NS}><w:p><w:r><w:t>สวัสดี</w:t></w:r></w:p></w:ftr>`;
    const rels = `<?xml version="1.0"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdF" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/></Relationships>`;
    const body = p("Latin") + `<w:sectPr><w:footerReference xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" w:type="default" r:id="rIdF"/></w:sectPr>`;
    const docx = makeDocx({ "word/document.xml": wrapDocument(body), "word/footer1.xml": footer, "word/_rels/document.xml.rels": rels });
    const { error } = await failure(docx);
    expect(error.code).toBe("unsupported_script");
    expect(error.message).toContain("Thai");
    expect(error.message).toContain("footer");
  });

  it("refuses an Indic comment only when comments are printed", async () => {
    const comments = `<?xml version="1.0"?>\n<w:comments ${W_NS}><w:comment w:id="0" w:author="A"><w:p><w:r><w:t>नमस्ते</w:t></w:r></w:p></w:comment></w:comments>`;
    const body = `<w:p><w:commentRangeStart w:id="0"/><w:r><w:t>Latin</w:t></w:r><w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r></w:p>`;
    const docx = makeDocx({ "word/document.xml": wrapDocument(body), "word/comments.xml": comments });
    const { error, bytes } = await failure(docx, { comments: true });
    expect(error.code).toBe("unsupported_script");
    expect(error.message).toContain("Devanagari");
    expect(bytes).toBe(0);
    const out = new Collect();
    await renderDocxToPdf(bytesSource(docx), out);
    expect(out.bytes).toBeGreaterThan(0);
  });

  it("refuses Arabic deleted text in a footer only when the markup prints it", async () => {
    const footer = `<?xml version="1.0"?>\n<w:ftr ${W_NS}><w:p><w:r><w:t>Latin</w:t></w:r><w:del w:id="1" w:author="A" w:date="2026-01-01T00:00:00Z"><w:r><w:delText>مرحبا</w:delText></w:r></w:del></w:p></w:ftr>`;
    const rels = `<?xml version="1.0"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdF" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/></Relationships>`;
    const body = p("Latin") + `<w:sectPr><w:footerReference xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" w:type="default" r:id="rIdF"/></w:sectPr>`;
    const docx = makeDocx({ "word/document.xml": wrapDocument(body), "word/footer1.xml": footer, "word/_rels/document.xml.rels": rels });
    const { error, bytes } = await failure(docx, { markup: true });
    expect(error.code).toBe("unsupported_script");
    expect(error.message).toContain("Arabic");
    expect(error.message).toContain("footer");
    expect(bytes).toBe(0);
    const out = new Collect();
    await renderDocxToPdf(bytesSource(docx), out);
    expect(out.bytes).toBeGreaterThan(0);
  });

  it("refuses an Arabic change author, in the body and in a header, when the markup names them", async () => {
    const change = (author: string): string => `<w:p><w:ins w:id="1" w:author="${author}" w:date="2026-01-01T00:00:00Z"><w:r><w:t>Latin</w:t></w:r></w:ins></w:p>`;
    const inBody = makeDocx({ "word/document.xml": wrapDocument(change("محمد")) });
    const { error, bytes } = await failure(inBody, { markup: true });
    expect(error.code).toBe("unsupported_script");
    expect(error.message).toContain("paragraph 1");
    expect(error.message).toContain("محمد");
    expect(bytes).toBe(0);
    const out = new Collect();
    await renderDocxToPdf(bytesSource(inBody), out);
    expect(out.bytes).toBeGreaterThan(0);

    const header = `<?xml version="1.0"?>\n<w:hdr ${W_NS}>${change("สมชาย")}</w:hdr>`;
    const rels = `<?xml version="1.0"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdH" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/></Relationships>`;
    const body = p("Latin") + `<w:sectPr><w:headerReference xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" w:type="default" r:id="rIdH"/></w:sectPr>`;
    const inHeader = makeDocx({ "word/document.xml": wrapDocument(body), "word/header1.xml": header, "word/_rels/document.xml.rels": rels });
    const found = await failure(inHeader, { markup: true });
    expect(found.error.code).toBe("unsupported_script");
    expect(found.error.message).toContain("Thai");
    expect(found.error.message).toContain("header");
  });

  it("refuses an Arabic comment author or initials when the comments are printed", async () => {
    const body = `<w:p><w:commentRangeStart w:id="0"/><w:r><w:t>Latin</w:t></w:r><w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r></w:p>`;
    for (const who of ['w:author="فاطمة" w:initials="F"', 'w:author="Fatima" w:initials="ف"']) {
      const comments = `<?xml version="1.0"?>\n<w:comments ${W_NS}><w:comment w:id="0" ${who}><w:p><w:r><w:t>Latin note</w:t></w:r></w:p></w:comment></w:comments>`;
      const docx = makeDocx({ "word/document.xml": wrapDocument(body), "word/comments.xml": comments });
      const { error, bytes } = await failure(docx, { comments: true });
      expect(error.code).toBe("unsupported_script");
      expect(error.message).toContain("Arabic");
      expect(error.message).toContain("comment");
      expect(bytes).toBe(0);
    }
  });

  it("refuses, in the shaper, any string that reaches the page in an unsupported script", () => {
    const fonts = new FontSet(defaultFontDirs());
    const font = { family: "Arial", size: 12, bold: false, italic: false };
    expect(fonts.shape("Latin only", font).length).toBeGreaterThan(0);
    expect(() => fonts.shape("Label मराठी", font)).toThrow(PdfRenderError);
    try {
      fonts.shape("مرحبا", font);
    } catch (error) {
      expect((error as PdfRenderError).code).toBe("unsupported_script");
      expect((error as PdfRenderError).message).toContain("Arabic");
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
