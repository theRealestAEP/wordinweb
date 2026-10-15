import { Writable } from "node:stream";
import { describe, expect, it } from "vitest";
import { PdfRenderError, unsupportedScript } from "../src/pdf/errors.js";
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

async function failure(docx: Uint8Array, options: { comments?: boolean } = {}): Promise<{ error: PdfRenderError; bytes: number }> {
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
