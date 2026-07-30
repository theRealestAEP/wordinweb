import { describe, expect, it } from "vitest";
import { DocxDocument } from "../src/index.js";
import { makeDocx } from "./helpers.js";

/**
 * B11: save() must be byte-deterministic. fflate stamps every zip entry
 * with `new Date()` at 2-second DOS granularity unless an mtime is given,
 * so two saves straddling a boundary used to differ in bytes — flaking
 * every byte-identity check downstream (docHash self-heal comparisons,
 * checkpoint dedup, the scoped-resync byte-compare suite). The fix pins
 * FIXED_ZIP_MTIME; this test pins the fix by saving twice across a
 * guaranteed DOS-tick boundary (>2s apart).
 */
describe("save() byte-determinism (B11 fixed zip mtime)", () => {
  it("two saves more than one DOS-time tick apart are byte-identical", async () => {
    const doc = DocxDocument.load(
      makeDocx({
        "word/document.xml":
          `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
          `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
          `<w:body><w:p><w:r><w:t xml:space="preserve">stable bytes</w:t></w:r></w:p></w:body></w:document>`,
      }),
    );
    const first = doc.save();
    await new Promise((r) => setTimeout(r, 2100)); // DOS timestamps tick every 2s
    const second = doc.save();
    expect(second.length).toBe(first.length);
    expect(Buffer.from(second).equals(Buffer.from(first))).toBe(true);
  }, 10_000);

  it("async save produces the same package bytes", async () => {
    const doc = DocxDocument.load(
      makeDocx({
        "word/document.xml":
          `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
          `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
          `<w:body><w:p><w:r><w:t>async bytes</w:t></w:r></w:p></w:body></w:document>`,
      }),
    );
    expect(Buffer.from(await doc.saveAsync()).equals(Buffer.from(doc.save()))).toBe(true);
  });
});
