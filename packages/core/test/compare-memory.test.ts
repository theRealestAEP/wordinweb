import { describe, expect, it } from "vitest";
import { unzipSync, zipSync, strToU8 } from "fflate";
import { estimateCompareMemory } from "../src/edit/compare/memory.js";
import { zipEntrySizes } from "../src/zip.js";
import { makeDocx, wrapDocument, p } from "./helpers.js";

const docx = (paragraphs: number) =>
  makeDocx({ "word/document.xml": wrapDocument(Array.from({ length: paragraphs }, (_, i) => p(`Paragraph ${i} of the test.`)).join("")) });

describe("zipEntrySizes", () => {
  it("reads every entry's uncompressed size without unzipping", () => {
    const bytes = docx(50);
    const unzipped = unzipSync(bytes);
    const sizes = zipEntrySizes(bytes);
    expect(new Set(sizes.keys())).toEqual(new Set(Object.keys(unzipped)));
    for (const [name, data] of Object.entries(unzipped)) expect(sizes.get(name)).toBe(data.length);
  });

  it("reads a zip with an archive comment", () => {
    const bytes = zipSync({ "a.xml": strToU8("<a/>") }, { comment: "x".repeat(300) } as never);
    expect(zipEntrySizes(bytes).get("a.xml")).toBe(4);
  });

  it("throws on bytes that are not a zip", () => {
    expect(() => zipEntrySizes(new TextEncoder().encode("not a docx at all"))).toThrow(/not a zip/);
  });
});

describe("estimateCompareMemory", () => {
  it("grows with the XML and is lower when the revised document is reused", () => {
    const small = docx(10);
    const large = docx(5000);
    const a = estimateCompareMemory(small, small);
    const b = estimateCompareMemory(large, large);
    expect(b.heapBytes).toBeGreaterThan(a.heapBytes);
    expect(estimateCompareMemory(large, large, { reuseRevised: true }).heapBytes).toBeLessThan(b.heapBytes);
    expect(b.totalBytes).toBeGreaterThan(b.heapBytes);
  });

  it("is higher when an input holds tracked changes", () => {
    const clean = docx(5000);
    const tracked = makeDocx({
      "word/document.xml": wrapDocument(
        Array.from({ length: 5000 }, (_, i) => p(`Paragraph ${i} of the test.`)).join("") +
          '<w:p><w:ins w:id="1" w:author="a"><w:r><w:t>new</w:t></w:r></w:ins></w:p>',
      ),
    });
    for (const options of [{}, { reuseRevised: true }]) {
      const base = estimateCompareMemory(clean, clean, options).heapBytes;
      expect(estimateCompareMemory(tracked, clean, options).heapBytes).toBeGreaterThan(base);
      expect(estimateCompareMemory(clean, tracked, options).heapBytes).toBeGreaterThan(base);
    }
  });

  it("is lower for compareDocx, unless an input needs the whole tree", () => {
    const clean = docx(5000);
    const percentTable = makeDocx({
      "word/document.xml": wrapDocument(
        Array.from({ length: 5000 }, (_, i) => p(`Paragraph ${i} of the test.`)).join("") +
          '<w:tbl><w:tblPr><w:tblW w:w="5000%" w:type="pct"/></w:tblPr><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>',
      ),
    });
    expect(estimateCompareMemory(clean, clean, { compareDocx: true }).heapBytes).toBeLessThan(
      estimateCompareMemory(clean, clean, { reuseRevised: true }).heapBytes,
    );
    expect(estimateCompareMemory(clean, percentTable, { compareDocx: true }).heapBytes).toBe(
      estimateCompareMemory(clean, percentTable, { reuseRevised: true }).heapBytes,
    );
  });
});
