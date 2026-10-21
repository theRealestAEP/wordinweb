import { mkdtempSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterAll, describe, expect, it } from "vitest";
import { makeDocx, p, wrapDocument } from "../../core/test/helpers.js";
import { xmlErrorCases } from "../../core/test/xml-error-fixtures.js";
import { compareFiles } from "../src/compare.js";
import { CompareError } from "../src/errors.js";

const dir = mkdtempSync(join(tmpdir(), "compare-xml-errors-"));
afterAll(() => rmSync(dir, { recursive: true, force: true }));
const good = join(dir, "good.docx");
writeFileSync(good, makeDocx({ "word/document.xml": wrapDocument(p("Text")) }));

describe("compareFiles XML errors", () => {
  it("preserves damaged errors from PDF rendering", async () => {
    const printable = join(dir, "printable.docx");
    await expect(compareFiles(good, good, join(dir, "result.docx"), {
      author: "Test",
      printable: { path: printable },
      pdf: { path: join(dir, "result.pdf") },
      onResultWritten: () => writeFileSync(printable, xmlErrorCases[0][1]),
    })).rejects.toMatchObject({ code: "damaged", cause: expect.objectContaining({ name: "PdfRenderError", code: "damaged" }) });
  });

  it.each(xmlErrorCases)("refuses %s as damaged for either input", async (_name, bytes) => {
    const bad = join(dir, "bad.docx");
    writeFileSync(bad, bytes);
    for (const side of ["original", "revised"] as const) {
      const result = compareFiles(side === "original" ? bad : good, side === "revised" ? bad : good, join(dir, "result.docx"), { author: "Test" });
      await expect(result).rejects.toBeInstanceOf(CompareError);
      await expect(result).rejects.toMatchObject({ code: "damaged", file: side, cause: expect.any(Error) });
    }
  });
});
