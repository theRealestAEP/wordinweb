import { mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { strFromU8, unzipSync } from "fflate";
import { afterAll, describe, expect, it } from "vitest";
import { DocxDocument } from "../../core/src/docx.js";
import { collectRevisions } from "../../core/src/edit/suggest.js";
import { parseXml, serializeXml } from "../../core/src/xml.js";
import { AUTHOR, DATE } from "../../core/test/compare-helpers.js";
import { TRANSITIONAL_DOCUMENT, strictPackage } from "../../core/test/strict-helpers.js";
import { compareFiles } from "../src/compare.js";

const dir = mkdtempSync(join(tmpdir(), "compare-strict-test-"));
afterAll(() => rmSync(dir, { recursive: true, force: true }));

/** compareFiles on the two packages: the result's revisions as collectRevisions sees them, and its parts. */
async function compare(name: string, original: Uint8Array, revised: Uint8Array): Promise<{ counts: Record<string, number>; signatures: string[]; parts: Record<string, string> }> {
  const paths = [join(dir, `${name}-o.docx`), join(dir, `${name}-r.docx`), join(dir, `${name}-out.docx`)];
  writeFileSync(paths[0], original);
  writeFileSync(paths[1], revised);
  const { revisions } = await compareFiles(paths[0], paths[1], paths[2], { author: AUTHOR, date: DATE });
  const result = new Uint8Array(readFileSync(paths[2]));
  const parts: Record<string, string> = {};
  for (const [part, bytes] of Object.entries(unzipSync(result))) if (/\.(xml|rels)$/.test(part)) parts[part] = strFromU8(bytes);
  const signatures = collectRevisions(DocxDocument.load(result)).map((r) => `${r.kind}:${serializeXml(r.el).replace(/ w:id="\d+"/g, "")}`);
  return { counts: revisions, signatures, parts };
}

describe("compareFiles on Strict Open XML", () => {
  it("gives the revisions of the Transitional pair, in a Transitional result", async () => {
    const transitional = await compare("transitional", strictPackage(false), strictPackage(false, "Hello there"));
    const strictOriginal = await compare("strict-original", strictPackage(true), strictPackage(false, "Hello there"));
    const both = await compare("strict-both", strictPackage(true), strictPackage(true, "Hello there"));
    expect(transitional.signatures.length).toBeGreaterThan(0);
    for (const result of [strictOriginal, both]) {
      expect(result.counts).toEqual(transitional.counts);
      expect(result.signatures).toEqual(transitional.signatures);
      for (const [part, xml] of Object.entries(result.parts)) {
        expect(xml, part).not.toContain("purl.oclc.org");
        parseXml(xml);
      }
      // The result's own parts are the Transitional pair's: the body carries
      // the Transitional measures and names, in a Transitional namespace.
      const sectPr = /<w:sectPr>.*<\/w:sectPr>/.exec(result.parts["word/document.xml"])?.[0];
      expect(sectPr).toBe(/<w:sectPr>.*<\/w:sectPr>/.exec(TRANSITIONAL_DOCUMENT)?.[0]);
      expect(result.parts["word/document.xml"]).toContain('<w:tblBorders><w:right w:val="single" w:sz="8" w:space="0" w:color="000000"/></w:tblBorders>');
      expect(result.parts["_rels/.rels"]).toContain("http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument");
    }
  });
});
