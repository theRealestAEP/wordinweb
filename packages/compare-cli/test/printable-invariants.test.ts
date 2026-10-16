import { execFileSync } from "node:child_process";
import { mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterAll, describe, expect, it } from "vitest";
import { makeDocx, wrapDocument } from "../../core/test/helpers.js";
import { para } from "../../core/test/compare-helpers.js";
import { compareFiles } from "../src/compare.js";
import { checkPrintable, comparePair, formatFailure } from "./invariants.js";
import { generateCase, serialize } from "./fuzz.js";

/*
 * The printable copy's invariants (invariants.ts) over fuzzed pairs
 * (fuzz.ts): a sample here, thousands with scripts/printable-invariants.ts.
 */

const dir = mkdtempSync(join(tmpdir(), "printable-invariants-"));
afterAll(() => rmSync(dir, { recursive: true, force: true }));
let hasPdftotext = true;
try {
  execFileSync("pdftotext", ["-v"], { stdio: "ignore" });
} catch {
  hasPdftotext = false;
}

async function fuzzed(seed: number, pdf: boolean): Promise<string[]> {
  const c = generateCase(seed);
  const work = join(dir, `seed-${seed}`);
  mkdirSync(work, { recursive: true });
  writeFileSync(join(work, "original.docx"), serialize(c.original));
  writeFileSync(join(work, "revised.docx"), serialize(c.revised));
  const report = await comparePair(join(work, "original.docx"), join(work, "revised.docx"), work, { pdf });
  if (report.error) return [`seed ${seed}: ${report.error}`];
  return report.failures.map((f) => `seed ${seed} (edits ${c.edits.map((e) => e.kind).join(",")}): ${formatFailure(f)}`);
}

describe("the printable copy's invariants", () => {
  it("hold over fuzzed pairs, without the PDF", async () => {
    const failures: string[] = [];
    for (let seed = 1; seed <= 80; seed++) failures.push(...(await fuzzed(seed, false)));
    expect(failures).toEqual([]);
  }, 120_000);

  it.skipIf(!hasPdftotext)("hold over fuzzed pairs, with the PDF", async () => {
    const failures: string[] = [];
    for (let seed = 1; seed <= 12; seed++) failures.push(...(await fuzzed(seed, true)));
    expect(failures).toEqual([]);
  }, 120_000);

  it("catch a copy that is the result itself, and one that lost its marks", async () => {
    const original = join(dir, "sanity-o.docx");
    const revised = join(dir, "sanity-r.docx");
    writeFileSync(original, makeDocx({ "word/document.xml": wrapDocument(para("The quick brown fox.") + para("Second.")) }));
    writeFileSync(revised, makeDocx({ "word/document.xml": wrapDocument(para("The quick red fox.") + para("Second.", { pPr: '<w:jc w:val="center"/>' })) }));
    const out = join(dir, "sanity-out.docx");
    const printable = join(dir, "sanity-print.docx");
    await compareFiles(original, revised, out, { author: "A", date: "2026-01-01T00:00:00Z", printable: { path: printable } });
    const result = readFileSync(out);
    expect(checkPrintable(result, readFileSync(printable))).toEqual([]);
    // The result as its own copy: the text shows (invariant 1 holds) but the marks and the markup fail.
    const asIs = checkPrintable(result, result);
    expect(new Set(asIs.map((f) => f.invariant))).toEqual(new Set([2, 3]));
    expect(asIs.map((f) => f.detail)).toContain("revision markup left: <w:del>");
    expect(asIs.some((f) => f.detail === "change bar")).toBe(true);
    // The revised document as the copy: the deleted word is lost.
    const lost = checkPrintable(result, readFileSync(revised));
    expect(lost.map((f) => f.invariant)).toContain(1);
    expect(lost[0].detail).toBe("text differs");
  });
});
