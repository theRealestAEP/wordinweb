import { spawn } from "node:child_process";
import { mkdtempSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { fileURLToPath } from "node:url";
import { afterAll, describe, expect, it } from "vitest";
import { giantParagraph, lyingBody, lyingStyles, manyParts, stylesBomb } from "./bombs.js";

// Each crafted file is a few hundred KB on disk and hundreds of MB inflated.
// Every way of opening a .docx must refuse it from the zip's directory, or
// stop inflating it at a bound, before it allocates that much: each way runs
// in a child process with a 1024 MB heap, which reports its error and how
// far its peak RSS rose (zip-limits-child.ts).

const dir = mkdtempSync(join(tmpdir(), "zip-limits-test-"));
afterAll(() => rmSync(dir, { recursive: true, force: true }));

const files = {
  "styles-bomb": stylesBomb(),
  "lying-styles": lyingStyles(),
  "many-parts": manyParts(),
  "giant-paragraph": giantParagraph(),
  "lying-body": lyingBody(),
};
for (const [name, bytes] of Object.entries(files)) writeFileSync(join(dir, `${name}.docx`), bytes);

interface Outcome {
  error: { name: string; code?: string; message: string } | null;
  grewMB: number;
  ms: number;
}

function open(file: keyof typeof files, way: "render" | "load" | "estimate"): Promise<Outcome> {
  const viteNode = fileURLToPath(new URL("../../../node_modules/.bin/vite-node", import.meta.url));
  const child = fileURLToPath(new URL("./zip-limits-child.ts", import.meta.url));
  return new Promise((resolve, reject) => {
    const run = spawn(process.execPath, ["--max-old-space-size=1024", viteNode, child, join(dir, `${file}.docx`), way], { stdio: ["ignore", "pipe", "pipe"] });
    let out = "";
    let err = "";
    run.stdout.on("data", (d) => (out += d));
    run.stderr.on("data", (d) => (err += d));
    run.on("close", (status) => (status === 0 ? resolve(JSON.parse(out.trim().split("\n").pop() as string)) : reject(new Error(`${file} ${way}: exit ${status}: ${err}`))));
  });
}

/** What each way must say, and the most its peak RSS may rise. */
const expected: [keyof typeof files, "render" | "load" | "estimate", Partial<Outcome["error"]>, number][] = [
  ["styles-bomb", "render", { code: "damaged", message: "word/styles.xml is 419430407 bytes, over the 67108864 an XML part may be" }, 64],
  ["styles-bomb", "load", { message: "too large: word/styles.xml is 419430407 bytes, over the 67108864 an XML part may be" }, 64],
  ["styles-bomb", "estimate", { message: "too large: word/styles.xml is 419430407 bytes, over the 67108864 an XML part may be" }, 64],
  // The directory says 1000 bytes: inflating stops just past them.
  ["lying-styles", "render", { code: "damaged", message: "word/styles.xml is corrupt" }, 128],
  ["lying-styles", "load", { message: "word/styles.xml is corrupt" }, 64],
  ["many-parts", "render", { code: "damaged", message: "the zip's parts come to over 536870912 bytes" }, 64],
  ["many-parts", "load", { message: "too large: the zip's parts come to over 536870912 bytes" }, 64],
  ["many-parts", "estimate", { message: "too large: the zip's parts come to over 536870912 bytes" }, 64],
  // The streamed body holds one block at a time, at most the XML limit.
  ["giant-paragraph", "render", { code: "damaged", message: "the body holds a block over 67108864 bytes" }, 320],
  ["giant-paragraph", "load", { message: "too large: word/document.xml is 419430602 bytes, over the 67108864 an XML part may be" }, 64],
  ["lying-body", "render", { code: "damaged", message: "the body holds a block over 67108864 bytes" }, 320],
];

describe("zip limits", () => {
  it("refuses every crafted file in every way, quickly and with no large allocation", async () => {
    const outcomes = await Promise.all(expected.map(([file, way]) => open(file, way)));
    expected.forEach(([file, way, error, maxMB], i) => {
      const outcome = outcomes[i];
      expect(outcome.error, `${file} ${way}`).toMatchObject(error as object);
      expect(outcome.grewMB, `${file} ${way}: peak RSS rise in MB`).toBeLessThan(maxMB);
      expect(outcome.ms, `${file} ${way}: ms`).toBeLessThan(5000);
    });
  }, 120_000);
});
