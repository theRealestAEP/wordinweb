import { spawn } from "node:child_process";
import { mkdtempSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { fileURLToPath } from "node:url";
import { afterAll, describe, expect, it } from "vitest";
import { emfDib, grayPng, pictBitmap, pictureDocx, pngHeader, wmfDib } from "./pictures.js";

// The picture budgets (pdf/images.ts): a picture's decoders hold at most
// BITMAP_BYTES of pixels (64 MB, 16,777,216 pixels of RGBA), taken before
// each bitmap is allocated; a streamed PNG may have at most 16 times the
// pixels, averaged down as it streams. Each document holds one picture
// just under, at or over a limit, or one whose header claims far more,
// and is rendered in a child process with a 1024 MB heap
// (picture-limits-child.ts), which reports the peak RSS, array buffers and
// external memory. A picture over a limit is listed as unsupported, and the
// render's peak is that of the same document with no picture at all.

const dir = mkdtempSync(join(tmpdir(), "picture-limits-test-"));
afterAll(() => rmSync(dir, { recursive: true, force: true }));

const SIDE = 4096; // 4096 x 4096 is exactly BITMAP_BYTES of RGBA.
const files: Record<string, [string, Uint8Array | undefined]> = {
  "no picture": ["png", undefined],
  "PICT bitmap under": ["pict", pictBitmap(SIDE, SIDE - 1)],
  "PICT bitmap at": ["pict", pictBitmap(SIDE, SIDE)],
  "PICT bitmap over": ["pict", pictBitmap(SIDE, SIDE + 1)],
  "PICT bitmap claiming 16383 x 32767": ["pict", pictBitmap(16383, 32767, 0)],
  "DIB under": ["wmf", wmfDib(SIDE, SIDE - 1)],
  "DIB at": ["wmf", wmfDib(SIDE, SIDE)],
  "DIB over": ["wmf", wmfDib(SIDE, SIDE + 1)],
  "DIB claiming 30000 x 30000": ["wmf", wmfDib(30000, 30000, 1)],
  // Interlaced, so decoded whole: 4 bytes a pixel of 8-bit gray.
  "PNG interlaced under": ["png", grayPng(SIDE, SIDE - 1, true)],
  "PNG interlaced at": ["png", grayPng(SIDE, SIDE, true)],
  "PNG interlaced over": ["png", grayPng(SIDE, SIDE + 1, true)],
  "PNG streamed over": ["png", pngHeader(16384, 16385)],
  "EMF bitmap claiming 8192 x 8192": ["emf", emfDib(8192, 8192)],
};
for (const [name, [ext, bytes]] of Object.entries(files)) writeFileSync(join(dir, `${name}.docx`), pictureDocx(ext, bytes));

interface Outcome {
  unsupportedImages: string[];
  pages: number;
  peakRssMB: number;
  arrayBuffersMB: number;
  externalMB: number;
  ms: number;
}

function render(name: string): Promise<Outcome> {
  const viteNode = fileURLToPath(new URL("../../../node_modules/.bin/vite-node", import.meta.url));
  const child = fileURLToPath(new URL("./picture-limits-child.ts", import.meta.url));
  return new Promise((resolve, reject) => {
    const run = spawn(process.execPath, ["--max-old-space-size=1024", viteNode, child, join(dir, `${name}.docx`)], { stdio: ["ignore", "pipe", "pipe"] });
    let out = "";
    let err = "";
    run.stdout.on("data", (d) => (out += d));
    run.stderr.on("data", (d) => (err += d));
    run.on("close", (status) => (status === 0 ? resolve(JSON.parse(out.trim().split("\n").pop() as string)) : reject(new Error(`${name}: exit ${status}: ${err}`))));
  });
}

describe("picture limits", () => {
  it("draws a bitmap up to the budget, and refuses one over it before it is allocated", async () => {
    const names = Object.keys(files);
    const outcomes = Object.fromEntries(await Promise.all(names.map(async (name) => [name, await render(name)] as const)));
    const none = outcomes["no picture"];
    for (const name of names.slice(1)) {
      const outcome = outcomes[name];
      const part = `word/media/image1.${files[name][0]}`;
      expect(outcome.pages, name).toBe(none.pages);
      if (/ (under|at)$/.test(name)) {
        expect(outcome.unsupportedImages, name).toEqual([]);
        // The bitmap, its samples and its PDF stream: a few times 64 MB at most.
        expect(outcome.peakRssMB, `${name}: peak RSS in MB`).toBeLessThan(none.peakRssMB + 320);
        expect(outcome.ms, `${name}: ms`).toBeLessThan(30_000);
      } else {
        expect(outcome.unsupportedImages, name).toEqual([part]);
        expect(outcome.peakRssMB, `${name}: peak RSS in MB`).toBeLessThan(none.peakRssMB + 48);
        expect(outcome.arrayBuffersMB, `${name}: array buffers in MB`).toBeLessThan(none.arrayBuffersMB + 48);
        expect(outcome.externalMB, `${name}: external in MB`).toBeLessThan(none.externalMB + 48);
        expect(outcome.ms, `${name}: ms`).toBeLessThan(10_000);
      }
    }
  }, 180_000);
});
