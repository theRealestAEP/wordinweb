import { spawn } from "node:child_process";
import { mkdtempSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { fileURLToPath } from "node:url";
import { afterAll, describe, expect, it } from "vitest";
import { BITMAP_BYTES } from "../src/pdf/images.js";
import { emfDib, grayPng, pictBitmap, pictureDocx, pngHeader, wmfDib } from "./pictures.js";

// The picture budgets (pdf/images.ts): a picture's decoders hold at most
// BITMAP_BYTES of pixels (64 MB, 16,777,216 pixels of RGBA), taken before
// each bitmap is allocated; a streamed PNG may have at most 16 times the
// pixels, averaged down as it streams. Each document holds one picture
// just under, at or over a limit, or one whose header claims far more. A
// child process with a 1024 MB heap (picture-limits-child.ts) warms up on
// the same document with no picture, then renders that and the document
// with the picture once each, and reports how far the picture's render rose
// over the other in peak RSS, array buffers and external memory, and the
// largest typed array each render made. A picture over a limit is listed as
// unsupported, and makes no array larger than the render with no picture
// makes: its bitmap, 64 MB or more, was never allocated.

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
  controlPages: number;
  rssRiseMB: number;
  arrayBuffersRiseMB: number;
  externalRiseMB: number;
  peakRssMB: number;
  largestArrayMB: number;
  controlLargestArrayMB: number;
  ms: number;
}

const BITMAP_MB = BITMAP_BYTES / (1024 * 1024);

function render(name: string): Promise<Outcome> {
  const viteNode = fileURLToPath(new URL("../../../node_modules/.bin/vite-node", import.meta.url));
  const child = fileURLToPath(new URL("./picture-limits-child.ts", import.meta.url));
  return new Promise((resolve, reject) => {
    const run = spawn(process.execPath, ["--max-old-space-size=1024", "--expose-gc", viteNode, child, join(dir, "no picture.docx"), join(dir, `${name}.docx`)], { stdio: ["ignore", "pipe", "pipe"] });
    let out = "";
    let err = "";
    run.stdout.on("data", (d) => (out += d));
    run.stderr.on("data", (d) => (err += d));
    run.on("close", (status) => (status === 0 ? resolve(JSON.parse(out.trim().split("\n").pop() as string)) : reject(new Error(`${name}: exit ${status}: ${err}`))));
  });
}

describe("picture limits", () => {
  it("draws a bitmap up to the budget, and refuses one over it before it is allocated", async () => {
    // One at a time, so no two renders share the machine.
    const names = Object.keys(files).filter((name) => name !== "no picture");
    for (const name of names) {
      const outcome = await render(name);
      const part = `word/media/image1.${files[name][0]}`;
      expect(outcome.pages, name).toBe(outcome.controlPages);
      if (/ (under|at)$/.test(name)) {
        expect(outcome.unsupportedImages, name).toEqual([]);
        // Its bitmap is the largest array, at most BITMAP_BYTES; with its
        // samples and its PDF stream, the render holds a few times that.
        expect(outcome.largestArrayMB, `${name}: largest array in MB`).toBeLessThanOrEqual(BITMAP_MB);
        expect(outcome.rssRiseMB, `${name}: RSS rise in MB`).toBeLessThan(4 * BITMAP_MB);
        expect(outcome.ms, `${name}: ms`).toBeLessThan(30_000);
      } else {
        expect(outcome.unsupportedImages, name).toEqual([part]);
        // No array of the bitmap's size was made: the largest is the one
        // the render with no picture makes too (a font's, 25 MB).
        expect(outcome.largestArrayMB, `${name}: largest array in MB`).toBe(outcome.controlLargestArrayMB);
        expect(outcome.largestArrayMB, `${name}: largest array in MB`).toBeLessThan(BITMAP_MB / 2);
        // The sampled peaks move by a font file (about 55 MB) from render to
        // render; a decoded bitmap and its samples would add 2 x BITMAP_BYTES.
        expect(outcome.rssRiseMB, `${name}: RSS rise in MB`).toBeLessThan(2 * BITMAP_MB);
        expect(outcome.arrayBuffersRiseMB, `${name}: array buffers rise in MB`).toBeLessThan(2 * BITMAP_MB);
        expect(outcome.externalRiseMB, `${name}: external rise in MB`).toBeLessThan(2 * BITMAP_MB);
        expect(outcome.ms, `${name}: ms`).toBeLessThan(10_000);
      }
    }
  }, 600_000);
});
