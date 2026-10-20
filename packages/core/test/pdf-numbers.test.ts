import { execFileSync } from "node:child_process";
import { mkdtempSync, writeFileSync } from "node:fs";
import { createRequire } from "node:module";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { Writable } from "node:stream";
import { inflateSync } from "node:zlib";
import { describe, expect, it } from "vitest";
import { RecordingCanvas } from "../src/pdf/canvas.js";
import { FontSet } from "../src/pdf/fonts.js";
import { convertMetafile } from "../src/pdf/metafile.js";
import { renderDocxToPdf } from "../src/pdf/render.js";
import { Resources } from "../src/pdf/resources.js";
import { PdfWriter, finiteOps, num } from "../src/pdf/writer.js";
import { bytesSource } from "../src/stream/source.js";
import { emf, emfPlus, emfPlusDriverString, emfPlusFont, pictureDocx } from "./pictures.js";

/*
 * Numbers in the PDF's content: never in exponent form (PDF has none, and
 * Poppler drops the rest of a page's content at one), and never a point a
 * metafile player misread: the recording canvas leaves out what reaches
 * past ten canvas sizes, and the picture is reported as unsupported.
 */

const require = createRequire(import.meta.url);
const fonts = new FontSet([dirname(require.resolve("@fontsource/carlito/files/carlito-latin-400-normal.woff"))]);

function resources(): Resources {
  return new Resources(new PdfWriter(new Writable({ write: (_chunk, _encoding, callback) => callback() })));
}

let hasPdftotext: boolean;
try {
  execFileSync("pdftotext", ["-v"], { stdio: "ignore" });
  hasPdftotext = true;
} catch {
  hasPdftotext = false;
}

describe("num", () => {
  it("writes every finite number in plain decimal, to three places", () => {
    expect(num(0)).toBe("0");
    expect(num(-0)).toBe("0");
    expect(num(12)).toBe("12");
    expect(num(-12.5)).toBe("-12.5");
    expect(num(0.1 + 0.2)).toBe("0.3");
    expect(num(1 / 3)).toBe("0.333");
    expect(num(1e-7)).toBe("0");
    expect(num(-1e-7)).toBe("0");
    expect(num(-0.0004)).toBe("0");
    expect(num(0.0005)).toBe("0.001");
    expect(num(123456789.125)).toBe("123456789.125");
    expect(num(1e21)).toBe("1000000000000000000000");
    expect(num(-2.6040111314378413e34)).toMatch(/^-26040111314378413\d{18}$/);
    expect(num(Number.MAX_VALUE)).toMatch(/^1797693134862315\d{293}$/);
  });

  it("never gives exponent form", () => {
    for (const value of [1e20, 1e21, 2.6e25, -1e30, 1e300, 1e-5, -3.3e-10, 7e-21]) expect(num(value)).toMatch(/^-?\d+(\.\d+)?$/);
  });
});

describe("finiteOps", () => {
  it("leaves out each entry with a NaN or an infinity, all but its q, Q, BT and ET", () => {
    const ops = ["q 0 0 1 rg 10 10 m 20 NaN l S Q", "q", "1 0 0 1 Infinity 5 cm /Im1 Do Q", "BT /F1 12 Tf 1 0 0 1 -Infinity 5 Tm <0001> Tj ET", "q 1 g 0 0 5 5 re f Q"];
    expect(finiteOps(ops)).toEqual(["q Q", "q", "Q", "BT ET", "q 1 g 0 0 5 5 re f Q"]);
  });

  it("reads no number in a string", () => {
    const ops = ["BT /F1 12 Tf (NaN Infinity) Tj ET", "[<4E614E>] TJ"];
    expect(finiteOps(ops)).toEqual(ops);
  });

  it("finds what num() writes for NaN and the infinities", () => {
    expect([num(Number.NaN), num(Infinity), num(-Infinity)]).toEqual(["NaN", "Infinity", "-Infinity"]);
    expect(finiteOps([`${num(Number.NaN)} 0 m`])).toEqual([""]);
  });
});

describe("RecordingCanvas", () => {
  it("leaves out a paint, text or image with a point out of reach, and counts it", () => {
    const canvas = new RecordingCanvas(100, 50, fonts, resources());
    const ctx = canvas.getContext("2d")!;
    ctx.fillRect(10, 10, 20, 20);
    ctx.fillRect(10, 10, 2e25, 20);
    ctx.beginPath();
    ctx.moveTo(0, 0);
    ctx.lineTo(Number.NaN, 5);
    ctx.stroke();
    ctx.fillText("far", 1e30, 10);
    ctx.fillText("near", 10, 10);
    ctx.setTransform(1e30, 0, 0, 1e30, 0, 0);
    ctx.drawImage({ pixels: ctx.createImageData(1, 1), width: 1, height: 1 }, 0, 0);
    expect(canvas.dropped).toBe(4);
    expect(canvas.images).toHaveLength(0);
    expect(canvas.ops).toHaveLength(2);
    expect(canvas.ops.join("\n")).not.toMatch(/e[+-]\d|NaN|Infinity/);
  });

  it("keeps points within ten canvas sizes of its origin", () => {
    const canvas = new RecordingCanvas(100, 50, fonts, resources());
    const ctx = canvas.getContext("2d")!;
    ctx.fillRect(-900, -900, 1800, 1800);
    expect(canvas.dropped).toBe(0);
    ctx.fillRect(0, 0, 1001, 1);
    expect(canvas.dropped).toBe(1);
  });

  it("clamps a clip out of reach to it", () => {
    const canvas = new RecordingCanvas(100, 50, fonts, resources());
    const ctx = canvas.getContext("2d")!;
    ctx.rect(-1e30, -1e30, 2e30, 2e30);
    ctx.clip();
    expect(canvas.dropped).toBe(0);
    expect(canvas.ops).toEqual(["-1000 -1000 m 1000 -1000 l 1000 1000 l -1000 1000 l h W n"]);
  });

  it("leaves out a clip out of reach with a slanted edge, which clamping would move", () => {
    const canvas = new RecordingCanvas(100, 50, fonts, resources());
    const ctx = canvas.getContext("2d")!;
    ctx.moveTo(-2000, 0);
    ctx.lineTo(2000, 0);
    ctx.lineTo(0, 50);
    ctx.clip();
    expect(canvas.dropped).toBe(1);
    expect(canvas.ops).toEqual([]);
  });
});

describe("a metafile with a text origin out of reach", () => {
  it("gives no picture for a text origin out of reach, and the text after the picture stays", async () => {
    const bytes = emf([emfPlus([emfPlusFont(0, 10), emfPlusDriverString("Away", 2.6e25, 30)])]);
    expect(await convertMetafile("emf", bytes, fonts, resources())).toBeNull();
    const chunks: Buffer[] = [];
    const out = new Writable({
      write(chunk: Buffer, _encoding, callback) {
        chunks.push(chunk);
        callback();
      },
    });
    const result = await renderDocxToPdf(bytesSource(pictureDocx("emf", bytes)), out);
    expect(result.unsupportedImages).toEqual(["word/media/image1.emf"]);
    const pdf = Buffer.concat(chunks);
    // The page's content stream: deflated, with no other dictionary entries.
    const content = [...pdf.toString("latin1").matchAll(/<< {2}\/Filter \/FlateDecode \/Length (\d+) >>\nstream\n/g)]
      .map((m) => inflateSync(pdf.subarray(m.index! + m[0].length, m.index! + m[0].length + Number(m[1]))).toString("latin1"))
      .join("\n");
    expect(content).toMatch(/ TJ/);
    expect(content).not.toMatch(/\d[eE][+-]\d|NaN|Infinity/);
    if (!hasPdftotext) return;
    const path = join(mkdtempSync(join(tmpdir(), "emf-numbers-")), "out.pdf");
    writeFileSync(path, pdf);
    const text = execFileSync("pdftotext", [path, "-"], { encoding: "utf8" });
    expect(text).toContain("After the picture.");
  });
});
