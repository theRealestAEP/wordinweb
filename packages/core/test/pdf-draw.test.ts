import { describe, expect, it } from "vitest";
import { embeddingLevels, mirrored } from "../src/pdf/bidi.js";
import { parseColor } from "../src/pdf/draw.js";
import { FontSet, defaultFontDirs } from "../src/pdf/fonts.js";
import { pathToPdf } from "../src/pdf/path.js";

describe("the order FontSet.shape draws a right-to-left run in", () => {
  const fonts = new FontSet(defaultFontDirs());
  const font = { family: "Arial", size: 12, bold: false, italic: false };
  /** The code points of `text`'s glyphs, left to right, as an RTL item draws them. */
  const drawn = (text: string): string => fonts.shape(text, font, true).flatMap((run) => run.codePoints).map((cp) => String.fromCodePoint(cp)).join("");
  const hasHebrew = fonts.shape("א", font).every((run) => !run.gids.includes(0));

  it("resolves the levels of a mixed run and mirrors a bracket", () => {
    expect([...embeddingLevels("אב abc", true)]).toEqual([1, 1, 1, 2, 2, 2]);
    expect(mirrored("(")).toBe(")");
    expect(mirrored("a")).toBeNull();
  });

  it.skipIf(!hasHebrew)("reverses a Hebrew run and mirrors its brackets", () => {
    expect(drawn("אב")).toBe("בא");
    expect(drawn("(אב)")).toBe("(בא)");
  });

  it.skipIf(!hasHebrew)("keeps a Latin word inside an RTL run readable left to right", () => {
    const out = drawn("אב abc גד");
    expect(out.indexOf("abc")).toBeGreaterThan(-1);
    expect(out.startsWith("דג")).toBe(true);
    expect(out.endsWith("בא")).toBe(true);
  });
});

describe("pathToPdf", () => {
  const identity = (x: number, y: number): [number, number] => [x, y];

  it("maps lines, curves and closes through the point transform", () => {
    expect(pathToPdf("M 0 0 L 10 0 Q 10 10 0 10 Z", (x, y) => [x * 2, y * 2])).toBe("0 0 m 20 0 l 20 13.333 13.333 20 0 20 c h");
  });

  it("turns an arc into cubic segments that end at the arc's end point", () => {
    const ops = pathToPdf("M 10 0 A 10 10 0 0 1 0 10", identity);
    expect(ops.startsWith("10 0 m")).toBe(true);
    expect(ops.endsWith("0 10 c")).toBe(true);
    expect(ops.match(/ c/g)?.length).toBe(1);
  });

  it("reads relative commands and the H/V shorthands", () => {
    expect(pathToPdf("m 1 1 h 4 v 4 l -2 -2", identity)).toBe("1 1 m 5 1 l 5 5 l 3 3 l");
  });
});

describe("parseColor", () => {
  it("reads hex, rgb() and rgba()", () => {
    expect(parseColor("#ff0000")).toEqual({ r: 255, g: 0, b: 0, a: 1 });
    expect(parseColor("#0f0")).toEqual({ r: 0, g: 255, b: 0, a: 1 });
    expect(parseColor("rgba(0, 0, 255, 0.5)")).toEqual({ r: 0, g: 0, b: 255, a: 0.5 });
    expect(parseColor("transparent")).toBeNull();
  });
});
