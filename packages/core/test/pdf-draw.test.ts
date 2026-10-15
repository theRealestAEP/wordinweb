import { describe, expect, it } from "vitest";
import { visualOrder } from "../src/pdf/bidi.js";
import { parseColor } from "../src/pdf/draw.js";
import { pathToPdf } from "../src/pdf/path.js";

describe("visualOrder", () => {
  it("reverses a Hebrew run and mirrors its brackets", () => {
    expect(visualOrder("אב")).toBe("בא");
    expect(visualOrder("(אב)")).toBe("(בא)");
  });

  it("keeps a Latin word inside an RTL run readable left to right", () => {
    const out = visualOrder("אב abc גד");
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
