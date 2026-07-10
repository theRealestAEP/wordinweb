import { afterEach, describe, expect, it, vi } from "vitest";
import { CanvasMeasurer, normalizeFamily } from "../src/layout/measure.js";

afterEach(() => vi.unstubAllGlobals());

describe("CanvasMeasurer", () => {
  it("uses the font-table alias for SimSun", () => {
    expect(normalizeFamily("宋体", false, false).family).toBe("SimSun");
  });

  it("measures fractional Word font sizes at 3x and scales the width back", () => {
    const measuredFonts: string[] = [];

    class FakeOffscreenCanvas {
      getContext() {
        return {
          font: "",
          measureText(text: string) {
            measuredFonts.push(this.font);
            const size = Number(/([\d.]+)px/.exec(this.font)?.[1]);
            const truncatedSize = Math.floor(size * 100) / 100;
            return { width: truncatedSize * text.length };
          },
        };
      }
    }

    vi.stubGlobal("OffscreenCanvas", FakeOffscreenCanvas);

    const measurer = new CanvasMeasurer();
    const width = measurer.width("abc", {
      family: "Times New Roman",
      size: 56 / 3,
      bold: false,
      italic: false,
    });

    expect(measuredFonts[0]).toContain("56px");
    expect(width).toBe(56);
  });
});
