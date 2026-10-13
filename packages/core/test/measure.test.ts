import { afterEach, describe, expect, it, vi } from "vitest";
import { ApproxMeasurer, CanvasMeasurer, cssFont, fontWidthScale, normalizeFamily } from "../src/layout/measure.js";

afterEach(() => vi.unstubAllGlobals());

describe("CanvasMeasurer", () => {
  it("uses the font-table alias for SimSun", () => {
    expect(normalizeFamily("宋体", false, false).family).toBe("SimSun");
  });

  it("lays Liberation faces out in the Office faces Word replaces them with", () => {
    // probe3-lo-provenance: Word's PDF embeds Times New Roman and Arial for
    // LibreOffice's Liberation Serif / Sans, even where those are installed.
    expect(normalizeFamily("Liberation Serif", false, false).family).toBe("Times New Roman");
    expect(normalizeFamily("Liberation Sans Bold", false, false)).toEqual({ family: "Arial", bold: true, italic: false });
    expect(normalizeFamily("Liberation Mono", false, false).family).toBe("Courier New");
  });

  it("uses the system UI face as the width-compatible Aptos fallback", () => {
    expect(cssFont({ family: "Aptos", size: 44 / 3, bold: false, italic: false })).toContain(
      "Aptos, system-ui, sans-serif",
    );
    expect(fontWidthScale({ family: "Aptos", size: 44 / 3, bold: false, italic: false })).toBe(0.935);
    expect(new ApproxMeasurer().width("TECHNICAL BRIEF", {
      family: "Aptos",
      size: 44 / 3,
      bold: false,
      italic: false,
    })).toBeLessThan(122.8);
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
