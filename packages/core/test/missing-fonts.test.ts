// @vitest-environment jsdom
/**
 * detectMissingFonts decides what `onMissingFonts` warns about: faces whose
 * absence moves line breaks away from Word's. A missing Calibri with Carlito
 * loaded is NOT that case — Carlito carries Calibri's exact advances, and the
 * anon-share demo ships it for precisely this reason (issue #18) — so the
 * detector must stay quiet then and speak only when the substitute is
 * missing too. The canvas fake measures a face as present when its name is
 * in the installed set; anything else collapses to the generic at the end of
 * the stack, which is exactly how a real browser measures a missing face.
 */
import { afterEach, describe, expect, it } from "vitest";
import { detectMissingFonts } from "../src/render/fonts.js";
import type { LayoutResult } from "../src/layout/types.js";

function installFonts(installed: string[]) {
  HTMLCanvasElement.prototype.getContext = function () {
    return {
      font: "",
      measureText(text: string) {
        const face = installed.find((name) => this.font.includes(`"${name}"`));
        const per = face ? 10 + installed.indexOf(face) : /monospace/.test(this.font) ? 8 : /serif/.test(this.font) ? 6 : 5;
        return { width: per * text.length } as TextMetrics;
      },
    } as unknown as CanvasRenderingContext2D;
  } as never;
}
const prevGetContext = HTMLCanvasElement.prototype.getContext;
afterEach(() => { HTMLCanvasElement.prototype.getContext = prevGetContext; });

function layoutIn(family: string): LayoutResult {
  return {
    pages: [{ items: [{ kind: "text", text: "Hello", font: { family, size: 16 } }] }],
  } as unknown as LayoutResult;
}

describe("detectMissingFonts", () => {
  it("stays quiet about Calibri when Carlito is loaded", () => {
    installFonts(["Carlito"]);
    expect(detectMissingFonts(layoutIn("Calibri"))).toEqual([]);
    expect(detectMissingFonts(layoutIn("Calibri Light"))).toEqual([]);
  });

  it("reports Calibri when neither it nor Carlito is available", () => {
    installFonts([]);
    expect(detectMissingFonts(layoutIn("Calibri")).map((m) => m.family)).toEqual(["Calibri"]);
  });

  it("still reports a face whose stand-in only approximates Word's widths", () => {
    installFonts(["Gill Sans"]);
    expect(detectMissingFonts(layoutIn("Gill Sans MT")).map((m) => m.family)).toEqual(["Gill Sans MT"]);
  });

  it("says nothing when the face itself is installed", () => {
    installFonts(["Calibri"]);
    expect(detectMissingFonts(layoutIn("Calibri"))).toEqual([]);
  });
});
