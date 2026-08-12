import { describe, expect, it } from "vitest";
import { AgentDocument } from "../src/index.js";
import { anchorTextBox, body, makeDocx } from "./helpers.js";

/**
 * `word_document_inspect { kind: "fit" }` — the rendered-output feedback the
 * in-app model asked for after it resized two drawings blind.
 *
 * The layout already measures shape text in order to place it, so these
 * assertions are about SURFACING that measurement, not about computing a new
 * one: every number here is the same number the renderer paints from.
 */

/** Comfortably more text than a 96x48px box can hold. */
const OVERFULL = "Lorem ipsum dolor sit amet consectetur adipiscing elit sed do eiusmod tempor incididunt";

function fitOf(autofit: string, text: string, height = 457200) {
  const shape = anchorTextBox({ id: 1, x: 457200, y: 457200, width: 914400, height, text, autofit });
  const agent = AgentDocument.load(makeDocx(body(`<w:p>${shape}<w:r><w:t>host</w:t></w:r></w:p>`)));
  const result = agent.inspect({ kind: "fit", pages: { start: 1, count: 1 } });
  if (!("drawings" in result)) throw new Error("fit result missing");
  return result;
}

describe("fit inspection", () => {
  it("reports the box, the measured text, and what the box hides", () => {
    const result = fitOf("<a:noAutofit/>", OVERFULL);
    expect(result.drawings).toHaveLength(1);
    const [fit] = result.drawings;
    expect(fit.objectRef).toMatch(/^object:\d+:\d+$/);
    expect(fit.page).toBe(1);
    expect(fit.boxPx).toEqual({ w: 96, h: 48 });
    // The text is measured inside the box's insets, so here it wraps narrower
    // than the frame and runs much taller than it.
    expect(fit.textPx.w).toBeGreaterThan(0);
    expect(fit.textPx.w).toBeLessThanOrEqual(fit.boxPx.w);
    expect(fit.textPx.h).toBeGreaterThan(fit.boxPx.h);
    expect(fit.overflow).toBe(true);
    expect(fit.clippedLines).toBeGreaterThan(0);
    expect(fit.autofit).toBe("none");
  });

  it("reports no overflow when the text fits", () => {
    const fit = fitOf("<a:noAutofit/>", "Short").drawings[0];
    expect(fit.textPx.h).toBeLessThan(fit.boxPx.h);
    expect(fit.overflow).toBe(false);
    expect(fit.clippedLines).toBe(0);
  });

  it("reports resizeShape as resolved: the box grew to the text", () => {
    const fit = fitOf("<a:spAutoFit/>", OVERFULL).drawings[0];
    expect(fit.autofit).toBe("resizeShape");
    // a:spAutoFit recomputes the frame from the measured text, so the box the
    // fit report names is the GROWN one and nothing is left over.
    expect(fit.boxPx.h).toBeGreaterThan(48);
    expect(fit.overflow).toBe(false);
    expect(fit.clippedLines).toBe(0);
  });

  it("reports shrinkText as unresolved, because Word does not apply it", () => {
    // probe-shapefit: fourteen a:normAutofit shapes through desktop Word all
    // paint at their authored size and clip at the box bottom, line for line
    // with the a:noAutofit control. The mode is reported so the model can tell
    // it apart from "none", but it does NOT make the overflow go away.
    const fit = fitOf("<a:normAutofit fontScale=\"62500\" lnSpcReduction=\"20000\"/>", OVERFULL).drawings[0];
    expect(fit.autofit).toBe("shrinkText");
    expect(fit.overflow).toBe(true);
    expect(fit.clippedLines).toBeGreaterThan(0);
  });

  it("reports how full each requested page is", () => {
    const result = fitOf("<a:noAutofit/>", "Short");
    expect(result.pages).toHaveLength(1);
    const [page] = result.pages;
    expect(page.page).toBe(1);
    expect(page.pageBottomPx).toBeGreaterThan(0);
    expect(page.contentBottomPx).toBeGreaterThan(0);
    expect(page.contentBottomPx).toBeLessThan(page.pageBottomPx);
    expect(result.revision).toBe(AgentDocument.load(makeDocx(body("<w:p/>"))).revision);
  });

  it("rejects a malformed page range like every other paged inspection", () => {
    const agent = AgentDocument.load(makeDocx(body("<w:p><w:r><w:t>host</w:t></w:r></w:p>")));
    expect(() => agent.inspect({ kind: "fit", pages: { start: 0, count: 1 } })).toThrow("Invalid page range");
  });

  it("says nothing about drawings that flow no text", () => {
    // A picture has a box but no laid-out text, so it carries no fit entry
    // rather than a zero-sized one that would read as "fits".
    const agent = AgentDocument.load(makeDocx(body("<w:p><w:r><w:t>host</w:t></w:r></w:p>")));
    const result = agent.inspect({ kind: "fit" });
    if (!("drawings" in result)) throw new Error("fit result missing");
    expect(result.drawings).toEqual([]);
  });
});
