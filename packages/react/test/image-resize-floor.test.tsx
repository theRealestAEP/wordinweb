// @vitest-environment jsdom
/**
 * USER-REPORTED: "the image disappears when I move it."
 *
 * It was not a move. Selection handles sit on the image's EDGES — exactly
 * where you aim to drag a picture — so a mis-grab resized instead of moved,
 * and one gesture collapsed a 624x351 photo to 624x8. That is indistinguishable
 * from the image being gone, and `resizeDrawing` has no inverse, so Cmd+Z could
 * not bring it back (measured in a real browser before the fix).
 *
 * The floor exists so the object stays GRABBABLE: below roughly three
 * handle-widths per axis there is nothing left to aim at, and the shrink
 * becomes a one-way trip. These pin both halves of that rule — a drag cannot
 * produce a sub-floor object, and an object that is ALREADY smaller is left
 * alone rather than forcibly grown.
 */
import { describe, expect, it } from "vitest";
import { createElement, act } from "react";
import { createRoot } from "react-dom/client";
import { DocxView, type DocxViewApi } from "../src/index.js";
import { blankDocxBytes } from "@wordinweb/server";

const PNG = new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10, 1, 2, 3, 4]);
const glob = globalThis as unknown as Record<string, unknown>;
if (typeof URL.createObjectURL !== "function") {
  URL.createObjectURL = () => "blob:resize-floor";
  URL.revokeObjectURL = () => {};
}
async function tick(ms = 5) { await act(async () => { await new Promise<void>((r) => setTimeout(r, ms)); }); }

/** Mount an editable local document holding one image of the given natural size. */
async function mountWithImage(naturalW: number, naturalH: number) {
  glob.createImageBitmap = async () => ({ width: naturalW, height: naturalH, close() {} });
  const container = document.createElement("div");
  document.body.appendChild(container);
  const root = createRoot(container);
  const seen: { api: DocxViewApi | null } = { api: null };
  await act(async () => {
    root.render(createElement(DocxView, {
      source: blankDocxBytes(), editable: true,
      onReady: (api: DocxViewApi) => { seen.api = api; },
    }));
  });
  for (let i = 0; i < 30 && !container.querySelector(".dxw-page"); i++) await tick();
  const page = container.querySelector<HTMLElement>(".dxw-page")!;
  const span = page.querySelector("span") ?? page;
  await act(async () => {
    const o = { bubbles: true, cancelable: true, clientX: 10, clientY: 10, button: 0 };
    span.dispatchEvent(new MouseEvent("mousedown", o));
    span.dispatchEvent(new MouseEvent("mouseup", o));
  });
  await tick();
  await act(async () => { await seen.api!.insertImage(new Blob([PNG], { type: "image/png" })); });
  await tick(20);
  const img = () => container.querySelector<HTMLElement>(".dxw-page img");
  expect(img(), "the image should have painted").toBeTruthy();
  /** Select it, then drag one handle by (dx, dy). */
  const dragHandle = async (dir: string, dx: number, dy: number) => {
    await act(async () => {
      const o = { bubbles: true, cancelable: true, clientX: 40, clientY: 40, button: 0 };
      img()!.dispatchEvent(new MouseEvent("mousedown", o));
      document.dispatchEvent(new MouseEvent("mouseup", { ...o, bubbles: true }));
    });
    await tick();
    const handle = container.querySelector<HTMLElement>(`[data-dxw-img-handle="${dir}"]`);
    expect(handle, `the ${dir} handle should exist once selected`).toBeTruthy();
    await act(async () => {
      handle!.dispatchEvent(new MouseEvent("mousedown", { bubbles: true, cancelable: true, clientX: 200, clientY: 200, button: 0 }));
      document.dispatchEvent(new MouseEvent("mousemove", { bubbles: true, clientX: 200 + dx, clientY: 200 + dy }));
      document.dispatchEvent(new MouseEvent("mouseup", { bubbles: true, clientX: 200 + dx, clientY: 200 + dy }));
    });
    await tick(20);
  };
  const size = () => {
    const el = img();
    return el ? { w: parseFloat(el.style.width), h: parseFloat(el.style.height) } : null;
  };
  return { container, size, dragHandle, unmount: async () => { await act(async () => { root.unmount(); }); } };
}

// Three handle-widths, mirroring MIN_DRAG_SIZE_PX in editor.ts. Written out
// rather than imported because it is private — if the editor's rule changes,
// this number has to be reconsidered deliberately.
const FLOOR = 27;

describe("a resize drag cannot annihilate a picture", () => {
  it("the SOUTH edge dragged far up stops at the floor (was 624x8)", async () => {
    const ed = await mountWithImage(800, 450);
    const before = ed.size()!;
    expect(before.h).toBeGreaterThan(FLOOR * 2); // a normal, large photo
    await ed.dragHandle("s", 0, -4000); // yank the bottom edge way past the top
    const after = ed.size()!;
    expect(after, "the image must still be there").not.toBeNull();
    expect(after.h, "height must not collapse below the grabbable floor").toBeGreaterThanOrEqual(FLOOR);
    expect(after.w, "an edge drag must not touch the other axis").toBeCloseTo(before.w, 1);
    await ed.unmount();
  });

  it("a CORNER dragged inward stops at the floor AND keeps the aspect lock", async () => {
    const ed = await mountWithImage(800, 450);
    const before = ed.size()!;
    await ed.dragHandle("se", -4000, -4000);
    const after = ed.size()!;
    expect(after.w).toBeGreaterThanOrEqual(FLOOR);
    expect(after.h).toBeGreaterThanOrEqual(FLOOR);
    // Aspect preserved: the floor is applied to the SCALE, not per-axis, so
    // clamping cannot quietly distort the picture.
    expect(after.w / after.h).toBeCloseTo(before.w / before.h, 1);
    await ed.unmount();
  });

  it("an image ALREADY smaller than the floor is left alone, never grown", async () => {
    // The floor bounds what a drag may PRODUCE; it is not a minimum the
    // document must satisfy. A genuinely tiny icon must not be inflated just
    // because someone touched its handle.
    const ed = await mountWithImage(20, 12);
    const before = ed.size()!;
    expect(before.w, "this fixture must actually be sub-floor").toBeLessThan(FLOOR);
    await ed.dragHandle("se", -4000, -4000);
    const after = ed.size()!;
    expect(after.w, "a sub-floor image must not be grown to the floor").toBeLessThanOrEqual(before.w + 0.5);
    expect(after.w, "…nor shrunk further into unreachability").toBeCloseTo(before.w, 1);
    await ed.unmount();
  });
});
