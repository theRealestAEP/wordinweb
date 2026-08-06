// @vitest-environment jsdom
/**
 * A crop drag has to READ as a crop.
 *
 * Writing a:srcRect alone selects a smaller part of the bitmap but leaves the
 * picture's frame the size it was, so the surviving content is blown up to
 * refill it — the gesture looks like a zoom, not a trim. Word shrinks the
 * frame to the kept region so the content holds its scale, and moves a
 * floating picture's anchor when the west or north edge is the one that moved.
 *
 * The gesture therefore carries two intents (srcRect + extent, plus a position
 * for a floating west/north crop) under ONE history checkpoint, so a single
 * undo puts both back.
 */
import { describe, expect, it } from "vitest";
import { createElement, act } from "react";
import { createRoot } from "react-dom/client";
import { DocxView, type DocxViewApi } from "../src/index.js";
import { blankDocxBytes } from "@wordinweb/server";
import { DocxDocument, localName, type XmlElement } from "@wordinweb/core";

const PNG = new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10, 1, 2, 3, 4]);
const glob = globalThis as unknown as Record<string, unknown>;
if (typeof URL.createObjectURL !== "function") {
  URL.createObjectURL = () => "blob:crop-frame";
  URL.revokeObjectURL = () => {};
}
async function tick(ms = 5) { await act(async () => { await new Promise<void>((r) => setTimeout(r, ms)); }); }

/** Depth-first search for the first element with this local name. */
function findEl(root: XmlElement, name: string): XmlElement | undefined {
  if (localName(root.name) === name) return root;
  for (const child of root.children) {
    const hit = findEl(child, name);
    if (hit) return hit;
  }
  return undefined;
}

function attrByLocal(el: XmlElement | undefined, name: string): string | undefined {
  if (!el) return undefined;
  const key = Object.keys(el.attrs).find((k) => localName(k) === name);
  return key ? el.attrs[key] : undefined;
}

/** The inline/anchor extent and the blipFill crop of the document's one picture. */
function pictureGeometry(bytes: Uint8Array) {
  const doc = DocxDocument.load(bytes);
  const drawing = findEl(doc.docRoot, "drawing")!;
  const extent = findEl(drawing, "extent")!;
  const srcRect = findEl(drawing, "srcRect");
  const anchor = findEl(drawing, "anchor");
  const num = (el: XmlElement | undefined, name: string) => {
    const raw = attrByLocal(el, name);
    return raw === undefined ? undefined : parseInt(raw, 10);
  };
  return {
    cx: num(extent, "cx")!,
    cy: num(extent, "cy")!,
    crop: { l: num(srcRect, "l"), t: num(srcRect, "t"), r: num(srcRect, "r"), b: num(srcRect, "b") },
    floating: !!anchor,
    posX: anchor ? readOffset(anchor, "positionH") : undefined,
  };
}

/** wp:positionH holds its offset in a child wp:posOffset element's text. */
function readOffset(anchor: XmlElement, which: string): number | undefined {
  const holder = findEl(anchor, which);
  const offset = holder && findEl(holder, "posOffset");
  return offset ? parseInt(offset.text, 10) : undefined;
}

/** Mount an editable blank document holding one image of the given natural size. */
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

  const api = () => seen.api!;
  /** Click the picture so the object toolbar and its handles appear. */
  const select = async () => {
    await act(async () => {
      const o = { bubbles: true, cancelable: true, clientX: 40, clientY: 40, button: 0 };
      img()!.dispatchEvent(new MouseEvent("mousedown", o));
      document.dispatchEvent(new MouseEvent("mouseup", { ...o, bubbles: true }));
    });
    await tick();
  };
  /** Enter crop mode and drag one crop handle by (dx, dy) unzoomed px. A
   * commit re-enters crop mode on its own, and the Crop command is a toggle,
   * so a second drag must NOT press it again. */
  const dragCrop = async (dir: string, dx: number, dy: number) => {
    if (!container.querySelector("[data-dxw-crop-handle]")) {
      await select();
      expect(api().runSelectedObjectCommand("crop"), "crop mode should open").toBe(true);
      await tick();
    }
    const handle = container.querySelector<HTMLElement>(`[data-dxw-crop-handle="${dir}"]`);
    expect(handle, `the ${dir} crop handle should exist`).toBeTruthy();
    await act(async () => {
      handle!.dispatchEvent(new MouseEvent("mousedown", { bubbles: true, cancelable: true, clientX: 200, clientY: 200, button: 0 }));
      document.dispatchEvent(new MouseEvent("mousemove", { bubbles: true, clientX: 200 + dx, clientY: 200 + dy }));
      document.dispatchEvent(new MouseEvent("mouseup", { bubbles: true, clientX: 200 + dx, clientY: 200 + dy }));
    });
    await tick(20);
  };
  // A cropped picture renders as an overflow:hidden viewport wrapping an
  // OVERSIZED <img>, so the frame's geometry is the viewport's; an uncropped
  // one renders as the <img> itself (render/dom.ts).
  const frame = () => {
    const el = img()!;
    const parent = el.parentElement;
    return parent && parent.style.overflow === "hidden" ? parent : el;
  };
  const boxWidth = () => parseFloat(frame().style.width);
  const boxLeft = () => parseFloat(frame().style.left) || 0;
  return {
    container, api, select, dragCrop, boxWidth, boxLeft,
    geometry: () => pictureGeometry(api().save()),
    unmount: async () => { await act(async () => { root.unmount(); }); },
  };
}

const EMU_PER_PX = 9525;

describe("a crop drag shrinks the picture's frame (Word's crop, not a zoom)", () => {
  it("cropping the WEST edge by a quarter writes srcRect l AND a quarter-narrower extent", async () => {
    const ed = await mountWithImage(800, 450);
    const before = ed.geometry();
    expect(before.crop.l, "an uncropped picture carries no srcRect").toBeUndefined();
    const w0 = ed.boxWidth();
    // An integer px drag: MouseEvent.clientX is an integer, so a fractional
    // one would be silently truncated and the expectations would drift.
    const dx = Math.round(w0 / 4);

    await ed.dragCrop("w", dx, 0);
    const after = ed.geometry();

    // The srcRect math is untouched: the removed fraction of the BITMAP, in
    // hundred-thousandths (CROP_UNIT).
    expect(after.crop.l).toBe(Math.round((dx / w0) * 100000));
    expect(after.crop.l! / 1000).toBeCloseTo(25, 0); // ~25% of the width
    // The frame is now the kept region's drawn size, rounded to whole EMU.
    expect(after.cx).toBe(Math.round((w0 - dx) * EMU_PER_PX));
    expect(after.cx / before.cx).toBeCloseTo(0.75, 2);
    // A west crop trims nothing off the height.
    expect(after.cy).toBe(before.cy);
    await ed.unmount();
  });

  it("cropping the EAST edge shrinks the frame too and leaves the anchor alone", async () => {
    const ed = await mountWithImage(800, 450);
    const before = ed.geometry();
    const w0 = ed.boxWidth();
    const dx = Math.round(w0 / 4);

    await ed.dragCrop("e", -dx, 0);
    const after = ed.geometry();

    expect(after.crop.r).toBe(Math.round((dx / w0) * 100000));
    expect(after.crop.l, "the west edge did not move").toBeUndefined();
    expect(after.cx).toBe(Math.round((w0 - dx) * EMU_PER_PX));
    await ed.unmount();
  });

  it("a FLOATING picture cropped from the west moves its anchor by the removed width", async () => {
    const ed = await mountWithImage(800, 450);
    await ed.select();
    expect(ed.api().runSelectedObjectCommand("wrapSquare"), "square wrap should float it").toBe(true);
    await tick(20);
    expect(ed.geometry().floating, "the picture should now be anchored").toBe(true);
    const w0 = ed.boxWidth();
    const left0 = ed.boxLeft();
    const dx = Math.round(w0 / 4);

    await ed.dragCrop("w", dx, 0);

    // Where the picture SITS is the claim; the anchor is only how it is
    // written. A west crop leaves the east edge on the page untouched and
    // walks the west edge in by what it removed.
    expect(ed.boxWidth()).toBeCloseTo(w0 - dx, 1);
    expect(ed.boxLeft()).toBeCloseTo(left0 + dx, 1);
    expect(ed.boxLeft() + ed.boxWidth(), "the east edge stays put").toBeCloseTo(left0 + w0, 1);
    // And it rode out as a position, not just an extent: the anchor now names
    // an absolute page offset matching where the picture landed.
    const after = ed.geometry();
    expect(after.cx).toBe(Math.round((w0 - dx) * EMU_PER_PX));
    expect(after.posX).toBe(Math.round((left0 + dx) * EMU_PER_PX));
    await ed.unmount();
  });

  it("a second crop composes: srcRect grows against the BITMAP, the frame against the box", async () => {
    // The two intents measure against different things, so a follow-up drag is
    // where they would come apart. srcRect l is a fraction of the whole bitmap
    // and accumulates; the frame is a fraction of what is left.
    const ed = await mountWithImage(800, 450);
    const w0 = ed.boxWidth();
    const dx1 = Math.round(w0 / 4);

    await ed.dragCrop("w", dx1, 0);
    const dx2 = Math.round(ed.boxWidth() / 4);
    await ed.dragCrop("w", dx2, 0);

    const after = ed.geometry();
    expect(after.crop.l).toBe(Math.round(((dx1 + dx2) / w0) * 100000));
    expect(after.cx).toBe(Math.round((w0 - dx1 - dx2) * EMU_PER_PX));
    expect(after.cx / (w0 * EMU_PER_PX)).toBeCloseTo(0.75 * 0.75, 2);
    await ed.unmount();
  });

  it("one undo restores the srcRect AND the extent together", async () => {
    const ed = await mountWithImage(800, 450);
    const before = ed.geometry();
    const dx = Math.round(ed.boxWidth() / 4);

    await ed.dragCrop("w", dx, 0);
    expect(ed.geometry().cx).not.toBe(before.cx);

    await act(async () => { ed.api().undo(); });
    await tick(20);
    const after = ed.geometry();
    expect(after.cx, "the frame is back").toBe(before.cx);
    expect(after.crop.l, "and so is the crop").toBeUndefined();
    await ed.unmount();
  });
});
