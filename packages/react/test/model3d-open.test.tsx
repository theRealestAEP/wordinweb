// @vitest-environment jsdom
/**
 * THE DEMO'S File > Open, for a document that contains a 3D model.
 *
 * Reported as "file open just doesn't work, at least for 3d objects". It
 * works — and this pins the two halves of it that a headless parse test
 * cannot see, because the demo's Open is not a mount, it is a PROP CHANGE:
 * `local-editor.tsx` calls `setBlank(bytes)` on the same mounted DocxView and
 * relies on the load effect re-running because `source` is in its deps and
 * the document cache is keyed on the source's IDENTITY. If that cache ever
 * stops distinguishing two different Uint8Arrays, Open silently keeps showing
 * the document you had before — which is exactly what "open doesn't work"
 * looks like from the user's seat, with no error anywhere.
 *
 * The second half is that the reopened document actually paints its model.
 */
import { describe, expect, it } from "vitest";
import { createElement, act } from "react";
import { createRoot } from "react-dom/client";
import { DocxView } from "../src/index.js";
import { blankDocxBytes } from "@wordinweb/server";
import { DocxDocument, insertModel3DAt, localName, type XmlElement } from "@wordinweb/core";

const glob = globalThis as unknown as Record<string, unknown>;
glob.createImageBitmap ??= async () => ({ width: 64, height: 48, close() {} });
if (typeof URL.createObjectURL !== "function") {
  URL.createObjectURL = () => "blob:model3d-open";
  URL.revokeObjectURL = () => {};
}
// A 3D model makes DocxView lazy-import @google/model-viewer, whose WebGL
// renderer throws inside jsdom. Claiming the tag name keeps the import from
// happening (DocxView checks customElements first). renderToDom still creates
// the <model-viewer> element itself, which is what these assertions read.
if (typeof customElements !== "undefined" && !customElements.get("model-viewer")) {
  customElements.define("model-viewer", class extends HTMLElement {});
}

const PNG = new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]);
const GLB = new Uint8Array([0x67, 0x6c, 0x54, 0x46, 2, 0, 0, 0, 12, 0, 0, 0]);

async function tick(ms = 5) {
  await act(async () => { await new Promise<void>((r) => setTimeout(r, ms)); });
}

/** Blank template with a 3D model dropped into its first run. */
function model3dDocxBytes(): Uint8Array {
  const doc = DocxDocument.load(blankDocxBytes());
  let anchor: XmlElement | null = null;
  const visit = (el: XmlElement): void => {
    if (anchor) return;
    if (localName(el.name) === "t") { anchor = el; return; }
    for (const child of el.children) visit(child);
  };
  visit(doc.docRoot);
  if (!anchor) throw new Error("blank template has no w:t to anchor on");
  if (!insertModel3DAt(doc, anchor, { data: GLB, poster: PNG })) throw new Error("insertModel3DAt refused");
  return doc.save();
}

describe("demo File > Open with a 3D model in the document", () => {
  it("re-sources the mounted view and paints the model", async () => {
    const container = document.createElement("div");
    document.body.appendChild(container);
    const root = createRoot(container);
    const render = async (source: Uint8Array) => {
      await act(async () => {
        root.render(createElement(DocxView, { source, editable: true }));
      });
    };

    // The demo lands on the blank template.
    await render(blankDocxBytes());
    for (let i = 0; i < 30 && !container.querySelector(".dxw-page"); i++) await tick();
    expect(container.querySelector(".dxw-page")).toBeTruthy();
    expect(container.querySelector("[data-dxw-model3d]")).toBeNull();

    // File > Open hands NEW bytes to the SAME view. No remount, no key.
    await render(model3dDocxBytes());
    for (let i = 0; i < 60 && !container.querySelector("[data-dxw-model3d]"); i++) await tick();

    // The reopened document is on screen…
    const model = container.querySelector<HTMLElement>("[data-dxw-model3d]");
    expect(model).toBeTruthy();
    // …and it is a live viewer over the .glb part, not just the poster.
    const viewer = container.querySelector<HTMLElement>("[data-dxw-model3d-viewer]");
    expect(viewer).toBeTruthy();
    expect(viewer!.getAttribute("src")).toMatch(/^blob:/);

    await act(async () => { root.unmount(); });
    container.remove();
  });
});

/**
 * ONE REPRESENTATION, AND THE OBJECT STILL MOVES (#140).
 *
 * Reported as the model showing "some other same copy artifact behind it
 * doubling it up". The interactive branch built a frame holding BOTH the
 * poster <img> and the <model-viewer>, and the viewer paints on a transparent
 * background — so every part of the box the model did not cover showed the
 * poster underneath, and rotating the model separated the two completely.
 *
 * Chasing that turned up a second defect nobody had reported in those words:
 * the viewer covered the whole box (inset:0) and its pointerdown called
 * stopPropagation, so EVERY drag rotated and the editor never saw a move
 * gesture. A 3D object could not be dragged anywhere; it just spun.
 */
describe("a 3D model paints once and stays draggable", () => {
  const mount = async (editable: boolean) => {
    const container = document.createElement("div");
    document.body.appendChild(container);
    const root = createRoot(container);
    await act(async () => {
      root.render(createElement(DocxView, { source: model3dDocxBytes(), editable }));
    });
    for (let i = 0; i < 60 && !container.querySelector("[data-dxw-model3d]"); i++) await tick();
    const model = container.querySelector<HTMLElement>("[data-dxw-model3d]");
    expect(model).toBeTruthy();
    return { container, root, model: model! };
  };

  it("shows the viewer INSTEAD of the poster, not on top of it", async () => {
    const { container, root, model } = await mount(true);

    // The doubling, stated as the two things that cannot both be in the box.
    expect(model.querySelectorAll("img")).toHaveLength(0);
    expect(model.querySelectorAll("[data-dxw-model3d-viewer]")).toHaveLength(1);
    // The poster is not lost — model-viewer paints it itself until the GLB
    // loads, which is the whole reason the second copy was redundant.
    const viewer = model.querySelector<HTMLElement>("[data-dxw-model3d-viewer]")!;
    expect(viewer.getAttribute("poster")).toMatch(/^blob:|^data:/);

    await act(async () => { root.unmount(); });
    container.remove();
  });

  it("leaves the viewer transparent to pointers, with rotation on its own handle", async () => {
    const { container, root, model } = await mount(true);
    const viewer = model.querySelector<HTMLElement>("[data-dxw-model3d-viewer]")!;
    const handle = model.querySelector<HTMLElement>("[data-dxw-model3d-rotate]");

    // The viewer must not eat the gesture the editor needs to move the object.
    expect(viewer.style.pointerEvents).toBe("none");
    // Rotation is still reachable, on a control of its own.
    expect(handle).toBeTruthy();
    expect(handle!.style.cursor).toBe("grab");

    await act(async () => { root.unmount(); });
    container.remove();
  });

  it("keeps the read-only render exactly as it was — the corpus measures that one", async () => {
    // A parity capture loads with editable=0, so the whole viewer branch is
    // skipped and the poster <img> IS the render. coverletter-anon scores
    // 0.00% on that path, so it must not move.
    const { container, root, model } = await mount(false);
    expect(model.querySelectorAll("[data-dxw-model3d-viewer]")).toHaveLength(0);
    expect(model.querySelectorAll("[data-dxw-model3d-rotate]")).toHaveLength(0);
    expect(model.tagName === "IMG" || model.querySelector("img")).toBeTruthy();

    await act(async () => { root.unmount(); });
    container.remove();
  });
});
