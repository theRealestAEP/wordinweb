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
