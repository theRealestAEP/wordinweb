// @vitest-environment jsdom
/**
 * Issue #18: a Calibri document opened on a machine without Office fonts
 * measured and painted in the browser's sans-serif fallback, ~10% wider, so
 * every line and page break drifted from Word with nothing on screen saying
 * why. The demo now ships the metric substitutes (main.tsx) and, for faces it
 * still cannot render, the local editor surfaces DocxView's `onMissingFonts`
 * report as a banner. The canvas fake below makes EVERY named face read as
 * missing — a face is missing when its stack measures exactly like the
 * generic it falls through to — so the banner must appear and name the face.
 */
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { createElement, act } from "react";
import { createRoot, type Root } from "react-dom/client";
import { zipSync, strToU8 } from "fflate";
import { InMemoryBundleStore } from "@wordinweb/collab/client";
import { LocalEditor } from "../src/local-editor";

function calibriDoc(text: string): Uint8Array {
  const xml = `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/></w:rPr><w:t xml:space="preserve">${text}</w:t></w:r></w:p></w:body></w:document>`;
  return zipSync({
    "[Content_Types].xml": strToU8(`<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`),
    "_rels/.rels": strToU8(`<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`),
    "word/document.xml": strToU8(xml),
  });
}

/** A canvas whose advances depend only on the GENERIC at the end of the
 * stack, which is exactly how a missing face measures in a real browser. */
function missingFaceContext() {
  return {
    font: "",
    measureText(text: string) {
      const per = /monospace/.test(this.font) ? 8 : /serif/.test(this.font) ? 6 : 5;
      return { width: per * text.length } as TextMetrics;
    },
    fillText() {}, clearRect() {}, setTransform() {}, scale() {}, translate() {},
    save() {}, restore() {}, beginPath() {}, fill() {}, drawImage() {},
  } as unknown as CanvasRenderingContext2D;
}

let mounted: { root: Root; host: HTMLElement }[] = [];
function render(node: ReturnType<typeof createElement>) {
  const host = document.createElement("div");
  document.body.appendChild(host);
  const root = createRoot(host);
  act(() => { root.render(node); });
  mounted.push({ root, host });
  return host;
}
afterEach(() => {
  for (const { root, host } of mounted) {
    act(() => { root.unmount(); });
    host.remove();
  }
  mounted = [];
});

const prevGetContext = HTMLCanvasElement.prototype.getContext;
beforeEach(() => {
  HTMLCanvasElement.prototype.getContext = (() => missingFaceContext()) as never;
  globalThis.fetch = vi.fn(async () => new Response(calibriDoc("blank") as unknown as BodyInit, { status: 200 })) as unknown as typeof fetch;
});
afterEach(() => { HTMLCanvasElement.prototype.getContext = prevGetContext; });

async function tick() { await act(async () => { await new Promise<void>((r) => setTimeout(r, 5)); }); }
const byId = (host: HTMLElement, id: string) => host.querySelector<HTMLElement>(`[data-testid="${id}"]`);

describe("missing-font banner (issue #18)", () => {
  it("names the face the browser substituted, and dismisses", async () => {
    const host = render(createElement(LocalEditor, {
      httpBase: "http://blankhost",
      onGoLive: async () => {},
      store: new InMemoryBundleStore(),
      autosaveMs: 30,
      initialBytes: calibriDoc("Set in Calibri"),
    }));
    for (let i = 0; i < 100 && !byId(host, "missing-fonts-banner"); i++) await tick();
    const banner = byId(host, "missing-fonts-banner");
    expect(banner, "the substituted face must be said on screen").toBeTruthy();
    expect(banner!.textContent).toContain("Calibri");

    await act(async () => {
      byId(host, "missing-fonts-dismiss")!.dispatchEvent(new MouseEvent("click", { bubbles: true }));
    });
    expect(byId(host, "missing-fonts-banner")).toBeNull();
  });
});
