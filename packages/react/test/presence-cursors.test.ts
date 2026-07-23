// @vitest-environment jsdom
import { describe, expect, it } from "vitest";
import { zipSync, strToU8 } from "fflate";
import { DocxDocument, localName, type Paragraph, type Run, type RenderHandle, type XmlElement } from "@wordinweb/core";
import { computePresenceCarets, presenceColor, drawPresenceCarets } from "../src/presence-cursors.js";

function makeDoc(text: string): DocxDocument {
  const xml = `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p></w:body></w:document>`;
  return DocxDocument.load(zipSync({
    "[Content_Types].xml": strToU8(`<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`),
    "_rels/.rels": strToU8(`<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`),
    "word/document.xml": strToU8(xml),
  }));
}

/** Build a render handle whose bindingsByText maps the run's w:t to a text
 * item with known geometry (x=100, width=80, lineTop=50, lineHeight=16). */
function handleFor(doc: DocxDocument, t: XmlElement, text: string): RenderHandle {
  const el = document.createElement("span");
  const item = { kind: "text", x: 100, baseline: 62, width: 80, text, props: {}, font: {}, lineTop: 50, lineHeight: 16 } as never;
  const bindingsByText = new Map<XmlElement, { el: HTMLElement; item: never }[]>();
  bindingsByText.set(t, [{ el, item }]);
  return { root: document.createElement("div"), bindings: [], bindingsByText, grips: [], images: [], drawings: [], wordarts: [], destroy: () => {} } as unknown as RenderHandle;
}

function runInfo(doc: DocxDocument) {
  const para = doc.sections[0].blocks[0] as Paragraph;
  const run = para.children[0] as Run;
  const t = run.content.find((c) => c.kind === "text")!.srcT!;
  return { blockId: doc.stableIds!.idOf(para.src!)!, runId: doc.stableIds!.idOf(run.src!)!, t };
}

describe("computePresenceCarets", () => {
  it("maps a stable-id position to page-pixel caret geometry", () => {
    const doc = makeDoc("HELLOWORLD"); // 10 chars
    doc.enableStableIds();
    const { blockId, runId, t } = runInfo(doc);
    const handle = handleFor(doc, t, "HELLOWORLD");
    const carets = computePresenceCarets(handle, doc, { alice: { anchor: { blockId, runId, offset: 5 } } });
    expect(carets).toHaveLength(1);
    const c = carets[0];
    // offset 5 of 10 -> halfway across the 80px run: x = 100 + 0.5*80 = 140.
    expect(c.x).toBe(140);
    expect(c.top).toBe(50);
    expect(c.height).toBe(16);
    expect(c.participant).toBe("alice");
    expect(c.color).toMatch(/^#/);
  });

  it("places the caret at the run start for offset 0 and end for full offset", () => {
    const doc = makeDoc("ABCD"); doc.enableStableIds();
    const { blockId, runId, t } = runInfo(doc);
    const handle = handleFor(doc, t, "ABCD");
    expect(computePresenceCarets(handle, doc, { a: { anchor: { blockId, runId, offset: 0 } } })[0].x).toBe(100);
    expect(computePresenceCarets(handle, doc, { a: { anchor: { blockId, runId, offset: 4 } } })[0].x).toBe(180); // 100+80
  });

  it("skips a position whose run isn't rendered, and null presence", () => {
    const doc = makeDoc("x"); doc.enableStableIds();
    const { blockId, t } = runInfo(doc);
    const handle = handleFor(doc, t, "x");
    // Unknown runId 9999 -> no caret; null presence -> no caret.
    expect(computePresenceCarets(handle, doc, { a: { anchor: { blockId, runId: 9999, offset: 0 } } })).toHaveLength(0);
    expect(computePresenceCarets(handle, doc, { a: null })).toHaveLength(0);
  });

  it("presenceColor is stable per participant", () => {
    expect(presenceColor("alice")).toBe(presenceColor("alice"));
  });

  it("drawPresenceCarets renders positioned caret bars into an overlay", () => {
    const doc = makeDoc("HELLOWORLD"); doc.enableStableIds();
    const { blockId, runId, t } = runInfo(doc);
    const handle = handleFor(doc, t, "HELLOWORLD");
    const carets = computePresenceCarets(handle, doc, { alice: { anchor: { blockId, runId, offset: 3 } } });
    const overlay = document.createElement("div");
    drawPresenceCarets(overlay, carets);
    const bar = overlay.querySelector<HTMLElement>(".dxw-presence-caret")!;
    expect(bar).toBeTruthy();
    expect(bar.dataset.participant).toBe("alice");
    expect(bar.style.left).toBe("124px"); // 100 + 0.3*80
    expect(bar.style.height).toBe("16px");
  });
});
