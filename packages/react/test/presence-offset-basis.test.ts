// @vitest-environment jsdom
/**
 * OFFSET-BASIS pin (checkpoint B1): the presence sender (editor.encodeCaret)
 * and receiver (computePresenceCarets) must agree on the wire offset basis —
 * CUMULATIVE within the run, with inline separators (tab/br/cr) counted as
 * one unit each so the mapping is one-to-one at w:t boundaries.
 *
 * Sender: `StableIds.encodeCaret(caret.t, caret.offset, parentOf)` encodes the
 * offset as the run's wire prefix before the caret's w:t (text lengths plus
 * separator units) plus the offset inside the caret's own w:t.
 *
 * Receiver: `computePresenceCarets` resolves runId -> run, then walks the
 * run's w:t ACCUMULATING (resolveRunOffset) to find the concrete w:t the
 * offset falls in, and interprets the local remainder against that w:t's
 * layout bindings (item.src.offset is "char offset within t").
 *
 * HISTORY: the wire basis used to be "offset within the sender's OWN w:t"
 * while every decoder resolved against the run's FIRST w:t — the bases agreed
 * only for single-w:t runs, so with an inline tab/break a caret in a later
 * w:t drew at the line start. The MULTI test below was the intentionally-
 * failing repro; it now pins the cumulative migration.
 */
import { describe, expect, it } from "vitest";
import { zipSync, strToU8 } from "fflate";
import { DocxDocument, type Paragraph, type Run, type RenderHandle, type XmlElement } from "@wordinweb/core";
import { computePresenceCarets } from "../src/presence-cursors.js";

function loadDoc(bodyInner: string): DocxDocument {
  const xml = `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${bodyInner}</w:body></w:document>`;
  return DocxDocument.load(zipSync({
    "[Content_Types].xml": strToU8(`<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`),
    "_rels/.rels": strToU8(`<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`),
    "word/document.xml": strToU8(xml),
  }));
}

/** All w:t elements under a run, in document order. */
function textsOf(runEl: XmlElement): XmlElement[] {
  const out: XmlElement[] = [];
  const walk = (el: XmlElement) => {
    if (el.name.endsWith(":t")) out.push(el);
    for (const c of el.children) walk(c);
  };
  walk(runEl);
  return out;
}

/**
 * Build a render handle that binds EACH w:t to a text item, laid out
 * left-to-right in document order. Each w:t occupies [x, x+width); a caret at
 * fraction f of w:t i sits at pxOf(i, f). src.offset = 0 (each item covers its
 * whole w:t, offset within that w:t).
 */
function handleForRun(texts: { t: XmlElement; text: string }[]): { handle: RenderHandle; pxOf: (i: number, frac: number) => number } {
  const root = document.createElement("div");
  const surface = document.createElement("div");
  root.appendChild(surface);
  const W = 80;
  const bindingsByText = new Map<XmlElement, { el: HTMLElement; item: never }[]>();
  texts.forEach(({ t, text }, i) => {
    const el = document.createElement("span");
    surface.appendChild(el);
    const item = { kind: "text", x: 100 + i * W, baseline: 62, width: W, text, props: {}, font: {}, lineTop: 50, lineHeight: 16, src: { offset: 0 } } as never;
    bindingsByText.set(t, [{ el, item }]);
  });
  const handle = { root, bindings: [], bindingsByText, grips: [], images: [], drawings: [], wordarts: [], destroy: () => {} } as unknown as RenderHandle;
  return { handle, pxOf: (i, frac) => 100 + i * W + frac * W };
}

describe("presence offset basis (sender encodeCaret vs receiver computePresenceCarets)", () => {
  it("SINGLE w:t run: bases agree — caret lands where the sender is", () => {
    const doc = loadDoc(`<w:p><w:r><w:t xml:space="preserve">HELLOWORLD</w:t></w:r></w:p>`);
    doc.enableStableIds();
    const para = doc.sections[0].blocks[0] as Paragraph;
    const runEl = (para.children[0] as Run).src!;
    const t = textsOf(runEl)[0];
    const { handle, pxOf } = handleForRun([{ t, text: "HELLOWORLD" }]);

    // Sender caret: 5 chars into the (only) w:t.
    const enc = doc.stableIds!.encodeCaret(t, 5, (el) => doc.findParentOf(el) ?? null)!;
    expect(enc).not.toBeNull();
    const carets = computePresenceCarets(handle, doc, { alice: { anchor: enc } });
    expect(carets).toHaveLength(1);
    // Caret should sit at 5/10 across the single w:t.
    expect(carets[0].x).toBe(pxOf(0, 0.5));
  });

  it("MULTI w:t run: caret in the SECOND w:t is decoded against the FIRST — WRONG SPOT", () => {
    // One run, two w:t split by an inline tab. Both w:t belong to run id N;
    // encodeCaret(secondT, off) sends offset within the SECOND w:t.
    const doc = loadDoc(`<w:p><w:r><w:t xml:space="preserve">HELLO</w:t><w:tab/><w:t xml:space="preserve">WORLD</w:t></w:r></w:p>`);
    doc.enableStableIds();
    const para = doc.sections[0].blocks[0] as Paragraph;
    const runEl = (para.children[0] as Run).src!;
    const texts = textsOf(runEl);
    expect(texts.length, "run must have two w:t").toBe(2);
    const firstT = texts[0];
    const secondT = texts[1];
    const { handle, pxOf } = handleForRun([{ t: firstT, text: "HELLO" }, { t: secondT, text: "WORLD" }]);

    // Sender caret: 2 chars into the SECOND w:t ("WO|RLD") -> visually item 1, frac 2/5.
    const enc = doc.stableIds!.encodeCaret(secondT, 2, (el) => doc.findParentOf(el) ?? null)!;
    expect(enc).not.toBeNull();
    // Wire basis: cumulative with SEPARATORS counted (one unit per inline
    // tab/br/cr, making encode one-to-one at w:t boundaries): "HELLO" (5)
    // + tab (1) + 2 into "WORLD" = wire offset 8.
    expect(enc.offset).toBe(8);
    const carets = computePresenceCarets(handle, doc, { alice: { anchor: enc } });
    expect(carets).toHaveLength(1);

    const expectedX = pxOf(1, 2 / 5); // second w:t, 2/5 across — where the sender is
    const actualX = carets[0].x;
    // Regression pin for the B1 migration: the receiver must walk the run's
    // w:t accumulating and land the caret on the SECOND w:t — under the old
    // first-w:t basis it drew on "HELLO" (the wrong w:t, at the line start).
    expect(actualX, "remote caret must land on the SECOND w:t where the sender is").toBe(expectedX);
  });

  it("GEOMETRY: presence x is a font-BLIND linear estimate, unlike the measured local caret", () => {
    // Even when run+offset resolve correctly, computePresenceCarets computes x as
    // `item.x + (offset/len) * item.width` — a UNIFORM per-character split. The
    // LOCAL caret (editor.positionCaret) instead measures the real glyph box via
    // a DOM Range (`range.getBoundingClientRect().left`), so it honors
    // proportional advances/kerning. In a real browser these DIVERGE; the remote
    // caret sits at the wrong x within the run. jsdom can't measure glyphs, so we
    // pin the MECHANISM: presence x depends only on offset/len/width and ignores
    // where the glyphs actually are.
    const doc = loadDoc(`<w:p><w:r><w:t xml:space="preserve">iWiWiWiWiW</w:t></w:r></w:p>`); // 10 chars, wildly non-uniform advances in any real font
    doc.enableStableIds();
    const para = doc.sections[0].blocks[0] as Paragraph;
    const runEl = (para.children[0] as Run).src!;
    const t = textsOf(runEl)[0];
    const { handle } = handleForRun([{ t, text: "iWiWiWiWiW" }]);

    // Caret after the first "i" (offset 1). A real font places this a few px in
    // (an "i" is narrow); the linear estimate places it at exactly 1/10 of width.
    const enc = doc.stableIds!.encodeCaret(t, 1, (el) => doc.findParentOf(el) ?? null)!;
    const x1 = computePresenceCarets(handle, doc, { a: { anchor: enc } })[0].x;
    // Caret after the first "iW" (offset 2): linear => exactly 2/10 of width.
    const enc2 = doc.stableIds!.encodeCaret(t, 2, (el) => doc.findParentOf(el) ?? null)!;
    const x2 = computePresenceCarets(handle, doc, { a: { anchor: enc2 } })[0].x;
    const W = 80;
    // The estimate is exactly linear in offset — proof it ignores glyph metrics.
    expect(x1).toBe(100 + (1 / 10) * W); // 108
    expect(x2).toBe(100 + (2 / 10) * W); // 116
    // A measured caret would NOT be evenly spaced here ("i" << "W"): x2 - x1 for
    // a proportional font is the width of "W", far more than W/10. The presence
    // caret being evenly spaced is the residual horizontal error.
    expect(x2 - x1).toBe(W / 10);
  });
});
