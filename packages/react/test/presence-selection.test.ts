// @vitest-environment jsdom
/**
 * REMOTE SELECTION HIGHLIGHT (receiver side): a participant's selection rides
 * the presence payload as `ranges` — wire-basis `[start, end)` inside one run,
 * the same addressing the caret's `anchor.offset` uses (cumulative within the
 * run, inline separators counting one unit each) — and computePresenceSelections
 * turns each range into the rects to paint in that participant's color.
 *
 * These pin the three things that can silently go wrong: the geometry within a
 * w:t, landing on the RIGHT w:t of a multi-w:t run (the separator-inclusive
 * basis, same trap as presence-offset-basis), and that a hostile/garbage
 * payload draws nothing instead of throwing or painting the whole page.
 *
 * jsdom can't measure glyphs (every Range rect is 0×0), so the numbers below
 * are the LINEAR fallback — offset fractions of the layout item's width. In a
 * browser the DOM-Range measurement takes over, exactly as for the caret.
 */
import { describe, expect, it } from "vitest";
import { zipSync, strToU8 } from "fflate";
import { DocxDocument, type Paragraph, type Run, type RenderHandle, type XmlElement } from "@wordinweb/core";
import { computePresenceSelections, drawPresenceSelections, presenceColor } from "../src/presence-cursors.js";

function loadDoc(bodyInner: string): DocxDocument {
  const xml = `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${bodyInner}</w:body></w:document>`;
  return DocxDocument.load(zipSync({
    "[Content_Types].xml": strToU8(`<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`),
    "_rels/.rels": strToU8(`<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`),
    "word/document.xml": strToU8(xml),
  }));
}

function textsOf(runEl: XmlElement): XmlElement[] {
  const out: XmlElement[] = [];
  const walk = (el: XmlElement) => {
    if (el.name.endsWith(":t")) out.push(el);
    for (const c of el.children) walk(c);
  };
  walk(runEl);
  return out;
}

const W = 80;

/** One layout item per w:t, laid out left-to-right (same shape as the
 * offset-basis harness): w:t i occupies [100 + i*W, 100 + (i+1)*W). */
function handleForRun(texts: { t: XmlElement; text: string }[]): RenderHandle {
  const root = document.createElement("div");
  const surface = document.createElement("div");
  root.appendChild(surface);
  const bindingsByText = new Map<XmlElement, { el: HTMLElement; item: never }[]>();
  texts.forEach(({ t, text }, i) => {
    const el = document.createElement("span");
    el.textContent = text;
    surface.appendChild(el);
    const item = { kind: "text", x: 100 + i * W, baseline: 62, width: W, text, props: {}, font: {}, lineTop: 50, lineHeight: 16, src: { offset: 0 } } as never;
    bindingsByText.set(t, [{ el, item }]);
  });
  return { root, bindings: [], bindingsByText, grips: [], images: [], drawings: [], wordarts: [], destroy: () => {} } as unknown as RenderHandle;
}

/** One w:t split across TWO wrap segments (word-wrap / justification), the
 * layout shape a long run actually produces: segment 0 covers chars [0,5),
 * segment 1 covers [5,10) on the next line. */
function handleForWrappedText(t: XmlElement, text: string, splitAt: number): RenderHandle {
  const root = document.createElement("div");
  const surface = document.createElement("div");
  root.appendChild(surface);
  const parts = [
    { offset: 0, text: text.slice(0, splitAt), lineTop: 50 },
    { offset: splitAt, text: text.slice(splitAt), lineTop: 70 },
  ];
  const bindings = parts.map((p) => {
    const el = document.createElement("span");
    el.textContent = p.text;
    surface.appendChild(el);
    const item = { kind: "text", x: 100, baseline: p.lineTop + 12, width: W, text: p.text, props: {}, font: {}, lineTop: p.lineTop, lineHeight: 16, src: { offset: p.offset } } as never;
    return { el, item };
  });
  const bindingsByText = new Map<XmlElement, { el: HTMLElement; item: never }[]>([[t, bindings]]);
  return { root, bindings: [], bindingsByText, grips: [], images: [], drawings: [], wordarts: [], destroy: () => {} } as unknown as RenderHandle;
}

describe("remote selection highlight (computePresenceSelections)", () => {
  it("SINGLE w:t: a range renders one rect at the selection's x/width", () => {
    const doc = loadDoc(`<w:p><w:r><w:t xml:space="preserve">HELLOWORLD</w:t></w:r></w:p>`);
    doc.enableStableIds();
    const para = doc.sections[0].blocks[0] as Paragraph;
    const runEl = (para.children[0] as Run).src!;
    const t = textsOf(runEl)[0];
    const runId = doc.stableIds!.idOf(runEl)!;
    const handle = handleForRun([{ t, text: "HELLOWORLD" }]);

    // "HE[LLOW]ORLD": wire offsets 2..7 of a 10-char single-w:t run.
    const rects = computePresenceSelections(handle, doc, {
      alice: { anchor: { blockId: 1, runId, offset: 7 }, ranges: [{ blockId: 1, runId, start: 2, end: 7 }] },
    });
    expect(rects).toHaveLength(1);
    expect(rects[0].x).toBe(100 + (2 / 10) * W); // 116
    expect(rects[0].width).toBe((5 / 10) * W); // 40
    expect(rects[0].top).toBe(50);
    expect(rects[0].height).toBe(16);
    expect(rects[0].color).toBe(presenceColor("alice"));

    // Drawn into the page SURFACE (the anchor's parent), inert, translucent.
    drawPresenceSelections(handle.root, rects);
    const boxes = handle.root.querySelectorAll<HTMLElement>(".dxw-presence-selection");
    expect(boxes).toHaveLength(1);
    expect(boxes[0].parentElement).toBe(rects[0].anchorEl.parentElement);
    expect(boxes[0].style.left).toBe("116px");
    expect(boxes[0].style.width).toBe("40px");
    expect(boxes[0].style.pointerEvents).toBe("none");
    expect(boxes[0].style.background).toMatch(/^rgba\(/);
    // Redrawing clears the previous pass (no accumulation across repaints).
    drawPresenceSelections(handle.root, rects);
    expect(handle.root.querySelectorAll(".dxw-presence-selection")).toHaveLength(1);
  });

  it("MULTI w:t run: a selection in the SECOND w:t lands on the second w:t", () => {
    // One run, two w:t split by an inline tab — the wire basis counts the tab
    // as one unit, so "WO" in the second w:t starts at 5 + 1 = 6, not 5.
    const doc = loadDoc(`<w:p><w:r><w:t xml:space="preserve">HELLO</w:t><w:tab/><w:t xml:space="preserve">WORLD</w:t></w:r></w:p>`);
    doc.enableStableIds();
    const para = doc.sections[0].blocks[0] as Paragraph;
    const runEl = (para.children[0] as Run).src!;
    const [firstT, secondT] = textsOf(runEl);
    const runId = doc.stableIds!.idOf(runEl)!;
    const handle = handleForRun([{ t: firstT, text: "HELLO" }, { t: secondT, text: "WORLD" }]);

    // Sender-equivalent encoding of "WOR|LD" start: encodeCaret(secondT, 0).
    const enc = doc.stableIds!.encodeCaret(secondT, 0, (el) => doc.findParentOf(el) ?? null)!;
    expect(enc.offset, "wire start counts HELLO(5) + tab(1)").toBe(6);
    const rects = computePresenceSelections(handle, doc, {
      bob: { anchor: { blockId: enc.blockId, runId, offset: 9 }, ranges: [{ blockId: enc.blockId, runId, start: 6, end: 9 }] },
    });
    expect(rects).toHaveLength(1);
    // Second w:t's item starts at x=180; "WOR" is its first 3 of 5 chars.
    expect(rects[0].x, "highlight must sit on the SECOND w:t").toBe(100 + W + 0);
    expect(rects[0].width).toBe((3 / 5) * W); // 48
    expect(rects[0].anchorEl.textContent).toBe("WORLD");
  });

  it("WRAPPED w:t: a selection crossing wrap segments emits one rect per line", () => {
    const doc = loadDoc(`<w:p><w:r><w:t xml:space="preserve">HELLOWORLD</w:t></w:r></w:p>`);
    doc.enableStableIds();
    const para = doc.sections[0].blocks[0] as Paragraph;
    const runEl = (para.children[0] as Run).src!;
    const t = textsOf(runEl)[0];
    const runId = doc.stableIds!.idOf(runEl)!;
    const handle = handleForWrappedText(t, "HELLOWORLD", 5);

    // Select chars 3..8 — the last 2 of line 1 and the first 3 of line 2.
    const rects = computePresenceSelections(handle, doc, {
      cara: { anchor: { blockId: 1, runId, offset: 8 }, ranges: [{ blockId: 1, runId, start: 3, end: 8 }] },
    });
    expect(rects).toHaveLength(2);
    expect(rects[0].top).toBe(50);
    expect(rects[0].x).toBe(100 + (3 / 5) * W); // from char 3 of segment [0,5)
    expect(rects[0].width).toBe((2 / 5) * W);
    expect(rects[1].top).toBe(70);
    expect(rects[1].x).toBe(100); // segment [5,10) starts at its own x
    expect(rects[1].width).toBe((3 / 5) * W);
  });

  it("MALFORMED ranges render nothing and never throw", () => {
    const doc = loadDoc(`<w:p><w:r><w:t xml:space="preserve">HELLOWORLD</w:t></w:r></w:p>`);
    doc.enableStableIds();
    const para = doc.sections[0].blocks[0] as Paragraph;
    const runEl = (para.children[0] as Run).src!;
    const t = textsOf(runEl)[0];
    const runId = doc.stableIds!.idOf(runEl)!;
    const handle = handleForRun([{ t, text: "HELLOWORLD" }]);
    const anchor = { blockId: 1, runId, offset: 0 };

    const junk = [
      { blockId: 1, runId, start: 5, end: 5 }, // empty
      { blockId: 1, runId, start: 8, end: 3 }, // inverted
      { blockId: 1, runId, start: -4, end: 3 }, // negative
      { blockId: 1, runId, start: NaN, end: 3 },
      { blockId: 1, runId, start: 0, end: Infinity },
      { blockId: 1, runId, start: "0", end: "9" }, // wrong types
      { blockId: 1, runId: 9999, start: 0, end: 3 }, // unknown run
      null,
      undefined,
      "nope",
    ] as never;
    expect(() => computePresenceSelections(handle, doc, { evil: { anchor, ranges: junk } })).not.toThrow();
    expect(computePresenceSelections(handle, doc, { evil: { anchor, ranges: junk } })).toEqual([]);

    // Non-array `ranges`, and a payload with no ranges at all (an OLD client:
    // caret-only presence must still be handled, drawing no highlight).
    expect(computePresenceSelections(handle, doc, { evil: { anchor, ranges: "all of it" as never } })).toEqual([]);
    expect(computePresenceSelections(handle, doc, { old: { anchor } })).toEqual([]);
    expect(computePresenceSelections(handle, doc, { gone: null })).toEqual([]);

    // Flooding: 500 valid ranges are capped at 64 rects, so one participant
    // can't make every other tab paint an unbounded number of nodes.
    const flood = Array.from({ length: 500 }, () => ({ blockId: 1, runId, start: 0, end: 4 }));
    expect(computePresenceSelections(handle, doc, { flood: { anchor, ranges: flood } })).toHaveLength(64);

    // An out-of-range end is clamped into the w:t rather than painting past it.
    const over = computePresenceSelections(handle, doc, {
      over: { anchor, ranges: [{ blockId: 1, runId, start: 8, end: 1000 }] },
    });
    expect(over).toHaveLength(1);
    expect(over[0].x + over[0].width).toBeLessThanOrEqual(100 + W);
  });
});
