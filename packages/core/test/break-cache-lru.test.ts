/**
 * The break cache evicts least-recently-used entries instead of clearing.
 *
 * An Enter mid-document re-lays every block after it (see
 * incremental-convergence.test.ts: the reflow is genuine), and each of those
 * blocks gets its line breaks from this cache. Clearing the whole cache on
 * overflow dropped entries the relay was about to ask for, so one keystroke
 * re-measured a large part of the document. LRU keeps the working set the
 * relay is walking.
 */
import { describe, expect, it, afterEach } from "vitest";
import { DocxDocument } from "../src/docx.js";
import { layoutDocument } from "../src/layout/engine.js";
import { ApproxMeasurer } from "../src/layout/measure.js";
import { __setBreakCacheMax } from "../src/layout/inline.js";
import type { FontSpec } from "../src/layout/types.js";
import type { XmlElement } from "../src/xml.js";
import { makeDocx, p, wrapDocument } from "./helpers.js";

/** Counts width() calls — the causal unit of layout work. */
class CountingMeasurer extends ApproxMeasurer {
  calls = 0;
  width(text: string, font: FontSpec, letterSpacing = 0): number {
    this.calls++;
    return super.width(text, font, letterSpacing);
  }
}

function splitParagraphAt(doc: DocxDocument, idx: number): XmlElement {
  const body = doc.docRoot.children.find((child) => child.name.endsWith("body"))!;
  const paragraphs = body.children.filter((child) => child.name.endsWith(":p"));
  const before = paragraphs[idx];
  const firstText = (el: XmlElement): XmlElement | null => {
    if (el.name.endsWith(":t")) return el;
    for (const child of el.children) {
      const hit = firstText(child);
      if (hit) return hit;
    }
    return null;
  };
  const after = JSON.parse(JSON.stringify(before)) as XmlElement;
  const head = firstText(before)!;
  const tail = firstText(after)!;
  const full = head.text;
  head.text = full.slice(0, 10);
  tail.text = full.slice(10);
  body.children.splice(body.children.indexOf(before) + 1, 0, after);
  doc.reparseDirectBodyParagraphSplit(before, after);
  return after;
}

afterEach(() => {
  __setBreakCacheMax(60000);
});

describe("break cache eviction", () => {
  it("keeps the relay's working set when the cache overflows", () => {
    // 1200 paragraphs, cache capped at 1000. The mount overflows the cap, so
    // eviction runs. The edit at 900 makes the relay walk blocks 900..1200 —
    // ~300 entries, comfortably inside the cap, so LRU must still serve them.
    __setBreakCacheMax(1000);
    const body = Array.from({ length: 1200 }, (_, i) =>
      p(`para-${i} alpha bravo charlie delta echo foxtrot golf hotel india juliet`),
    ).join("");
    const doc = DocxDocument.load(makeDocx({ "word/document.xml": wrapDocument(body) }));
    const measurer = new CountingMeasurer();
    const first = layoutDocument(doc, { measurer, windowModel: true });
    const mountCalls = measurer.calls;

    const inserted = splitParagraphAt(doc, 900);
    measurer.calls = 0;
    layoutDocument(doc, { measurer, windowModel: true, prev: first, dirtyHint: inserted });

    // Only the edited paragraph and its neighbours need measuring; the rest of
    // the relay is served from cache. Measured: 24 calls with LRU eviction
    // against 2462 when overflow cleared the cache instead.
    expect(measurer.calls).toBeLessThan(mountCalls / 100);
  });
});
