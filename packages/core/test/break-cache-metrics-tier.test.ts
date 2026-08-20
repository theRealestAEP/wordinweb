/**
 * The break cache's metrics tier, and the paginate-only paragraph path it feeds.
 *
 * Under a page window the engine keeps positioned items for a dozen pages and
 * rebuilds the rest on demand, so the line breaks of every later paragraph are
 * used only to decide where pages end. Those paragraphs are broken into per-line
 * METRICS with no spans, which is what took the layout heap of a 483-page
 * document from 192 MB to 30 MB.
 *
 * The full-span tier is capped far below the document's paragraph count while
 * that is true, so these tests exist to catch a change that quietly stops the
 * metrics tier from answering: the window-determinism suite would stay green
 * (the output is identical either way) while every edit re-measured the tail.
 */
import { describe, expect, it } from "vitest";
import { DocxDocument } from "../src/docx.js";
import { layoutDocument } from "../src/layout/engine.js";
import { ApproxMeasurer } from "../src/layout/measure.js";
import type { FontSpec, LayoutResult } from "../src/layout/types.js";
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

/** One-line paragraphs — 1,600 of them is ~35 pages, well past the 12-page
 * window, so most of the document can only be paginated from metrics. */
function bigDoc(paras = 1600): DocxDocument {
  const body = Array.from({ length: paras }, (_, i) =>
    p(`para-${i} alpha bravo charlie delta echo foxtrot golf hotel india juliet`),
  ).join("");
  return DocxDocument.load(makeDocx({ "word/document.xml": wrapDocument(body) }));
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

function pageProjection(result: LayoutResult, pageIndex: number): string {
  return JSON.stringify(result.pages[pageIndex], (key, value) =>
    key === "src" || key === "tbl" ? undefined : value,
  );
}

describe("break cache metrics tier", () => {
  it("leaves the pages it never paints for the window to rebuild", () => {
    const measurer = new ApproxMeasurer();
    const win = layoutDocument(bigDoc(), { measurer, windowModel: true });
    expect(win.totalPages).toBeGreaterThan(20);
    expect(win._window).toBeTruthy();

    // Everything past the window is unpainted, so the retained set is the
    // window itself rather than the whole document.
    expect(win._window!.retainedPages().size).toBeLessThanOrEqual(20);
    expect(win.pages.slice(25).every((page) => page.items.length === 0)).toBe(true);
  });

  it("rebuilds a page it paginated from metrics byte-for-byte", () => {
    const measurer = new ApproxMeasurer();
    const doc = bigDoc();
    const full = layoutDocument(doc, { measurer });
    const win = layoutDocument(doc, { measurer, windowModel: true });
    expect(win.totalPages).toBe(full.totalPages);

    // A page laid entirely by the paginate-only path, well past the window.
    const index = win.totalPages - 3;
    win._window!.materialize([index]);
    expect(pageProjection(win, index)).toBe(pageProjection(full, index));
  });

  it("paginates an edit's cascade re-lay without re-measuring the tail", () => {
    // The tail this edit re-lays is far longer than the windowed full tier, so
    // without the metrics tier every entry would be evicted just before the
    // relay asked for it — the measured cliff this cache is built around.
    const doc = bigDoc();
    const measurer = new CountingMeasurer();
    const first = layoutDocument(doc, { measurer, windowModel: true });
    const mountCalls = measurer.calls;

    const inserted = splitParagraphAt(doc, 400);
    measurer.calls = 0;
    const second = layoutDocument(doc, { measurer, windowModel: true, prev: first, dirtyHint: inserted });

    expect(second._incremental).toBe(true);
    expect(measurer.calls).toBeLessThan(mountCalls / 10);
  });

  it("costs an edit the same on a document twice as long", () => {
    // The sharpest statement of what the metrics tier buys, and the one a
    // regression cannot fake: what an edit re-measures is the handful of pages
    // the relay actually paints, so it must not grow with the length of the
    // tail behind it. Re-measuring that tail is precisely the cliff.
    const editCost = (paras: number): number => {
      const doc = bigDoc(paras);
      const measurer = new CountingMeasurer();
      const first = layoutDocument(doc, { measurer, windowModel: true });
      const inserted = splitParagraphAt(doc, 400);
      measurer.calls = 0;
      layoutDocument(doc, { measurer, windowModel: true, prev: first, dirtyHint: inserted });
      return measurer.calls;
    };
    const short = editCost(1600);
    const long = editCost(3200);
    expect(long).toBeLessThan(short * 1.5);
  });
});
