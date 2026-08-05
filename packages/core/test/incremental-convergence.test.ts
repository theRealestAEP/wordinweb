/**
 * CONVERGENCE GATE for the incremental relay across a block-count change.
 *
 * A splitParagraph (Enter) shifts every later block index by one. The relay
 * remaps prev-layout capture points through that shift (`shiftedBlockIdx` in
 * engine.ts) so it can still recognise a re-converged state. These tests pin
 * that remap: without it, an Enter re-lays every block to the end of the
 * document even when the layout demonstrably re-converges a page later.
 *
 * The uniform-prose case is the counterpart: there the relay CANNOT converge,
 * because the reflow is real — the added line pushes every later page boundary
 * by one paragraph. That test documents the distinction so a future reader does
 * not mistake the block count for a convergence bug.
 */
import { describe, expect, it } from "vitest";
import { DocxDocument } from "../src/docx.js";
import { layoutDocument, __incrStats } from "../src/layout/engine.js";
import { ApproxMeasurer } from "../src/layout/measure.js";
import type { XmlElement } from "../src/xml.js";
import { makeDocx, p, wrapDocument } from "./helpers.js";

const measurer = new ApproxMeasurer();

function load(body: string): DocxDocument {
  return DocxDocument.load(makeDocx({ "word/document.xml": wrapDocument(body) }));
}

/** Prose in groups separated by hard page breaks. Each break is a natural
 * re-convergence point: the page after it always restarts at bodyTop. */
function pagedProse(groups: number, perGroup: number): string {
  return Array.from({ length: groups }, (_, g) =>
    Array.from({ length: perGroup }, (_, i) => p(`g${g}-${i} alpha bravo charlie delta echo foxtrot`)).join("") +
    (g < groups - 1 ? `<w:p><w:r><w:br w:type="page"/></w:r></w:p>` : ""),
  ).join("");
}

function uniformProse(count: number): string {
  return Array.from({ length: count }, (_, i) =>
    p(`para-${i} alpha bravo charlie delta echo foxtrot golf hotel india juliet kilo lima`),
  ).join("");
}

/** Split the body paragraph at `idx` in two, as Enter does: the tail text moves
 * to a new sibling w:p and the parsed model is patched in place (no refresh, so
 * the incremental relay stays eligible). Returns the new paragraph's source. */
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

describe("incremental relay convergence across a split", () => {
  it("converges within a page of the edit when the document re-converges", () => {
    const doc = load(pagedProse(60, 20));
    const blockCount = doc.sections[0].blocks.length;
    const first = layoutDocument(doc, { measurer, windowModel: true });
    const inserted = splitParagraphAt(doc, 600);
    const second = layoutDocument(doc, { measurer, windowModel: true, prev: first, dirtyHint: inserted });

    expect(second).toBeTruthy();
    expect(__incrStats.fallbackReason).toBe("");
    // The relay must recognise the re-converged state past the shifted block
    // indices, and stop. Convergence lands at the next page break.
    expect(__incrStats.convergedBlock).toBeGreaterThan(600);
    expect(__incrStats.convergedBlock).toBeLessThan(640);
    // O(edited neighbourhood), not O(rest of document).
    expect(__incrStats.blocksLaid).toBeLessThan(50);
    expect(__incrStats.blocksLaid).toBeLessThan(blockCount / 10);
  });

  it("re-lays the remaining document only when the reflow is genuine", () => {
    // Continuous prose with no breaks: the split adds a line, which pushes the
    // last line of the page onto the next page, and so on to the end. Every
    // later page really does hold different content, so there is nothing to
    // converge onto and the block count is the necessary work, not waste.
    const doc = load(uniformProse(2000));
    const first = layoutDocument(doc, { measurer, windowModel: true });
    const inserted = splitParagraphAt(doc, 1000);
    layoutDocument(doc, { measurer, windowModel: true, prev: first, dirtyHint: inserted });

    expect(__incrStats.fallbackReason).toBe("");
    expect(__incrStats.convergedBlock).toBe(-1);
    // Page boundaries after the edit genuinely move: page 20 starts a
    // paragraph earlier than it did before the split.
    const before = layoutDocument(load(uniformProse(2000)), { measurer });
    const pageText = (result: { pages: { items: { kind: string; text?: string }[] }[] }, i: number): string =>
      result.pages[i].items.filter((it) => it.kind === "text").map((it) => it.text ?? "").join("").slice(0, 20);
    const after = layoutDocument(doc, { measurer });
    expect(pageText(after, 20)).not.toBe(pageText(before, 20));
  });
});
