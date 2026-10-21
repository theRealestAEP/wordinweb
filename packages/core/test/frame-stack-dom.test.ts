// @vitest-environment jsdom
import { describe, expect, it } from "vitest";
import { DocxDocument } from "../src/docx.js";
import { layoutDocument } from "../src/layout/engine.js";
import { ApproxMeasurer } from "../src/layout/measure.js";
import { renderToDom } from "../src/render/dom.js";
import { makeDocx, wrapDocument } from "./helpers.js";

// The frame fix lives in the shared layout engine, so the DOM renderer (the
// editor's view) stacks a frame's paragraphs as the PDF does (word95 r3).
const FRAME = `<w:framePr w:w="4000" w:hSpace="187" w:wrap="around" w:vAnchor="text" w:hAnchor="page" w:x="1494" w:y="200"/>`;
const framed = (text: string) =>
  `<w:p><w:pPr>${FRAME}<w:spacing w:after="0"/></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`;

describe("a frame of adjacent paragraphs in the DOM renderer", () => {
  it("paints the frame's paragraphs one under another", () => {
    const body =
      `<w:p><w:r><w:t>Intro</w:t></w:r></w:p>` +
      framed("FrameOne") +
      framed("FrameTwo") +
      framed("FrameThree") +
      `<w:p><w:r><w:t xml:space="preserve">Body text after the frame, long enough to wrap beside it.</w:t></w:r></w:p>` +
      `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800"/></w:sectPr>`;
    const doc = DocxDocument.load(makeDocx({ "word/document.xml": wrapDocument(body) }));
    const layout = layoutDocument(doc, { measurer: new ApproxMeasurer() });
    const container = document.createElement("div");
    document.body.appendChild(container);
    const render = renderToDom(doc, layout, container);
    const top = (text: string) => {
      const binding = render.bindings.find((b) => b.el.textContent === text);
      if (!binding) throw new Error(`no element for "${text}"`);
      return { top: parseFloat(binding.el.style.top), left: parseFloat(binding.el.style.left), height: binding.item.lineHeight };
    };
    const lines = ["FrameOne", "FrameTwo", "FrameThree"].map(top);
    for (let k = 1; k < lines.length; k++) {
      // One line further down (the painted top is the glyph top, a little
      // below the line top, so allow a fraction of a line).
      expect(lines[k].top - lines[k - 1].top).toBeGreaterThan(lines[k - 1].height * 0.9);
      expect(lines[k].left).toBeCloseTo(lines[0].left, 1);
    }
  });
});
