// @vitest-environment jsdom
import { describe, expect, it } from "vitest";
import { DocxDocument } from "../src/docx.js";
import { invalidateParagraphSignature } from "../src/layout/inline.js";
import { ApproxMeasurer } from "../src/layout/measure.js";
import { layoutDocument } from "../src/layout/engine.js";
import { renderToDom } from "../src/render/dom.js";
import type { Paragraph, Run, TextContent } from "../src/model.js";
import { makeDocx, p, wrapDocument } from "./helpers.js";

describe("incremental DOM render", () => {
  it("updates text in the retained page when geometry stays compatible", () => {
    const doc = DocxDocument.load(
      makeDocx({
        "word/document.xml": wrapDocument(p("alpha bravo charlie") + p("delta echo foxtrot")),
      }),
    );
    const measurer = new ApproxMeasurer();
    const first = layoutDocument(doc, { measurer });
    const container = document.createElement("div");
    document.body.appendChild(container);
    const firstRender = renderToDom(doc, first, container);
    const pageElement = firstRender._pages![0].el;
    const textElements = firstRender.bindings.map((binding) => binding.el);

    const paragraph = doc.sections[0].blocks[0] as Paragraph;
    const run = paragraph.children[0] as Run;
    const text = run.content[0] as TextContent;
    text.text += "Z";
    text.srcT!.text = text.text;
    invalidateParagraphSignature(paragraph.src!);
    const second = layoutDocument(doc, {
      measurer,
      prev: first,
      dirtyHint: paragraph.src,
      dirtySource: text.srcT,
    });
    const secondRender = renderToDom(doc, second, container, {}, firstRender);

    expect(secondRender._pages![0].el).toBe(pageElement);
    expect(secondRender.bindings.map((binding) => binding.el)).toEqual(textElements);
    expect(secondRender.root.textContent).toContain("charlieZ");
  });
});
