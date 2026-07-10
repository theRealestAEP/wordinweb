import { describe, expect, it } from "vitest";
import { DocxDocument } from "../src/docx.js";
import { layoutDocument } from "../src/layout/engine.js";
import { ApproxMeasurer } from "../src/layout/measure.js";
import { makeDocx, wrapDocument } from "./helpers.js";

describe("hyperlink URL rendering", () => {
  it("renders a long unbreakable hyperlink URL completely", () => {
    const para = `<w:p><w:pPr><w:ind w:left="1440" w:right="1440"/></w:pPr>
      <w:r><w:t xml:space="preserve">This baji is wamuqum at:</w:t></w:r>
      <w:hyperlink r:id="rId99" w:history="1">
        <w:r><w:t xml:space="preserve"> </w:t></w:r>
        <w:r><w:rPr><w:color w:val="2B60DE"/></w:rPr><w:t>wamuv://deta.oja.ata/papekew/pajuhujosenorosepa/vojekoqewip/Corinazib/Harujipaguduh.loh</w:t></w:r>
        <w:r><w:t xml:space="preserve"> </w:t></w:r>
      </w:hyperlink>
      <w:r><w:t>.</w:t></w:r></w:p>`;
    const doc = DocxDocument.load(makeDocx({ "word/document.xml": wrapDocument(para) }));
    const result = layoutDocument(doc, { measurer: new ApproxMeasurer() });
    const texts = result.pages[0].items
      .filter((i) => i.kind === "text")
      .map((i) => (i as { text: string }).text);
    expect(texts.join("")).toContain("wamuv://deta.oja.ata/papekew");
    expect(texts.join("")).toContain("Harujipaguduh.loh");
  });
});
