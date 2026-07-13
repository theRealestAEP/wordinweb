import { it, expect } from "vitest";
import { DocxDocument } from "../src/docx.js";
import { layoutDocument } from "../src/layout/engine.js";
import { ApproxMeasurer } from "../src/layout/measure.js";
import { makeDocx, wrapDocument } from "./helpers.js";
it("cell trailing spaces carry caretClampX", () => {
  const tbl =
    '<w:tbl><w:tblPr><w:tblW w:w="4000" w:type="dxa"/></w:tblPr>' +
    '<w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>' +
    '<w:tr><w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>' +
    '<w:p><w:r><w:t xml:space="preserve">content-        </w:t></w:r></w:p></w:tc>' +
    '<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>right cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>';
  const sect = '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>';
  const doc = DocxDocument.load(makeDocx({ "word/document.xml": wrapDocument(tbl + sect) }));
  const result = layoutDocument(doc, { measurer: new ApproxMeasurer() });
  const spaces = result.pages[0].items.filter((i) => i.kind === "text" && i.text === " ");
  console.log("space items:", spaces.length, "clamped:", spaces.filter((s: any) => s.caretClampX !== undefined).length);
  for (const s of spaces.slice(0, 3)) console.log("  x=", (s as any).x.toFixed(1), "clamp=", (s as any).caretClampX);
  expect(spaces.some((s: any) => s.caretClampX !== undefined)).toBe(true);
});
