import { describe, expect, it } from "vitest";
import { DocxDocument } from "../src/docx.js";
import { layoutDocument } from "../src/layout/engine.js";
import { ApproxMeasurer } from "../src/layout/measure.js";
import { makeDocx, wrapDocument } from "./helpers.js";

it("routes Lao runs to the bundled Noto face (spaces stay ascii)", () => {
  const body =
    `<w:p><w:r><w:rPr><w:rFonts w:ascii="Tahoma" w:hAnsi="Tahoma" w:cs="Tahoma"/></w:rPr>` +
    `<w:t>ສະບາຍດີ ໂລກ</w:t></w:r></w:p>` +
    `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>`;
  const doc = DocxDocument.load(makeDocx({ "word/document.xml": wrapDocument(body) }));
  const result = layoutDocument(doc, { measurer: new ApproxMeasurer() });
  const items = result.pages[0].items.filter((i) => i.kind === "text");
  expect(items.some((i) => /Noto Sans Lao/.test(i.font.family))).toBe(true);
});

it("Lao survives the hAnsi splitter (0x0E80-0x0EFF excluded like Thai)", () => {
  // hAnsi != ascii forces the splitter path that once re-clobbered Lao to hAnsi.
  const body =
    `<w:p><w:r><w:rPr><w:rFonts w:ascii="Tahoma" w:hAnsi="Calibri" w:cs="Tahoma"/></w:rPr>` +
    `<w:t>ຂໍ້ຄວາມ</w:t></w:r></w:p>` +
    `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>`;
  const doc = DocxDocument.load(makeDocx({ "word/document.xml": wrapDocument(body) }));
  const result = layoutDocument(doc, { measurer: new ApproxMeasurer() });
  const lao = result.pages[0].items.filter((i) => i.kind === "text" && /[຀-໿]/.test(i.text));
  expect(lao.length).toBeGreaterThan(0);
  for (const i of lao) expect(i.kind === "text" && i.font.family).toBe("Noto Sans Lao Looped");
});
