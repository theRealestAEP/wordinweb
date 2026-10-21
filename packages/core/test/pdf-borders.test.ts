import { Writable } from "node:stream";
import { inflateSync } from "node:zlib";
import { describe, expect, it } from "vitest";
import { DocxDocument } from "../src/docx.js";
import { layoutDocument } from "../src/layout/engine.js";
import { FontSet } from "../src/pdf/fonts.js";
import { NodeMeasurer } from "../src/pdf/measure.js";
import { renderDocxToPdf } from "../src/pdf/render.js";
import { bytesSource } from "../src/stream/source.js";
import { makeDocx, wrapDocument } from "./helpers.js";

function borderedDocx(sizes = [24, 24, 24, 24], tableBorders = false, shadow = true): Uint8Array {
  const borders = ["top", "left", "bottom", "right"].map((side, i) =>
    `<w:${side} w:val="single" w:sz="${sizes[i]}" w:space="4" w:color="000000" w:shadow="${shadow ? 1 : 0}"/>`,
  ).join("");
  const body = `<w:p><w:pPr><w:pBdr>${borders}</w:pBdr><w:shd w:fill="DDEBF7"/><w:spacing w:after="480"/></w:pPr><w:r><w:t>Paragraph: thick border and shadow</w:t></w:r></w:p>` +
    `<w:tbl><w:tblPr><w:tblW w:w="6000" w:type="dxa"/><w:tblLayout w:type="fixed"/>${tableBorders ? `<w:tblBorders>${borders}</w:tblBorders>` : ""}</w:tblPr><w:tblGrid><w:gridCol w:w="6000"/></w:tblGrid>` +
    `<w:tr><w:trPr><w:trHeight w:val="900" w:hRule="atLeast"/></w:trPr><w:tc><w:tcPr><w:tcW w:w="6000" w:type="dxa"/>${tableBorders ? "" : `<w:tcBorders>${borders}</w:tcBorders>`}<w:shd w:fill="FFF2CC"/></w:tcPr><w:p><w:r><w:t>Cell: thick border and shadow</w:t></w:r></w:p></w:tc></w:tr></w:tbl>` +
    `<w:p/><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440"/></w:sectPr>`;
  return makeDocx({ "word/document.xml": wrapDocument(body) });
}

async function contentStream(docx: Uint8Array): Promise<string> {
  const chunks: Buffer[] = [];
  await renderDocxToPdf(bytesSource(docx), new Writable({
    write(chunk: Buffer, _encoding, callback) { chunks.push(chunk); callback(); },
  }));
  const bytes = Buffer.concat(chunks);
  const text = bytes.toString("latin1");
  return [...text.matchAll(/\/Length (\d+) >>\nstream\n/g)].map((m) =>
    inflateSync(bytes.subarray(m.index! + m[0].length, m.index! + m[0].length + Number(m[1]))).toString("latin1"),
  ).join("\n");
}

// PDF coordinates and widths are rounded to three decimal places.
const rounded = (value: number) => Math.round(value * 1000) / 1000;

describe("PDF box border corners", () => {
  it.each([
    { name: "3pt cell borders", sizes: [24, 24, 24, 24], tableBorders: false },
    { name: "unequal cell borders", sizes: [8, 16, 32, 48], tableBorders: false },
    { name: "thin declared cell borders", sizes: [2, 4, 6, 8], tableBorders: false },
    { name: "inherited table borders", sizes: [24, 24, 24, 24], tableBorders: true },
  ])("joins all four corners and adjacent shadow strips: $name", async ({ sizes, tableBorders }) => {
    const ops = await contentStream(borderedDocx(sizes, tableBorders));
    const rules = [...ops.matchAll(/q 0 J 0 0 0 RG ([\d.]+) w ([\d.]+) ([\d.]+) m ([\d.]+) ([\d.]+) l S Q/g)]
      .map((m) => ({ w: +m[1], x1: +m[2], y1: +m[3], x2: +m[4], y2: +m[5] }));
    const shadows = [...ops.matchAll(/q 0 0 0 rg ([\d.]+) ([\d.]+) ([\d.]+) ([\d.]+) re f Q/g)]
      .map((m) => ({ x: +m[1], y: +m[2], w: +m[3], h: +m[4] }));
    expect(rules).toHaveLength(8);
    expect(shadows).toHaveLength(4);
    for (const i of [0, 1]) {
      const [top, bottom, left, right] = rules.slice(i * 4, i * 4 + 4);
      const x0 = rounded(left.x1 - left.w / 2);
      const x1 = rounded(right.x1 + right.w / 2);
      const y0 = rounded(bottom.y1 - bottom.w / 2);
      const y1 = rounded(top.y1 + top.w / 2);
      // Every butt-capped segment ends exactly at the adjacent rule's outer edge.
      for (const h of [top, bottom]) expect([h.x1, h.x2]).toEqual([x0, x1]);
      for (const v of [left, right]) expect([v.y1, v.y2]).toEqual([y1, y0]);
      expect([top.w, left.w, bottom.w, right.w]).toEqual(sizes.map((sz) => sz / 8));
      const [bottomShadow, rightShadow] = shadows.slice(i * 2, i * 2 + 2);
      expect(bottomShadow).toEqual({ x: x0, y: rounded(y0 - bottom.w), w: expect.closeTo(x1 - x0 + right.w, 2), h: bottom.w });
      expect(rightShadow).toEqual({ x: x1, y: y0, w: right.w, h: expect.closeTo(y1 - y0, 2) });
    }
    // Both shaded interiors remain filled.
    expect(ops).toContain("0.867 0.922 0.969 rg");
    expect(ops).toContain("1 0.949 0.8 rg");
  });

  it("keeps text positions and line breaks when shadow is enabled", () => {
    const measurer = new NodeMeasurer(new FontSet());
    const textPositions = (shadow: boolean) => layoutDocument(DocxDocument.load(borderedDocx(undefined, false, shadow)), { measurer })
      .pages.map((page) => page.items.filter((item) => item.kind === "text")
        .map(({ text, x, width, baseline, lineTop, lineHeight }) => ({ text, x, width, baseline, lineTop, lineHeight })));
    expect(textPositions(true)).toEqual(textPositions(false));
  });
});
