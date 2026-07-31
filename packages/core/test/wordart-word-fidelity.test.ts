import { describe, expect, it } from "vitest";
import { DocxDocument } from "../src/docx.js";
import {
  drawingWordArtText,
  insertWordArtAt,
  setDrawingWordArtText,
  type WordArtPreset,
} from "../src/edit/drawings.js";
import { setFloatingPagePosition } from "../src/edit/images.js";
import { resizeDrawing } from "../src/edit/tables.js";
import type { Paragraph, Run, TextContent } from "../src/model.js";
import { makeDocx, p, wrapDocument } from "./helpers.js";

const presets: Array<[WordArtPreset, string]> = [
  ["plain", "textNoShape"],
  ["archUp", "textArchUp"],
  ["archDown", "textArchDown"],
  ["wave", "textWave1"],
  ["chevron", "textChevron"],
];

function documentWithWordArt(text: string, preset: WordArtPreset) {
  const doc = DocxDocument.load(makeDocx({ "word/document.xml": wrapDocument(p("Anchor")) }));
  const paragraph = doc.sections[0].blocks[0] as Paragraph;
  const run = paragraph.children[0] as Run;
  const sourceText = (run.content[0] as TextContent).srcT!;
  const drawing = insertWordArtAt(doc, sourceText, text, preset);
  if (!drawing) throw new Error("WordArt missing");
  return { doc, drawing };
}

describe("Microsoft Word WordArt fidelity", () => {
  it("writes centered fixed-box text geometry for every preset", () => {
    for (const [preset, warp] of presets) {
      const { doc } = documentWithWordArt("WordArt text", preset);
      const xml = DocxDocument.load(doc.save()).pkg.text("word/document.xml");

      expect(xml).toContain('<w:jc w:val="center"/>');
      expect(xml).toContain(`<a:prstTxWarp prst="${warp}">`);
      expect(xml).toContain(
        '<wps:bodyPr rot="0" spcFirstLastPara="1" vertOverflow="overflow" horzOverflow="overflow" ' +
        'vert="horz" wrap="none" lIns="91440" tIns="45720" rIns="91440" bIns="45720" ' +
        'numCol="1" spcCol="0" rtlCol="0" fromWordArt="0" anchor="t" anchorCtr="0" forceAA="0" compatLnSpc="1">',
      );
      expect(xml).toContain("<a:noAutofit/>");
    }
  });

  it("round-trips edited text, size, and page position through DOCX", () => {
    const { doc, drawing } = documentWithWordArt("Before", "archUp");
    expect(setDrawingWordArtText(doc, drawing, "Edited WordArt")).toBe(true);
    expect(resizeDrawing(doc, drawing, 360, 144)).toBe(true);
    expect(setFloatingPagePosition(doc, drawing, 140, 210)).toBe(true);

    const reopened = DocxDocument.load(doc.save());
    const paragraph = reopened.sections[0].blocks[0] as Paragraph;
    const anchor = paragraph.children
      .flatMap((child) => child.type === "run" ? child.content : [])
      .find((content) => content.kind === "anchor");
    if (!anchor || anchor.kind !== "anchor" || anchor.shape.type !== "textbox" || !anchor.shape.srcDrawing) {
      throw new Error("Reopened WordArt missing");
    }

    expect(drawingWordArtText(anchor.shape.srcDrawing)).toBe("Edited WordArt");
    const resavedXml = DocxDocument.load(reopened.save()).pkg.text("word/document.xml");
    expect(resavedXml).toContain('<wp:extent cx="3429000" cy="1371600"/>');
    expect(resavedXml).toContain('<a:ext cx="3429000" cy="1371600"/>');
    expect(resavedXml).toContain(
      '<wp:positionH relativeFrom="page"><wp:posOffset>1333500</wp:posOffset></wp:positionH>',
    );
    expect(resavedXml).toContain(
      '<wp:positionV relativeFrom="page"><wp:posOffset>2000250</wp:posOffset></wp:positionV>',
    );
    expect(resavedXml).toContain('<w:jc w:val="center"/>');
    expect(resavedXml).toContain("<a:noAutofit/>");
  });
});
