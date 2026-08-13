/**
 * #159: a floating object we create or drag has to land ON TOP of what is
 * already there, the way Word orders them.
 *
 * Every creation site stamped the constant `relativeHeight="251658240"`. That
 * is Word's own LOWEST ordinary value, so anything we made or floated landed
 * UNDER every picture Word had written — measured across the corpus at
 * 251659264 and above (coverletter-anon, YN, dense-skewtest, wild2-math), and
 * up to 251699200 in wild2-med. From then on the picture took every click
 * aimed at the object, and because the number is written into the file it
 * stayed buried across save and reopen.
 *
 * WHY THE FIXTURES ARE BUILT HERE RATHER THAN CHECKED IN. A pristine corpus
 * document does not reproduce this: opening a Word file gives correct z-order,
 * and it takes an edit by US to create the bad ordering. A test built on an
 * untouched document would pass whatever the code did. So each case below
 * makes the document AND performs the edit.
 */
import { describe, expect, it } from "vitest";
import { DocxDocument } from "../src/docx.js";
import { insertShapeAt } from "../src/edit/drawings.js";
import { setImageWrap } from "../src/edit/images.js";
import { layoutDocument } from "../src/layout/engine.js";
import { ApproxMeasurer } from "../src/layout/measure.js";
import { localName, XmlElement } from "../src/xml.js";
import { makeDocxWithMedia, wrapDocument, p } from "./helpers.js";
import type { Paragraph, Run } from "../src/model.js";

/** What Word writes for the first floating object in a document. */
const WORD_FIRST = 251658240;
/** What Word had written in the documents this was measured on. */
const WORD_EXISTING = 251659264;

const EMU = 9525;
const PNG = new Uint8Array([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
  0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x02, 0x00, 0x00, 0x00, 0x90, 0x77, 0x53,
  0xde, 0x00, 0x00, 0x00, 0x0c, 0x49, 0x44, 0x41, 0x54, 0x08, 0xd7, 0x63, 0xf8, 0xcf, 0xc0, 0x00,
  0x00, 0x03, 0x01, 0x01, 0x00, 0x18, 0xdd, 0x8d, 0xb0, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e,
  0x44, 0xae, 0x42, 0x60, 0x82,
]);

/** A floating picture, exactly as Word anchors one. */
function floatingPicture(relativeHeight: number, size = 150): string {
  const px = String(size * EMU);
  return `<w:p><w:r><w:drawing>
    <wp:anchor xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
        distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="${relativeHeight}"
        behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">
      <wp:simplePos x="0" y="0"/>
      <wp:positionH relativeFrom="page"><wp:posOffset>${120 * EMU}</wp:posOffset></wp:positionH>
      <wp:positionV relativeFrom="page"><wp:posOffset>${120 * EMU}</wp:posOffset></wp:positionV>
      <wp:extent cx="${px}" cy="${px}"/><wp:wrapNone/><wp:docPr id="9" name="Pic"/>
      <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
          <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
            <pic:nvPicPr><pic:cNvPr id="0" name="p.png"/><pic:cNvPicPr/></pic:nvPicPr>
            <pic:blipFill><a:blip r:embed="rId9" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
            <pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${px}" cy="${px}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>
          </pic:pic>
        </a:graphicData>
      </a:graphic>
    </wp:anchor>
  </w:drawing></w:r></w:p>`;
}

/** An INLINE picture — the thing a drag converts to floating. */
const INLINE_PICTURE = `<w:p><w:r><w:drawing>
  <wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">
    <wp:extent cx="${100 * EMU}" cy="${100 * EMU}"/><wp:docPr id="7" name="Inline"/>
    <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
        <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
          <pic:nvPicPr><pic:cNvPr id="0" name="p.png"/><pic:cNvPicPr/></pic:nvPicPr>
          <pic:blipFill><a:blip r:embed="rId9" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
          <pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${100 * EMU}" cy="${100 * EMU}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>
        </pic:pic>
      </a:graphicData>
    </a:graphic>
  </wp:inline>
</w:drawing></w:r></w:p>`;

function load(body: string): DocxDocument {
  return DocxDocument.load(makeDocxWithMedia(
    {
      "word/document.xml": wrapDocument(body),
      "word/_rels/document.xml.rels": `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/p.png"/>
</Relationships>`,
    },
    { "word/media/p.png": PNG },
  ));
}

/** Every relativeHeight in the saved file, in document order. */
function savedHeights(doc: DocxDocument): number[] {
  const xml = new TextDecoder().decode(
    // Re-open so we read what a user's Word would read, not the in-memory tree.
    DocxDocument.load(doc.save()).save(),
  );
  void xml;
  const out: number[] = [];
  const walk = (node: XmlElement): void => {
    if (localName(node.name) === "anchor") {
      const key = Object.keys(node.attrs).find((k) => localName(k) === "relativeHeight");
      if (key) out.push(parseInt(node.attrs[key], 10));
    }
    for (const c of node.children) walk(c);
  };
  for (const section of DocxDocument.load(doc.save()).sections) {
    for (const block of section.blocks) {
      const src = (block as { src?: XmlElement }).src;
      if (src) walk(src);
    }
  }
  return out;
}

function lastParagraphText(doc: DocxDocument): XmlElement {
  const paras = doc.sections.flatMap((s) => s.blocks).filter((b): b is Paragraph => b.type === "paragraph");
  for (let i = paras.length - 1; i >= 0; i--) {
    const run = paras[i].children.flatMap((pc) => (pc.type === "run" ? [pc] : pc.runs))[0] as Run | undefined;
    const t = run?.content.find((c) => c.kind === "text");
    if (t && t.kind === "text" && t.srcT) return t.srcT;
  }
  throw new Error("no text to put the caret on");
}

function drawingElOf(doc: DocxDocument): XmlElement {
  let found: XmlElement | undefined;
  const walk = (node: XmlElement): void => {
    if (!found && localName(node.name) === "drawing") found = node;
    for (const c of node.children) if (!found) walk(c);
  };
  for (const section of doc.sections) {
    for (const block of section.blocks) {
      const src = (block as { src?: XmlElement }).src;
      if (src && !found) walk(src);
    }
  }
  if (!found) throw new Error("no drawing in the document");
  return found;
}

describe("#159 · a new floating object lands on top", () => {
  it("gives an inserted shape a z above the picture Word already anchored", () => {
    const doc = load(floatingPicture(WORD_EXISTING) + p("caret here"));
    expect(insertShapeAt(doc, lastParagraphText(doc), "roundedRectangle", "NEW")).toBeTruthy();
    const heights = savedHeights(doc);
    expect(heights).toContain(WORD_EXISTING);
    const made = heights.filter((h) => h !== WORD_EXISTING);
    expect(made.length, "the insert wrote no anchor").toBe(1);
    // Was 251658240 — 1024 BELOW the picture, so the picture covered it.
    expect(made[0], `a new shape at ${made[0]} sits under the existing ${WORD_EXISTING}`)
      .toBeGreaterThan(WORD_EXISTING);
  });

  it("gives a dragged-out picture a z above the one already anchored", () => {
    const doc = load(floatingPicture(WORD_EXISTING) + INLINE_PICTURE + p("tail"));
    // The inline picture is the second drawing; take it by skipping the first.
    const drawings: XmlElement[] = [];
    const walk = (node: XmlElement): void => {
      if (localName(node.name) === "drawing") drawings.push(node);
      for (const c of node.children) walk(c);
    };
    for (const section of doc.sections) {
      for (const block of section.blocks) {
        const src = (block as { src?: XmlElement }).src;
        if (src) walk(src);
      }
    }
    expect(drawings.length).toBe(2);
    expect(setImageWrap(doc, drawings[1], "square", { x: 40, y: 40 })).toBe(true);
    const made = savedHeights(doc).filter((h) => h !== WORD_EXISTING);
    expect(made.length, "the drag wrote no anchor").toBe(1);
    expect(made[0], "a dragged picture landed under the one already there")
      .toBeGreaterThan(WORD_EXISTING);
  });

  it("starts where Word starts when nothing is floating yet", () => {
    const doc = load(p("caret here"));
    expect(insertShapeAt(doc, lastParagraphText(doc), "roundedRectangle", "FIRST")).toBeTruthy();
    expect(savedHeights(doc)).toEqual([WORD_FIRST]);
  });

  it("keeps stacking upwards over repeated inserts", () => {
    const doc = load(floatingPicture(WORD_EXISTING) + p("caret here"));
    for (let i = 0; i < 5; i++) {
      expect(insertShapeAt(doc, lastParagraphText(doc), "roundedRectangle", `S${i}`)).toBeTruthy();
    }
    const heights = savedHeights(doc);
    const sorted = [...heights].sort((a, b) => a - b);
    expect(new Set(heights).size, "two objects share a z, so their order is arbitrary").toBe(heights.length);
    // Six inserts from 251659264 reach 251659270 — nowhere near the 2^31 a
    // browser clamps a z-index at, and Word's own bring-to-front does the same.
    expect(sorted[sorted.length - 1]).toBeLessThan(2 ** 31 - 1);
    expect(sorted[sorted.length - 1] - WORD_EXISTING).toBe(5);
  });

  it("paints the new shape above the existing picture, not under it", () => {
    const doc = load(floatingPicture(WORD_EXISTING) + p("caret here"));
    expect(insertShapeAt(doc, lastParagraphText(doc), "roundedRectangle", "NEW")).toBeTruthy();
    const pages = layoutDocument(DocxDocument.load(doc.save()), new ApproxMeasurer());
    const items = pages.pages.flatMap((page) => page.items);
    const picture = items.find((i) => i.kind === "image");
    const hit = items.find((i) => i.kind === "drawingHit" && i.anchored);
    expect(picture, "the picture stopped rendering").toBeTruthy();
    expect(hit, "the inserted shape has no hit target").toBeTruthy();
    // This is the effect the user feels: the click goes to whichever is on top.
    expect((hit as { z?: number }).z ?? -1).toBeGreaterThan((picture as { z?: number }).z ?? 0);
  });
});
