import { describe, expect, it } from "vitest";
import { DocxDocument } from "../src/docx.js";
import { moveDrawingTo } from "../src/edit/tables.js";
import { insertShapeAt } from "../src/edit/drawings.js";
import { localName, XmlElement } from "../src/xml.js";
import { Paragraph, Run } from "../src/model.js";
import { makeDocxWithMedia, wrapDocument, p } from "./helpers.js";

// A header image's r:embed relationship is scoped to the header part. Moving
// its run into the body part (or vice versa) leaves the rel dangling and the
// picture renders nowhere — "the image disappeared". moveDrawingTo must refuse
// cross-part moves; the editor's region gate keeps such a drag from ever
// starting, and this guard is the model-level backstop.

const INLINE_DRAWING = `<w:r><w:drawing>
  <wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">
    <wp:extent cx="914400" cy="914400"/>
    <wp:docPr id="1" name="HdrPic"/>
    <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
        <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
          <pic:blipFill><a:blip r:embed="rIdImg" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/></pic:blipFill>
          <pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm></pic:spPr>
        </pic:pic>
      </a:graphicData>
    </a:graphic>
  </wp:inline>
</w:drawing></w:r>`;

const PNG = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);

/** Doc with body text + a header holding an inline image (and, optionally, a
 * second header paragraph of text as an in-part move target). */
function loadHeaderImageDoc(extraHeaderPara = ""): DocxDocument {
  const header = `<?xml version="1.0"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p>${INLINE_DRAWING}</w:p>${extraHeaderPara}</w:hdr>`;
  return DocxDocument.load(
    makeDocxWithMedia(
      {
        "word/document.xml": wrapDocument(
          p("BODYTEXT") +
            `<w:sectPr><w:headerReference xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" w:type="default" r:id="rId5"/><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>`,
        ),
        "word/_rels/document.xml.rels": `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
</Relationships>`,
        "word/header1.xml": header,
        "word/_rels/header1.xml.rels": `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdImg" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
</Relationships>`,
      },
      { "word/media/image1.png": PNG },
    ),
  );
}

/** The header's inline w:drawing element. */
function headerDrawing(doc: DocxDocument): XmlElement {
  const hdr = doc.headers.get("rId5")!;
  const run = (hdr.blocks[0] as Paragraph).children[0] as Run;
  const img = run.content.find((c) => c.kind === "image");
  if (!img || img.kind !== "image" || !img.srcDrawing) throw new Error("no header image");
  return img.srcDrawing;
}

function firstText(para: Paragraph): XmlElement {
  const run = para.children.flatMap((pc) => (pc.type === "run" ? [pc] : pc.runs))[0] as Run;
  const t = run.content.find((c) => c.kind === "text");
  if (!t || t.kind !== "text" || !t.srcT) throw new Error("no text");
  return t.srcT;
}

describe("header image drag safety (moveDrawingTo cross-part guard)", () => {
  it("refuses to move a header image onto body text (would dangle its rel)", () => {
    const doc = loadHeaderImageDoc();
    const drawing = headerDrawing(doc);
    const bodyT = firstText(doc.sections[0].blocks[0] as Paragraph);
    expect(moveDrawingTo(doc, drawing, bodyT)).toBe(false);
    // The image is still in the header, unharmed.
    const hdr = doc.headers.get("rId5")!;
    const stillThere = (hdr.blocks[0] as Paragraph).children
      .flatMap((pc) => (pc.type === "run" ? [pc] : pc.runs))
      .some((r) => r.content.some((c) => c.kind === "image"));
    expect(stillThere).toBe(true);
  });

  it("allows moving a header image within the same header part", () => {
    const doc = loadHeaderImageDoc(`<w:p><w:r><w:t>HDRLABEL</w:t></w:r></w:p>`);
    const drawing = headerDrawing(doc);
    const hdr = doc.headers.get("rId5")!;
    const labelT = firstText(hdr.blocks[1] as Paragraph);
    expect(moveDrawingTo(doc, drawing, labelT)).toBe(true);
    // The same drawing element is still in the header part (its root is w:hdr):
    // refresh() re-derives the model but keeps the live XML tree.
    let root: XmlElement = drawing;
    for (let par = doc.findParentOf(root); par; par = doc.findParentOf(par)) root = par;
    expect(localName(root.name)).toBe("hdr");
  });

  it("refuses to move a shape run into its own editable textbox", () => {
    const doc = DocxDocument.load(makeDocxWithMedia({ "word/document.xml": wrapDocument(p("ANCHOR")) }, {}));
    const anchor = firstText(doc.sections[0].blocks[0] as Paragraph);
    const drawing = insertShapeAt(doc, anchor, "roundedRectangle", "SELF TARGET")!;
    let ownText: XmlElement | undefined;
    const findOwnText = (node: XmlElement): void => {
      if (localName(node.name) === "t" && node.text === "SELF TARGET") ownText = node;
      for (const child of node.children) findOwnText(child);
    };
    findOwnText(drawing);
    expect(ownText).toBeDefined();
    expect(moveDrawingTo(doc, drawing, ownText!)).toBe(false);

    const reloaded = DocxDocument.load(doc.save());
    const xml = reloaded.pkg.text("word/document.xml");
    expect((xml.match(/<w:drawing>/g) ?? [])).toHaveLength(1);
    expect(xml).toContain("SELF TARGET");
  });
});
