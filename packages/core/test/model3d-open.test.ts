/**
 * OPENING a document that contains an Office 3D model (am3d).
 *
 * Reported as "file open just doesn't work, at least for 3d objects". The
 * round trip turned out to work — this file is what makes that claim
 * checkable rather than a memory, because the only 3D coverage the suite had
 * (edit.test.ts "packages a native 3D model…") asserts on a document THIS
 * codebase wrote, inline, one paragraph long. Our writer and our parser can
 * agree with each other and both be wrong about the format Word emits.
 *
 * So the fixture below is the shape a REAL Word file has, transcribed from a
 * Word 16.0.8326 document ("Hot Face Emoji" inserted into a cover-letter
 * template). Three things in it differ from anything our writer produces, and
 * each one is a way the open path could break without a single existing test
 * noticing:
 *
 *   1. wp:anchor, NOT wp:inline — Word floats a 3D model by default. That is
 *      a completely different branch of parseDrawing (AnchorContent, not
 *      ImageContent), and model3D has to survive it.
 *   2. a non-zero am3d:rot — the model carries an orientation, which the
 *      viewer needs and which only this path reads.
 *   3. several am3d:ptLight siblings, wp14:sizeRel*, and a wrapNone anchor.
 *
 * The unreadable-drawing cases at the bottom carry the most weight of all:
 * a 3D model we cannot make sense of must never stop the DOCUMENT from
 * opening. Losing a model is a missing picture; throwing is a file the user
 * cannot open at all, and those are not the same bug.
 */
import { describe, expect, it } from "vitest";
import { DocxDocument } from "../src/docx.js";
import { insertModel3DAt } from "../src/edit/objects.js";
import { layoutDocument } from "../src/layout/engine.js";
import { Model3DReference, Paragraph, TextContent } from "../src/model.js";
import { makeDocxWithMedia, p, wrapDocument } from "./helpers.js";

const POSTER = new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]);
const GLB = new Uint8Array([0x67, 0x6c, 0x54, 0x46, 2, 0, 0, 0, 12, 0, 0, 0]);

/** Every 3D model the laid-out document actually paints. */
function laidOutModels(doc: DocxDocument): Model3DReference[] {
  return layoutDocument(doc).pages
    .flatMap((page) => page.items)
    .map((item) => (item as { model3D?: Model3DReference }).model3D)
    .filter((m): m is Model3DReference => !!m);
}

/** Concatenated text of the laid-out document — "did the rest of it open". */
function laidOutText(doc: DocxDocument): string {
  return layoutDocument(doc).pages
    .flatMap((page) => page.items)
    .map((item) => (item as { text?: string }).text ?? "")
    .join("");
}

/** Laid-out items whose geometry came out non-finite. A drawing we cannot
 * read must not put NaN into the page — that survives the parse but paints a
 * blank or collapsed page, which reads to a user as "it didn't open". */
function itemsWithBrokenGeometry(doc: DocxDocument): string[] {
  const broken: string[] = [];
  for (const page of layoutDocument(doc).pages) {
    for (const item of page.items) {
      const box = item as { x?: number; y?: number; width?: number; height?: number };
      for (const [name, value] of Object.entries(box)) {
        if (typeof value === "number" && !Number.isFinite(value)) broken.push(`${item.kind}.${name}`);
      }
    }
  }
  return broken;
}

describe("opening a document with an Office 3D model", () => {
  it("reads a model out of a Word-shaped FLOATING drawing", () => {
    const doc = DocxDocument.load(wordAuthoredModelDocx());

    const models = laidOutModels(doc);
    expect(models).toHaveLength(1);
    expect(models[0].part).toBe("word/media/model3d1.glb");
    expect(models[0].posterPart).toBe("word/media/image1.png");
    // am3d:rot is in 60000ths of a degree: ax=1647008 -> 27.45°.
    expect(models[0].rotation?.x).toBeCloseTo(27.45, 2);
    expect(models[0].rotation?.y).toBeCloseTo(-22.64, 2);
    expect(models[0].rotation?.z).toBeCloseTo(-11.31, 2);
    // The model rides ON the document's own text, not instead of it.
    expect(laidOutText(doc)).toContain("Curriculum vitae");
  });

  it("keeps the model through a save/reopen of a Word-shaped document", () => {
    const doc = DocxDocument.load(wordAuthoredModelDocx());
    const reopened = DocxDocument.load(doc.save());

    // The .glb part itself — byte-identical, not merely present.
    expect(reopened.pkg.binary("word/media/model3d1.glb")).toEqual(GLB);
    // Its relationship, under Office's 2017 model3d type.
    expect(reopened.pkg.text("word/_rels/document.xml.rels"))
      .toContain('Type="http://schemas.microsoft.com/office/2017/06/relationships/model3d"');
    // And the content-type default without which Word rejects the package.
    expect(reopened.pkg.text("[Content_Types].xml"))
      .toContain('Extension="glb" ContentType="model/gltf-binary"');
    // Still a model after the round trip, not just a loose part in the zip.
    expect(laidOutModels(reopened)).toHaveLength(1);
  });

  it("round-trips a model this codebase inserted", () => {
    const doc = DocxDocument.load(makeDocxWithMedia(
      { "word/document.xml": wrapDocument(p("Anchor")) },
      {},
    ));
    const para = doc.sections[0].blocks[0] as Paragraph;
    const anchorT = ((para.children[0] as { content: TextContent[] }).content[0] as TextContent).srcT!;
    expect(insertModel3DAt(doc, anchorT, { data: GLB, poster: POSTER })).toBe(true);

    const reopened = DocxDocument.load(doc.save());
    expect(laidOutModels(reopened)).toEqual([
      { part: "word/media/model3d1.glb", posterPart: "word/media/image1.png" },
    ]);
  });

  // THE ONE THAT MATTERS. A 3D model we can't read is a missing picture. A 3D
  // model that throws is a document nobody can open.
  it.each([
    ["model3d with no children at all", `<am3d:model3d r:embed="rId5"/>`],
    ["raster with no blip", `<am3d:model3d r:embed="rId5"><am3d:raster/></am3d:model3d>`],
    ["model relationship missing from .rels", `<am3d:model3d r:embed="rId404"><am3d:raster><am3d:blip r:embed="rId4"/></am3d:raster></am3d:model3d>`],
    ["a graphicData nobody has ever heard of", `<zz:whatever xmlns:zz="http://example.invalid/2099"><zz:inner v="1"/></zz:whatever>`],
  ])("still opens the document when the drawing is unreadable: %s", (_name, graphicBody) => {
    const doc = DocxDocument.load(malformedModelDocx(graphicBody));
    expect(laidOutText(doc)).toContain("Curriculum vitae");
    expect(itemsWithBrokenGeometry(doc)).toEqual([]);
  });

  // Garbage numerics are their own case: the drawing is READABLE, so losing
  // it would be a silent downgrade rather than a crash. It has to survive,
  // fall back to the wp:extent for its size, and put no NaN in the page.
  it("keeps a model whose am3d numbers are unparseable", () => {
    const doc = DocxDocument.load(malformedModelDocx(
      `<am3d:model3d r:embed="rId5">` +
      `<am3d:spPr><a:xfrm><a:off x="NaN" y="oops"/><a:ext cx="-" cy="[]"/></a:xfrm></am3d:spPr>` +
      `<am3d:trans><am3d:rot ax="abc" ay=""/></am3d:trans>` +
      `<am3d:raster><am3d:blip r:embed="rId4"/></am3d:raster></am3d:model3d>`,
    ));
    const models = laidOutModels(doc);
    expect(models).toHaveLength(1);
    expect(models[0].part).toBe("word/media/model3d1.glb");
    // Unreadable angles read as no rotation, never NaN.
    expect(models[0].rotation).toBeUndefined();
    expect(itemsWithBrokenGeometry(doc)).toEqual([]);
    expect(laidOutText(doc)).toContain("Curriculum vitae");
  });
});

// ---------------------------------------------------------------- fixtures

const NS = [
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"',
  'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"',
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"',
  'xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"',
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"',
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"',
  'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"',
  'xmlns:am3d="http://schemas.microsoft.com/office/drawing/2017/model3d"',
  'mc:Ignorable="w14 wp14"',
].join(" ");

const RELS =
  `<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>` +
  `<Relationship Id="rId5" Type="http://schemas.microsoft.com/office/2017/06/relationships/model3d" Target="media/model3d1.glb"/>`;

const CONTENT_TYPES =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
  `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
  `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
  `<Default Extension="xml" ContentType="application/xml"/>` +
  `<Default Extension="png" ContentType="image/png"/>` +
  `<Default Extension="glb" ContentType="model/gltf-binary"/>` +
  `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
  `</Types>`;

function pack(body: string): Uint8Array {
  return makeDocxWithMedia(
    {
      "word/document.xml":
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${NS}><w:body>` +
        `${body}<w:p><w:r><w:t xml:space="preserve">Curriculum vitae</w:t></w:r></w:p>` +
        `</w:body></w:document>`,
      "word/_rels/document.xml.rels":
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
        `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${RELS}</Relationships>`,
      "[Content_Types].xml": CONTENT_TYPES,
    },
    { "word/media/image1.png": POSTER, "word/media/model3d1.glb": GLB },
  );
}

/**
 * The am3d drawing as Word 16.0.8326 writes it: FLOATING (wp:anchor +
 * wrapNone), a real orientation on am3d:rot, three point lights, and the
 * wp14 size-relative pair Word appends. Transcribed from a genuine file;
 * trimmed only of the surrounding table and content controls.
 */
function wordAuthoredModelDocx(): Uint8Array {
  const anchor = (graphic: string) =>
    `<w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0"` +
    ` relativeHeight="251659264" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"` +
    ` wp14:anchorId="74953950" wp14:editId="3040AA58">` +
    `<wp:simplePos x="0" y="0"/>` +
    `<wp:positionH relativeFrom="column"><wp:posOffset>71755</wp:posOffset></wp:positionH>` +
    `<wp:positionV relativeFrom="paragraph"><wp:posOffset>255905</wp:posOffset></wp:positionV>` +
    `<wp:extent cx="1828800" cy="2082800"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>` +
    `<wp:docPr id="1677382972" name="3D Model 43" descr="Hot Face Emoji"/>` +
    `<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>` +
    `${graphic}` +
    `<wp14:sizeRelH relativeFrom="margin"><wp14:pctWidth>0</wp14:pctWidth></wp14:sizeRelH>` +
    `<wp14:sizeRelV relativeFrom="margin"><wp14:pctHeight>0</wp14:pctHeight></wp14:sizeRelV>` +
    `</wp:anchor></w:drawing>`;
  const modelGraphic =
    `<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/drawing/2017/model3d">` +
    `<am3d:model3d r:embed="rId5">` +
    `<am3d:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="2082800"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></am3d:spPr>` +
    `<am3d:camera><am3d:pos x="0" y="0" z="77741944"/><am3d:up dx="0" dy="36000000" dz="0"/>` +
    `<am3d:lookAt x="0" y="0" z="0"/><am3d:perspective fov="2700000"/></am3d:camera>` +
    `<am3d:trans><am3d:meterPerModelUnit n="88662" d="1000000"/>` +
    `<am3d:preTrans dx="2152" dy="-16757246" dz="-1248832"/>` +
    `<am3d:scale><am3d:sx n="1000000" d="1000000"/><am3d:sy n="1000000" d="1000000"/>` +
    `<am3d:sz n="1000000" d="1000000"/></am3d:scale>` +
    `<am3d:rot ax="1647008" ay="-1358698" az="-678617"/>` +
    `<am3d:postTrans dx="0" dy="0" dz="0"/></am3d:trans>` +
    `<am3d:raster rName="Office3DRenderer" rVer="16.0.8326"><am3d:blip r:embed="rId4"/></am3d:raster>` +
    `<am3d:objViewport viewportSz="3200400"/>` +
    `<am3d:ambientLight><am3d:clr><a:scrgbClr r="50000" g="50000" b="50000"/></am3d:clr>` +
    `<am3d:illuminance n="500000" d="1000000"/></am3d:ambientLight>` +
    `<am3d:ptLight rad="0"><am3d:clr><a:scrgbClr r="100000" g="75000" b="50000"/></am3d:clr>` +
    `<am3d:intensity n="9765625" d="1000000"/><am3d:pos x="21959998" y="70920001" z="16344003"/></am3d:ptLight>` +
    `<am3d:ptLight rad="0"><am3d:clr><a:scrgbClr r="40000" g="60000" b="95000"/></am3d:clr>` +
    `<am3d:intensity n="12250000" d="1000000"/><am3d:pos x="-37964106" y="51130435" z="57631972"/></am3d:ptLight>` +
    `<am3d:ptLight rad="0"><am3d:clr><a:scrgbClr r="86837" g="72700" b="100000"/></am3d:clr>` +
    `<am3d:intensity n="3125000" d="1000000"/><am3d:pos x="-37739122" y="58056624" z="-34769649"/></am3d:ptLight>` +
    `</am3d:model3d></a:graphicData></a:graphic>`;
  const posterGraphic =
    `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
    `<pic:pic><pic:nvPicPr><pic:cNvPr id="1677382972" name="Picture 43"/><pic:cNvPicPr/></pic:nvPicPr>` +
    `<pic:blipFill><a:blip r:embed="rId4"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
    `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="2082800"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic>`;
  return pack(
    `<w:p><w:r><w:rPr><w:noProof/></w:rPr><mc:AlternateContent>` +
    `<mc:Choice Requires="am3d">${anchor(modelGraphic)}</mc:Choice>` +
    `<mc:Fallback>${anchor(posterGraphic)}</mc:Fallback>` +
    `</mc:AlternateContent></w:r></w:p>`,
  );
}

/** The same drawing with its graphic replaced by something unreadable, and
 * NO mc:Fallback to rescue it. */
function malformedModelDocx(graphicBody: string): Uint8Array {
  return pack(
    `<w:p><w:r><mc:AlternateContent><mc:Choice Requires="am3d"><w:drawing>` +
    `<wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="1828800" cy="2082800"/>` +
    `<wp:docPr id="1" name="3D Model 1"/>` +
    `<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/drawing/2017/model3d">` +
    `${graphicBody}</a:graphicData></a:graphic>` +
    `</wp:inline></w:drawing></mc:Choice></mc:AlternateContent></w:r></w:p>`,
  );
}
