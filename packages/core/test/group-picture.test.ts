// @vitest-environment jsdom
import { describe, expect, it } from "vitest";
import { DocxDocument } from "../src/docx.js";
import { layoutDocument } from "../src/layout/engine.js";
import { ApproxMeasurer } from "../src/layout/measure.js";
import { renderToDom } from "../src/render/dom.js";
import { makeDocxWithMedia, p, wrapDocument } from "./helpers.js";

// A Word 97 figure is a drawing group that holds a picture and a caption text
// box. The .doc reader writes it as wpg:wgp (pic:pic + wps:wsp); older .docx
// files hold the same figure as a VML v:group. govdocs 071539: every figure
// printed its caption only, and the picture was lost.

const PNG = Uint8Array.from(
  atob("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=="),
  (c) => c.charCodeAt(0),
);

const RELS = `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdImg" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
</Relationships>`;

const NS =
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"' +
  ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"' +
  ' xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"' +
  ' xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"' +
  ' xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"' +
  ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"' +
  ' xmlns:v="urn:schemas-microsoft-com:vml"' +
  ' xmlns:w10="urn:schemas-microsoft-com:office:word"';

// 6in x 3.5in group: the picture fills the top 3in, the caption box the
// 0.5in below it. Both in EMU, in the group's child space.
const GROUP_CHILDREN =
  "<pic:pic><pic:nvPicPr><pic:cNvPr id=\"0\" name=\"Picture\"/><pic:cNvPicPr/></pic:nvPicPr>" +
  '<pic:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>' +
  '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="5486400" cy="2743200"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>' +
  '<wps:wsp><wps:cNvPr id="2" name="Caption"/><wps:cNvSpPr txBox="1"/>' +
  '<wps:spPr><a:xfrm><a:off x="0" y="2743200"/><a:ext cx="5486400" cy="457200"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>' +
  '<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></wps:spPr>' +
  "<wps:txbx><w:txbxContent><w:p><w:r><w:t>Caption</w:t></w:r></w:p></w:txbxContent></wps:txbx>" +
  '<wps:bodyPr lIns="0" tIns="0" rIns="0" bIns="0"/></wps:wsp>';

const GROUP_XFRM =
  '<wpg:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="5486400" cy="3200400"/>' +
  '<a:chOff x="0" y="0"/><a:chExt cx="5486400" cy="3200400"/></a:xfrm></wpg:grpSpPr>';

const GRAPHIC =
  '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup">' +
  `<wpg:wgp>${GROUP_XFRM}${GROUP_CHILDREN}</wpg:wgp></a:graphicData></a:graphic>`;

const ANCHORED_WGP =
  `<w:p><w:r><w:t>Before</w:t></w:r><w:r><w:drawing ${NS}>` +
  '<wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">' +
  '<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>' +
  '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>' +
  '<wp:extent cx="5486400" cy="3200400"/><wp:wrapTopAndBottom/><wp:docPr id="1" name="Group 1"/>' +
  `${GRAPHIC}</wp:anchor></w:drawing></w:r></w:p>`;

const INLINE_WGP =
  `<w:p><w:r><w:drawing ${NS}><wp:inline><wp:extent cx="5486400" cy="3200400"/><wp:docPr id="1" name="Group 1"/>` +
  `${GRAPHIC}</wp:inline></w:drawing></w:r></w:p>`;

// The same figure as Word writes it in VML: group coordinates 4320 x 2520
// over a 432pt x 252pt box (10 units a point).
const VML_GROUP =
  `<w:p><w:r><w:t>Before</w:t></w:r><w:r><w:pict ${NS}>` +
  '<v:group style="position:absolute;margin-left:0;margin-top:0;width:432pt;height:252pt;z-index:1;' +
  'mso-position-horizontal-relative:text;mso-position-vertical-relative:text" coordorigin="0,0" coordsize="4320,2520">' +
  '<v:shape style="position:absolute;left:0;top:0;width:4320;height:2160" type="#_x0000_t75"><v:imagedata r:id="rIdImg" o:title=""/></v:shape>' +
  '<v:shape type="#_x0000_t202" style="position:absolute;left:0;top:2160;width:4320;height:360" stroked="f">' +
  '<v:textbox inset="0,0,0,0"><w:txbxContent><w:p><w:r><w:t>Caption</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape>' +
  '<w10:wrap type="topAndBottom"/></v:group></w:pict></w:r></w:p>';

function load(body: string) {
  const doc = DocxDocument.load(
    makeDocxWithMedia(
      { "word/document.xml": wrapDocument(body), "word/_rels/document.xml.rels": RELS },
      { "word/media/image1.png": PNG },
    ),
  );
  return { doc, result: layoutDocument(doc, { measurer: new ApproxMeasurer() }) };
}

function figure(result: ReturnType<typeof layoutDocument>) {
  const items = result.pages[0].items;
  const image = items.find((i) => i.kind === "image" && i.part === "word/media/image1.png");
  const caption = items.find((i) => i.kind === "text" && i.text === "Caption");
  const after = items.find((i) => i.kind === "text" && i.text === "After");
  if (image?.kind !== "image" || caption?.kind !== "text" || after?.kind !== "text") {
    throw new Error("the picture, the caption and the next paragraph must all be laid out");
  }
  return { image, caption, after };
}

describe("a drawing group of a picture and a caption text box", () => {
  for (const [name, body] of [
    ["wpg:wgp", ANCHORED_WGP],
    ["VML v:group", VML_GROUP],
  ] as const) {
    it(`draws the picture of a ${name} at its place in the group, and the caption below it`, () => {
      const { result } = load(body + p("After"));
      const { image, caption, after } = figure(result);
      // 6in x 3in at 96 px an inch.
      expect(image.width).toBeCloseTo(576, 0);
      expect(image.height).toBeCloseTo(288, 0);
      // The caption box starts where the picture ends.
      expect(caption.lineTop).toBeGreaterThanOrEqual(image.y + image.height - 0.5);
      // wrapTopAndBottom: the next paragraph starts below the whole group (3.5in).
      expect(after.lineTop).toBeGreaterThanOrEqual(image.y + 336 - 0.5);
    });
  }

  it("draws the picture of an inline wpg:wgp group", () => {
    const { result } = load(INLINE_WGP + p("After"));
    const { image, caption } = figure(result);
    expect(image.height).toBeCloseTo(288, 0);
    expect(caption.lineTop).toBeGreaterThanOrEqual(image.y + image.height - 0.5);
  });

  it("places a character-relative wpg:wgp at the drawing run after leading text", () => {
    const body = ANCHORED_WGP
      .replace("Before", "Lead words before the anchor ")
      .replace('relativeFrom="column"', 'relativeFrom="character"')
      .replace('relativeFrom="paragraph"', 'relativeFrom="line"')
      .replace("<wp:wrapTopAndBottom/>", "<wp:wrapNone/>");
    const { result } = load(body + p("After"));
    const { image } = figure(result);
    const lead = result.pages[0].items.find((i) => i.kind === "text" && i.text.startsWith("Lead"));
    if (lead?.kind !== "text") throw new Error("the leading text must be laid out");
    expect(image.x).toBeGreaterThan(lead.x);
  });

  it("draws a VML v:group with space-separated coordinate vectors", () => {
    const body = VML_GROUP
      .replace('coordorigin="0,0"', 'coordorigin="0 0"')
      .replace('coordsize="4320,2520"', 'coordsize="4320 2520"');
    const { result } = load(body + p("After"));
    const { image, caption } = figure(result);
    expect(image.width).toBeCloseTo(576, 0);
    expect(image.height).toBeCloseTo(288, 0);
    expect(caption.lineTop).toBeGreaterThanOrEqual(image.y + image.height - 0.5);
  });

  it("draws the picture of a nested VML v:group in its own coordinate space", () => {
    const body = VML_GROUP.replace(
      '<v:shape style="position:absolute;left:0;top:0;width:4320;height:2160" type="#_x0000_t75"><v:imagedata r:id="rIdImg" o:title=""/></v:shape>',
      '<v:group style="position:absolute;left:0;top:0;width:4320;height:2160" coordorigin="100,100" coordsize="200,100">' +
      '<v:shape style="position:absolute;left:100;top:100;width:200;height:100" type="#_x0000_t75"><v:imagedata r:id="rIdImg" o:title=""/></v:shape></v:group>',
    );
    const { result } = load(body + p("After"));
    const { image, caption, after } = figure(result);
    expect(image.width).toBeCloseTo(576, 0);
    expect(image.height).toBeCloseTo(288, 0);
    expect(caption.lineTop).toBeGreaterThanOrEqual(image.y + image.height - 0.5);
    expect(after.lineTop).toBeGreaterThanOrEqual(image.y + 336 - 0.5);
  });

  it("draws a v:image picture in a VML v:group", () => {
    const body = VML_GROUP.replace(
      '<v:shape style="position:absolute;left:0;top:0;width:4320;height:2160" type="#_x0000_t75"><v:imagedata r:id="rIdImg" o:title=""/></v:shape>',
      '<v:image style="position:absolute;left:0;top:0;width:4320;height:2160" r:id="rIdImg"/>',
    );
    const { result } = load(body + p("After"));
    const { image } = figure(result);
    expect(image.width).toBeCloseTo(576, 0);
    expect(image.height).toBeCloseTo(288, 0);
  });

  it("lets text flow behind a floating VML v:group with no wrap child", () => {
    const body = VML_GROUP.replace('<w10:wrap type="topAndBottom"/>', "");
    const { result } = load(body + p("After"));
    const { image, after } = figure(result);
    expect(after.lineTop).toBeLessThan(image.y + 336);
  });

  it("uses the VML group's wrapNone when its picture shape has square wrap", () => {
    const body = VML_GROUP
      .replace('<v:imagedata r:id="rIdImg" o:title=""/>', '<v:imagedata r:id="rIdImg" o:title=""/><w10:wrap type="square"/>')
      .replace('<w10:wrap type="topAndBottom"/>', '<w10:wrap type="none"/>');
    const { result } = load(body + p("After"));
    const { image, after } = figure(result);
    expect(after.lineTop).toBeLessThan(image.y + 336);
  });

  it("draws a filled VML rect with empty text below the group picture", () => {
    const body = VML_GROUP.replace(
      '<v:shape type="#_x0000_t202" style="position:absolute;left:0;top:2160;width:4320;height:360" stroked="f">' +
      '<v:textbox inset="0,0,0,0"><w:txbxContent><w:p><w:r><w:t>Caption</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape>',
      '<v:rect fillcolor="#FF0000" style="position:absolute;left:0;top:2160;width:4320;height:360"/>',
    );
    const { result } = load(body + p("After"));
    const items = result.pages[0].items;
    const image = items.find((i) => i.kind === "image" && i.part === "word/media/image1.png");
    const rect = items.find((i) => i.kind === "rect" && i.fill === "#FF0000");
    if (image?.kind !== "image" || rect?.kind !== "rect") {
      throw new Error("the picture and the red-filled rect must both be laid out");
    }
    expect(rect.x).toBeCloseTo(image.x, 0);
    expect(rect.y).toBeCloseTo(image.y + 288, 0);
    expect(rect.width).toBeCloseTo(576, 0);
    expect(rect.height).toBeCloseTo(48, 0);
  });

  it("draws the picture and caption of an inline VML v:group", () => {
    const body = VML_GROUP
      .replace("position:absolute;margin-left", "margin-left")
      .replace('<w10:wrap type="topAndBottom"/>', "");
    const { result } = load(body + p("After"));
    const { image, caption } = figure(result);
    expect(image.width).toBeCloseTo(576, 0);
    expect(image.height).toBeCloseTo(288, 0);
    expect(caption.lineTop).toBeGreaterThanOrEqual(image.y + image.height - 0.5);
  });

  it("lays out a wpg:wgp with only a text box as a lone floating text box", () => {
    const body = ANCHORED_WGP.replace(
      GROUP_CHILDREN,
      GROUP_CHILDREN.slice(GROUP_CHILDREN.indexOf("<wps:wsp>")),
    );
    const { result } = load(body + p("After"));
    const items = result.pages[0].items;
    expect(items.some((i) => i.kind === "text" && i.text === "Caption")).toBe(true);
    expect(items.some((i) => i.kind === "image")).toBe(false);
  });

  it("puts the picture and the caption in the DOM page", () => {
    // jsdom has no blob URLs; the renderer hands each picture one.
    URL.createObjectURL ??= () => "blob:test";
    const { doc, result } = load(ANCHORED_WGP + p("After"));
    const container = document.createElement("div");
    document.body.appendChild(container);
    renderToDom(doc, result, container);
    expect(container.querySelectorAll("img").length).toBe(1);
    expect(container.textContent).toContain("Caption");
  });
});
