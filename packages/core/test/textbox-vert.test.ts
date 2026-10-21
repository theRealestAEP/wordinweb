// @vitest-environment jsdom
// Vertical figure captions reduced from GovDocs 045416 (m21.docx).
import { describe, expect, it } from "vitest";
import { DocxDocument } from "../src/docx.js";
import { layoutDocument } from "../src/layout/engine.js";
import { ApproxMeasurer } from "../src/layout/measure.js";
import type { PageItem } from "../src/layout/types.js";
import { renderToDom } from "../src/render/dom.js";
import { makeDocx, wrapDocument } from "./helpers.js";

const measurer = new ApproxMeasurer();
const caption = "Figure 7. Power versus true underlying effect";
const para = (text: string) => `<w:p><w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:t>${text}</w:t></w:r></w:p>`;
const LETTER =
  `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>` +
  `<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/></w:sectPr>`;

function box(opts: { vert?: string; rotation?: number; insets?: string; story?: string } = {}): string {
  return (
    `<w:r><w:drawing><wp:anchor xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ` +
    `distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">` +
    `<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="page"><wp:posOffset>914400</wp:posOffset></wp:positionH>` +
    `<wp:positionV relativeFrom="page"><wp:posOffset>914400</wp:posOffset></wp:positionV>` +
    `<wp:extent cx="457200" cy="3657600"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>` +
    `<wp:docPr id="1" name="Box"/><wp:cNvGraphicFramePr/>` +
    `<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
    `<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">` +
    `<wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:cNvSpPr txBox="1"/><wps:spPr>` +
    `<a:xfrm rot="${(opts.rotation ?? 0) * 60000}"><a:off x="0" y="0"/><a:ext cx="457200" cy="3657600"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></wps:spPr>` +
    `<wps:txbx><w:txbxContent>${opts.story ?? para(caption)}</w:txbxContent></wps:txbx>` +
    `<wps:bodyPr ${opts.vert ? `vert="${opts.vert}"` : ""} wrap="square" anchor="t" ` +
    `${opts.insets ?? 'lIns="0" tIns="0" rIns="0" bIns="0"'}><a:noAutofit/></wps:bodyPr>` +
    `</wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>`
  );
}

function layout(drawing: string) {
  const doc = DocxDocument.load(makeDocx({ "word/document.xml": wrapDocument(`<w:p>${drawing}</w:p>${LETTER}`) }));
  return { doc, result: layoutDocument(doc, { measurer }) };
}

function texts(items: PageItem[]) {
  return items.filter((it): it is Extract<PageItem, { kind: "text" }> => it.kind === "text");
}

function expectVertical(items: PageItem[], deg: number, bounds = { left: 96, top: 96, right: 144, bottom: 480 }) {
  const ts = texts(items);
  expect(ts.map((t) => t.text).join("")).toBe(caption);
  expect(new Set(ts.map((t) => t.lineTop)).size).toBe(1);
  const centers = ts.map((t) => [t.x + t.rotate!.ox, (t.glyphTop ?? t.lineTop) + t.rotate!.oy]);
  for (let i = 0; i < ts.length; i++) {
    const t = ts[i];
    expect(t.rotate?.deg).toBe(deg);
    expect(centers[i][0]).toBeCloseTo(centers[0][0], 6);
    expect(centers[i][1]).toBeCloseTo(centers[0][1], 6);
    const [cx, cy] = centers[i];
    const a = deg * Math.PI / 180;
    for (const x of [t.x, t.x + t.width]) {
      for (const y of [t.lineTop, t.lineTop + t.lineHeight]) {
        const rx = cx + Math.cos(a) * (x - cx) - Math.sin(a) * (y - cy);
        const ry = cy + Math.sin(a) * (x - cx) + Math.cos(a) * (y - cy);
        expect(rx).toBeGreaterThanOrEqual(bounds.left - 0.01);
        expect(rx).toBeLessThanOrEqual(bounds.right + 0.01);
        expect(ry).toBeGreaterThanOrEqual(bounds.top - 0.01);
        expect(ry).toBeLessThanOrEqual(bounds.bottom + 0.01);
      }
    }
  }
}

describe("vertical anchored text boxes (GovDocs 045416)", () => {
  it("lays out vert270 as one bottom-to-top line inside the box", () => {
    const { result } = layout(box({ vert: "vert270" }));
    expectVertical(result.pages[0].items, -90);
  });

  it.each(["vert", "eaVert", "wordArtVert", "wordArtVertRtl", "mongolianVert"])("draws %s top-to-bottom", (vert) => {
    const { result } = layout(box({ vert }));
    expectVertical(result.pages[0].items, 90);
  });

  it("reads bottom-to-top layout-flow from a VML textbox", () => {
    const { result } = layout(
      `<w:r><w:pict><v:shape xmlns:v="urn:schemas-microsoft-com:vml" ` +
      `style="position:absolute;margin-left:72pt;margin-top:72pt;width:36pt;height:288pt;` +
      `mso-position-horizontal-relative:page;mso-position-vertical-relative:page" filled="f" stroked="f">` +
      `<v:textbox inset="0,0,0,0" style="layout-flow:vertical;mso-layout-flow-alt:bottom-to-top">` +
      `<w:txbxContent>${para(caption)}</w:txbxContent></v:textbox></v:shape></w:pict></w:r>`,
    );
    expectVertical(result.pages[0].items, -90);
  });

  it.each([undefined, "horz"])("keeps horizontal text unrotated (%s)", (vert) => {
    const { result } = layout(box({ vert }));
    const ts = texts(result.pages[0].items);
    expect(ts.map((t) => t.text).join("")).toBe(caption);
    expect(ts.every((t) => t.rotate === undefined)).toBe(true);
    expect(new Set(ts.map((t) => t.lineTop)).size).toBeGreaterThan(1);
  });

  it.each(["vert", "vert270"])("keeps asymmetric insets on the box axes (%s)", (vert) => {
    const { result } = layout(box({ vert, insets: 'lIns="19050" tIns="38100" rIns="57150" bIns="76200"' }));
    expectVertical(result.pages[0].items, vert === "vert" ? 90 : -90, { left: 98, top: 100, right: 138, bottom: 472 });
  });

  it("adds the shape rotation to the text direction", () => {
    const { result } = layout(box({ vert: "vert270", rotation: 30 }));
    expect(texts(result.pages[0].items).every((t) => t.rotate?.deg === -60)).toBe(true);
    const hit = result.pages[0].items.find((it) => it.kind === "drawingHit")!;
    expect(hit.rotate?.deg).toBe(30);
  });

  it("hides overflow across the narrow axis and reports the text fit", () => {
    const { result } = layout(box({ vert: "vert270", story: ["FIRST", "SECOND", "THIRD", "FOURTH"].map(para).join("") }));
    const page = result.pages[0];
    // Lines past the first that cross the box edge are hidden whole, as in a
    // horizontal box.
    expect(texts(page.items).some((t) => t.text === "FIRST")).toBe(true);
    expect(page.hiddenText?.some((t) => t.text === "FOURTH")).toBe(true);
    const hit = page.items.find((it) => it.kind === "drawingHit");
    expect(hit?.kind === "drawingHit" && hit.textFit?.overflow).toBe(true);
  });

  it("draws the vert270 caption with a DOM rotate transform", () => {
    const { doc, result } = layout(box({ vert: "vert270" }));
    const container = document.createElement("div");
    document.body.appendChild(container);
    const rendered = renderToDom(doc, result, container);
    const ts = rendered.bindings.map((b) => b.el).filter((el) => el.textContent?.includes("Figure"));
    expect(ts.length).toBeGreaterThan(0);
    expect(ts.every((el) => el.style.transform.includes("rotate(-90deg)"))).toBe(true);
  });
});
