// @vitest-environment jsdom
// Text boxes and anchored objects whose text the layout used to lose. Each
// case comes from a GovDocs1 .doc (doc97-check-a) and is reduced here to the
// .docx shape the .doc reader writes.
import { describe, expect, it } from "vitest";
import { DocxDocument } from "../src/docx.js";
import { layoutDocument } from "../src/layout/engine.js";
import { ApproxMeasurer } from "../src/layout/measure.js";
import type { PageItem } from "../src/layout/types.js";
import { renderToDom } from "../src/render/dom.js";
import { makeDocx, wrapDocument } from "./helpers.js";

const measurer = new ApproxMeasurer();
const WP = 'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"';
const A = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';
const WPS = 'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"';
const LETTER =
  `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>` +
  `<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/></w:sectPr>`;

/** A text box anchored in a run: `story` is the wps:txbx / wps:linkedTxbx element. */
function box(opts: { x: number; y: number; cx: number; cy: number; story: string; from?: string; insets?: string }): string {
  const from = opts.from ?? "page";
  return (
    `<w:r><w:drawing><wp:anchor ${WP} distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">` +
    `<wp:simplePos x="0" y="0"/>` +
    `<wp:positionH relativeFrom="page"><wp:posOffset>${opts.x}</wp:posOffset></wp:positionH>` +
    `<wp:positionV relativeFrom="${from}"><wp:posOffset>${opts.y}</wp:posOffset></wp:positionV>` +
    `<wp:extent cx="${opts.cx}" cy="${opts.cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>` +
    `<wp:docPr id="1" name="Box"/><wp:cNvGraphicFramePr/>` +
    `<a:graphic ${A}><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">` +
    `<wps:wsp ${WPS}><wps:cNvSpPr txBox="1"/><wps:spPr>` +
    `<a:xfrm><a:off x="0" y="0"/><a:ext cx="${opts.cx}" cy="${opts.cy}"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></wps:spPr>` +
    opts.story +
    `<wps:bodyPr wrap="square" ${opts.insets ?? 'lIns="0" tIns="0" rIns="0" bIns="0"'} anchor="t"><a:noAutofit/></wps:bodyPr>` +
    `</wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>`
  );
}

const para = (text: string, sz = 24) => `<w:p><w:r><w:rPr><w:sz w:val="${sz}"/></w:rPr><w:t>${text}</w:t></w:r></w:p>`;

function layout(body: string) {
  const doc = DocxDocument.load(makeDocx({ "word/document.xml": wrapDocument(body + LETTER) }));
  return { doc, result: layoutDocument(doc, { measurer }) };
}

function texts(items: PageItem[]): Extract<PageItem, { kind: "text" }>[] {
  return items.filter((i): i is Extract<PageItem, { kind: "text" }> => i.kind === "text");
}

describe("a text box shorter than one line of its text (GovDocs 042652)", () => {
  // SF-424 instructions: "INSTRUCTION FOR THE SF424" in a page-positioned
  // box 156845 EMU (12.35 pt) high with zero insets and 12 pt text. Word
  // draws the line and clips it at the box bottom; the layout dropped it.
  const cy = 156845;
  const boxPx = (cy / 914400) * 96;

  it("paints the line, clipped to the box", () => {
    const story = `<wps:txbx><w:txbxContent>${para("LABEL")}</w:txbxContent></wps:txbx>`;
    const { result } = layout(`<w:p>${box({ x: 914400, y: 914400, cx: 4998720, cy, story })}<w:r><w:t>Anchor</w:t></w:r></w:p>`);
    const label = texts(result.pages[0].items).find((t) => t.text === "LABEL");
    expect(label).toBeDefined();
    expect(label!.hidden).toBeUndefined();
    expect(label!.lineHeight).toBeGreaterThan(boxPx);
    expect(label!.clip).toBeDefined();
    expect(label!.clip!.top).toBeCloseTo(96, 1);
    expect(label!.clip!.bottom - label!.clip!.top).toBeCloseTo(boxPx, 1);
  });

  it("still hides a second line that starts below the box", () => {
    const story = `<wps:txbx><w:txbxContent>${para("LABEL")}${para("BELOW")}</w:txbxContent></wps:txbx>`;
    const { result } = layout(`<w:p>${box({ x: 914400, y: 914400, cx: 4998720, cy, story })}<w:r><w:t>Anchor</w:t></w:r></w:p>`);
    expect(texts(result.pages[0].items).some((t) => t.text === "BELOW")).toBe(false);
    expect(result.pages[0].hiddenText?.some((t) => t.text === "BELOW")).toBe(true);
  });

  it("paints the line of a box the .doc reader wrote as a chain of one", () => {
    // The .doc reader gives a box an id when its shape names a next shape,
    // even with no box continuing it (all 558 boxes of 042652).
    const story = `<wps:txbx id="468"><w:txbxContent>${para("LABEL")}</w:txbxContent></wps:txbx>`;
    const { result } = layout(`<w:p>${box({ x: 914400, y: 914400, cx: 4998720, cy, story })}<w:r><w:t>Anchor</w:t></w:r></w:p>`);
    const label = texts(result.pages[0].items).find((t) => t.text === "LABEL");
    expect(label?.clip).toBeDefined();
    expect(label!.clip!.bottom - label!.clip!.top).toBeCloseTo(boxPx, 1);
  });

  it("clips the line in the DOM renderer too", () => {
    const story = `<wps:txbx><w:txbxContent>${para("LABEL")}</w:txbxContent></wps:txbx>`;
    const { doc, result } = layout(`<w:p>${box({ x: 914400, y: 914400, cx: 4998720, cy, story })}<w:r><w:t>Anchor</w:t></w:r></w:p>`);
    const container = document.createElement("div");
    document.body.appendChild(container);
    const rendered = renderToDom(doc, result, container);
    const el = rendered.bindings.map((b) => b.el).find((e) => e.textContent === "LABEL");
    expect(el).toBeDefined();
    expect(el!.style.clipPath).toMatch(/^inset\(/);
  });
});
