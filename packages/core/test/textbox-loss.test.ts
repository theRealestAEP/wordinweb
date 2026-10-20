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

describe("a linked chain of three or more text boxes (GovDocs 032370, 027984)", () => {
  // Three one-line boxes side by side; the first holds the story FIRST /
  // SECOND / THIRD / FOURTH. The third box used to repeat the second's text.
  const cy = 190500; // 20px: one 12pt line
  const xs = [914400, 2914650, 4914900];
  const story = `<wps:txbx id="24"><w:txbxContent>${para("FIRST")}${para("SECOND")}${para("THIRD")}${para("FOURTH")}</w:txbxContent></wps:txbx>`;
  const chainBox = (seq: number) =>
    box({ x: xs[seq], y: 914400, cx: 1809750, cy, story: seq === 0 ? story : `<wps:linkedTxbx id="24" seq="${seq}"/>` });
  // The left edge of each word's text item on page 1, for every copy of it.
  const shown = (order: number[]) => {
    const { result } = layout(`<w:p>${order.map(chainBox).join("")}<w:r><w:t>Anchor</w:t></w:r></w:p>`);
    const page = texts(result.pages[0].items);
    return ["FIRST", "SECOND", "THIRD", "FOURTH"].map((word) => page.filter((t) => t.text === word).map((t) => Math.round(t.x)));
  };
  const left = xs.map((x) => Math.round((x / 914400) * 96));

  it("continues the story in each box, and hides what is left", () => {
    expect(shown([0, 1, 2])).toEqual([[left[0]], [left[1]], [left[2]], []]);
  });

  it("does not depend on the order the anchors are laid out in", () => {
    expect(shown([0, 2, 1])).toEqual([[left[0]], [left[1]], [left[2]], []]);
    expect(shown([1, 0, 2])).toEqual([[left[0]], [left[1]], [left[2]], []]);
  });

  it("continues a chain in a header, which the prescan does not walk", () => {
    const R = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
    const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
    const doc = DocxDocument.load(
      makeDocx({
        "word/document.xml": wrapDocument(
          `<w:p><w:r><w:t>Body</w:t></w:r></w:p><w:sectPr><w:headerReference ${R} w:type="default" r:id="rIdH"/>` +
            `<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/></w:sectPr>`,
        ),
        "word/_rels/document.xml.rels":
          `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
          `<Relationship Id="rIdH" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/></Relationships>`,
        "word/header1.xml": `<?xml version="1.0"?><w:hdr ${W}><w:p>${[0, 1, 2].map(chainBox).join("")}</w:p></w:hdr>`,
      }),
    );
    const page = texts(layoutDocument(doc, { measurer }).pages[0].items);
    const at = (word: string) => page.filter((t) => t.text === word).map((t) => Math.round(t.x));
    expect([at("FIRST"), at("SECOND"), at("THIRD"), at("FOURTH")]).toEqual([[left[0]], [left[1]], [left[2]], []]);
  });

  it("paints each box's part once in the DOM renderer", () => {
    const { doc, result } = layout(`<w:p>${[0, 2, 1].map(chainBox).join("")}<w:r><w:t>Anchor</w:t></w:r></w:p>`);
    const container = document.createElement("div");
    document.body.appendChild(container);
    const rendered = renderToDom(doc, result, container);
    const words = rendered.bindings.map((b) => b.el.textContent);
    expect(words.filter((w) => w === "SECOND")).toHaveLength(1);
    expect(words.filter((w) => w === "THIRD")).toHaveLength(1);
  });
});
