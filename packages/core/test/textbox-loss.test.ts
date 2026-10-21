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

function layout(body: string, styles?: string) {
  const doc = DocxDocument.load(makeDocx({ "word/document.xml": wrapDocument(body + LETTER), ...(styles ? { "word/styles.xml": styles } : {}) }));
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

  it("paints the first line of a chained box shorter than the line", () => {
    // A chain of one: a box with an id that no wps:linkedTxbx continues.
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

  it("continues the story in outlined boxes", () => {
    // render-eval probe3-linked-textboxes: each box has a filled, outlined
    // shape. The continuation box's outline must not make it a drawing.
    const outlined = (seq: number) =>
      chainBox(seq).replace(
        "<a:noFill/><a:ln><a:noFill/></a:ln>",
        '<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:ln w="9525"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>',
      );
    const { result } = layout(`<w:p>${[0, 1, 2].map(outlined).join("")}<w:r><w:t>Anchor</w:t></w:r></w:p>`);
    const page = texts(result.pages[0].items);
    const at = (word: string) => page.filter((t) => t.text === word).map((t) => Math.round(t.x));
    expect([at("FIRST"), at("SECOND"), at("THIRD"), at("FOURTH")]).toEqual([[left[0]], [left[1]], [left[2]], []]);
  });
});

describe("a compatibility-mode 11 footer that starts with a page-anchored frame (GovDocs 030685)", () => {
  // The first-page footer is a framed logo at a page position, then centred
  // address lines. The whole footer hung down from the footer distance, so
  // the address lines were drawn below the page.
  const R = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
  const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
  const footer =
    `<?xml version="1.0"?><w:ftr ${W}>` +
    `<w:p><w:pPr><w:framePr w:w="540" w:h="360" w:hRule="exact" w:vAnchor="page" w:hAnchor="page" w:x="5945" w:y="13537"/></w:pPr><w:r><w:t>L</w:t></w:r></w:p>` +
    `<w:p/><w:p><w:r><w:t>ADDRESS</w:t></w:r></w:p><w:p><w:r><w:t>EMPLOYER</w:t></w:r></w:p></w:ftr>`;
  const settings =
    `<?xml version="1.0"?><w:settings ${W}><w:compat><w:compatSetting w:name="compatibilityMode" ` +
    `w:uri="http://schemas.microsoft.com/office/word" w:val="11"/></w:compat></w:settings>`;
  const load = (footerXml = footer) =>
    DocxDocument.load(
      makeDocx({
        "word/document.xml": wrapDocument(
          `<w:p><w:r><w:t>Body</w:t></w:r></w:p><w:sectPr><w:footerReference ${R} w:type="default" r:id="rIdF"/>` +
            `<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="432"/></w:sectPr>`,
        ),
        "word/_rels/document.xml.rels":
          `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
          `<Relationship Id="rIdF" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/></Relationships>`,
        "word/footer1.xml": footerXml,
        "word/settings.xml": settings,
      }),
    );

  it("stacks the footer's lines up from the footer distance and puts the frame at its page position", () => {
    const doc = load();
    expect(doc.compatibilityMode).toBe(11);
    const page = texts(layoutDocument(doc, { measurer }).pages[0].items);
    const at = (word: string) => page.find((t) => t.text === word)!;
    // Page 1056px, footer distance 432tw = 28.8px.
    const employer = at("EMPLOYER");
    expect(employer.lineTop + employer.lineHeight).toBeCloseTo(1056 - 28.8, 0);
    expect(at("ADDRESS").lineTop).toBeLessThan(employer.lineTop);
    const logo = at("L");
    expect(logo.lineTop).toBeCloseTo(13537 / 15, 0);
    expect(logo.x).toBeCloseTo(5945 / 15, 0);
  });

  it("stacks paragraphs that share a frame, and aligns the frame by xAlign", () => {
    const framePr = `<w:framePr w:w="1440" w:vAnchor="page" w:hAnchor="margin" w:xAlign="right" w:y="13537"/>`;
    const framed = (text: string) => `<w:p><w:pPr>${framePr}</w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`;
    const doc = load(`<?xml version="1.0"?><w:ftr ${W}>${framed("NAME")}${framed("PLACE")}<w:p><w:r><w:t>FLOW</w:t></w:r></w:p></w:ftr>`);
    const page = texts(layoutDocument(doc, { measurer }).pages[0].items);
    const at = (word: string) => page.find((t) => t.text === word)!;
    expect(at("NAME").lineTop).toBeCloseTo(13537 / 15, 0);
    expect(at("PLACE").lineTop).toBeCloseTo(at("NAME").lineTop + at("NAME").lineHeight, 0);
    // Right-aligned in the margin box: the 1in frame ends at the right margin (720px).
    expect(at("NAME").x).toBeCloseTo(720 - 96, 0);
  });

  it("draws the footer lines inside the page in the DOM renderer", () => {
    const doc = load();
    const container = document.createElement("div");
    document.body.appendChild(container);
    const rendered = renderToDom(doc, layoutDocument(doc, { measurer }), container);
    const el = rendered.root.querySelectorAll<HTMLElement>("*");
    const employer = [...el].find((e) => e.children.length === 0 && e.textContent === "EMPLOYER");
    expect(employer).toBeDefined();
    expect(parseFloat(employer!.style.top)).toBeLessThan(1056 - 28.8);
  });
});

describe("a top-and-bottom art shape anchored above an empty paragraph (break-table br_smallTB_br)", () => {
  // A red rectangle 1in tall, 12pt above its empty anchor paragraph. The
  // body flow ignored an art shape's wrap, so the text after it was drawn
  // over the rectangle.
  const rect =
    `<w:r><w:drawing><wp:anchor ${WP} distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">` +
    `<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>` +
    `<wp:positionV relativeFrom="paragraph"><wp:posOffset>-152400</wp:posOffset></wp:positionV>` +
    `<wp:extent cx="914400" cy="914400"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapTopAndBottom/><wp:docPr id="4" name="R"/><wp:cNvGraphicFramePr/>` +
    `<a:graphic ${A}><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">` +
    `<wps:wsp ${WPS}><wps:cNvSpPr/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></wps:spPr><wps:bodyPr/></wps:wsp>` +
    `</a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>`;
  const pageBreak = `<w:p><w:r><w:br w:type="page"/></w:r></w:p>`;
  const bodies = {
    "without a page break": `${para("AAA")}<w:p>${rect}</w:p>${para("AFTER")}`,
    "after a page break": `${para("AAA")}${pageBreak}<w:p>${rect}</w:p>${para("AFTER")}`,
  };
  for (const [name, body] of Object.entries(bodies)) {
    it(`puts the following text below the rectangle ${name}`, () => {
      const { result } = layout(body);
      const page = result.pages.find((pg) => pg.items.some((i) => i.kind === "path"))!;
      const red = page.items.find((i) => i.kind === "path")!;
      if (red.kind !== "path") throw new Error("no rectangle");
      const after = texts(page.items).find((t) => t.text === "AFTER")!;
      expect(after.lineTop).toBeGreaterThanOrEqual(red.y + red.height);
    });
  }

  it("puts the following text below the rectangle in the DOM renderer", () => {
    const { doc, result } = layout(bodies["without a page break"]);
    const container = document.createElement("div");
    document.body.appendChild(container);
    const rendered = renderToDom(doc, result, container);
    const after = rendered.bindings.map((b) => b.el).find((e) => e.textContent === "AFTER")!;
    // The rectangle runs from 12pt above the anchor line to 1in below that.
    const red = result.pages[0].items.find((i) => i.kind === "path")!;
    if (red.kind !== "path") throw new Error("no rectangle");
    expect(parseFloat(after.style.top)).toBeGreaterThanOrEqual(red.y + red.height);
  });
});

describe("a square-wrapped box in a balanced 3-column section (GovDocs 083839)", () => {
  // The paragraph before the anchor draws the page-positioned box from its
  // lookahead. The balancing pass rolled the page back, but the box stayed
  // marked as drawn, so no pass drew it again and its text was lost.
  const story = `<wps:txbx><w:txbxContent>${para("BOXTEXT")}</w:txbxContent></wps:txbx>`;
  const squareBox = box({ x: 2237105, y: 688340, cx: 3200400, cy: 2743200, story }).replace("<wp:wrapNone/>", '<wp:wrapSquare wrapText="bothSides"/>');
  const threeColumns =
    `<w:p><w:pPr><w:sectPr><w:type w:val="continuous"/><w:pgSz w:w="12240" w:h="15840"/>` +
    `<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/><w:cols w:num="3" w:space="346"/></w:sectPr></w:pPr></w:p>`;
  const load = () =>
    DocxDocument.load(
      makeDocx({
        "word/document.xml": wrapDocument(
          `${para("INTRO")}<w:p>${squareBox}<w:r><w:t>ANCHOR</w:t></w:r></w:p>${para("MORE")}${threeColumns}${para("NEXT")}` +
            LETTER.replace("<w:sectPr>", '<w:sectPr><w:type w:val="continuous"/>'),
        ),
      }),
    );

  it("draws the box", () => {
    const page = layoutDocument(load(), { measurer }).pages[0];
    expect(texts(page.items).some((t) => t.text === "BOXTEXT")).toBe(true);
    expect(page.items.some((i) => i.kind === "drawingHit")).toBe(true);
  });

  it("draws the box's text in the DOM renderer", () => {
    const doc = load();
    const container = document.createElement("div");
    document.body.appendChild(container);
    const rendered = renderToDom(doc, layoutDocument(doc, { measurer }), container);
    expect(rendered.bindings.some((b) => b.el.textContent === "BOXTEXT")).toBe(true);
  });
});

describe("a tall box anchored near the page bottom (GovDocs 026966, 029456, 074645, 071877)", () => {
  // Letter page 1056px, body 96..960px. 40 filler lines of 16.9px end near
  // 770px; the anchor paragraph comes next with a 5.3in (509px) box at
  // offset 0 from it. The box ran past the page bottom, and nothing below
  // the page edge is printed.
  const filler = (n: number) => Array.from({ length: n }, (_, i) => para(`filler${i}`)).join("");
  const lines = Array.from({ length: 25 }, (_, i) => para(`boxline${i}`)).join("");
  const tall = (wrap: string, cy = 4846320) =>
    box({ x: 4572000, y: 0, cx: 2743200, cy, from: "paragraph", story: `<wps:txbx><w:txbxContent>${lines}</w:txbxContent></wps:txbx>` })
      .replace("<wp:wrapNone/>", wrap)
      .replace("<a:noAutofit/>", "");
  const anchored = (wrap: string, cy?: number) => `<w:p>${tall(wrap, cy)}<w:r><w:t>ANCHOR</w:t></w:r></w:p>${para("AFTER")}`;

  const shown = (result: ReturnType<typeof layoutDocument>, text: string) =>
    result.pages.flatMap((pg, n) => texts(pg.items).filter((t) => t.text === text).map((t) => ({ page: n + 1, t, height: pg.height })));

  for (const wrap of ['<wp:wrapSquare wrapText="bothSides"/>', "<wp:wrapTopAndBottom/>"]) {
    it(`moves the anchor paragraph and its box to the next page (${wrap.slice(4, 18)})`, () => {
      const { result } = layout(filler(40) + anchored(wrap));
      const [anchor] = shown(result, "ANCHOR");
      expect(anchor.page).toBe(2);
      for (let i = 0; i < 25; i++) {
        const [line] = shown(result, `boxline${i}`);
        expect(line.page).toBe(2);
        expect(line.t.lineTop + line.t.lineHeight).toBeLessThanOrEqual(line.height);
      }
    });
  }

  it("keeps a box taller than the room below a page-top anchor on the page, moved up", () => {
    // 11in box: taller than the page; it sits at the page top.
    const { result } = layout(anchored('<wp:wrapSquare wrapText="bothSides"/>', 10058400));
    const hit = result.pages[0].items.find((i) => i.kind === "drawingHit")!;
    if (hit.kind !== "drawingHit") throw new Error("no box");
    expect(hit.y).toBe(0);
  });

  it("starts the next page at its top after a top-and-bottom box that pushed a line off the page", () => {
    // The box fits on page 1 below its anchor but leaves no room under it
    // for AFTER, which used to land on page 2 below a blank band as tall as
    // the push.
    // doc97-check-c box30: the anchor at about 603px, the 349px box 70px
    // below it, so the box ends at about 1022px, below the body bottom.
    const box30 = tall("<wp:wrapTopAndBottom/>", 3324225).replace(
      '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset>',
      '<wp:positionV relativeFrom="paragraph"><wp:posOffset>669290</wp:posOffset>',
    );
    const { result } = layout(filler(30) + `<w:p>${box30}<w:r><w:t>ANCHOR</w:t></w:r></w:p>${para("AFTER")}`);
    const [after] = shown(result, "AFTER");
    expect(after.page).toBe(2);
    // Page 2 starts at the body top, and AFTER sits in the flow there.
    const tops = texts(result.pages[1].items).map((t) => t.lineTop);
    expect(Math.min(...tops)).toBeCloseTo(96, 0);
    expect(after.t.lineTop).toBeLessThan(96 + 2 * after.t.lineHeight);
  });

  it("draws the moved box inside the page in the DOM renderer", () => {
    const { doc, result } = layout(filler(40) + anchored('<wp:wrapSquare wrapText="bothSides"/>'));
    const container = document.createElement("div");
    document.body.appendChild(container);
    const rendered = renderToDom(doc, result, container);
    const last = rendered.bindings.map((b) => b.el).find((e) => e.textContent === "boxline24")!;
    expect(parseFloat(last.style.top)).toBeLessThan(1056);
  });
});

describe("a blank paragraph's mark size in a text box (GovDocs 092399)", () => {
  const styles =
    `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
    `<w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="36"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>`;
  const paragraphs = para("Line A", 16) + `<w:p><w:pPr><w:rPr><w:sz w:val="12"/></w:rPr></w:pPr></w:p>` + para("Line B", 16);
  const story = `<wps:txbx><w:txbxContent>${paragraphs}</w:txbxContent></wps:txbx>`;
  const body = `<w:p>${box({ x: 914400, y: 914400, cx: 3000000, cy: 1500000, story })}</w:p>`;
  const gap = (result: ReturnType<typeof layoutDocument>) => {
    const lines = texts(result.pages[0].items);
    return lines.find((t) => t.src?.t?.text === "Line B")!.lineTop - lines.find((t) => t.src?.t?.text === "Line A")!.lineTop;
  };

  it("advances the same distance as the body", () => {
    const bodyGap = gap(layout(paragraphs, styles).result);
    const boxGap = gap(layout(body, styles).result);
    expect(Math.abs(boxGap - bodyGap)).toBeLessThanOrEqual(0.5);
  });

  it("draws Line B at its laid-out line top in the DOM", () => {
    const { doc, result } = layout(body, styles);
    const container = document.createElement("div");
    const rendered = renderToDom(doc, result, container);
    const line = texts(result.pages[0].items).find((t) => t.src?.t?.text === "Line B")!;
    const el = rendered.bindings.find((b) => b.item.src?.t?.text === "Line B")!.el;
    expect(Math.abs(parseFloat(el.style.top) - line.lineTop)).toBeLessThanOrEqual(0.5);
  });
});

describe("a paragraph pushed off the page by a top-and-bottom float (review of 29b557c)", () => {
  // A page-positioned rectangle covers y = 120..955px of the 96..960px body.
  // The next paragraph's first line goes below it, past the body bottom, so
  // the whole paragraph moves to page 2. The break plan used to keep the
  // float's push for the moved line and break again after it.
  const rect =
    `<w:r><w:drawing><wp:anchor ${WP} distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">` +
    `<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>` +
    `<wp:positionV relativeFrom="page"><wp:posOffset>1143000</wp:posOffset></wp:positionV>` +
    `<wp:extent cx="914400" cy="7953375"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapTopAndBottom/><wp:docPr id="4" name="R"/><wp:cNvGraphicFramePr/>` +
    `<a:graphic ${A}><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">` +
    `<wps:wsp ${WPS}><wps:cNvSpPr/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="7953375"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></wps:spPr><wps:bodyPr/></wps:wsp>` +
    `</a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>`;
  const tenLines =
    `<w:p><w:pPr><w:widowControl w:val="0"/></w:pPr>` +
    Array.from({ length: 10 }, (_, i) => `<w:r><w:t>row${i}</w:t>${i < 9 ? "<w:br/>" : ""}</w:r>`).join("") +
    `</w:p>`;

  it("keeps the moved paragraph together on page 2", () => {
    const { result } = layout(`<w:p>${rect}</w:p>${tenLines}`);
    const rows = result.pages.flatMap((pg, n) => texts(pg.items).filter((t) => /^row\d$/.test(t.text)).map((t) => ({ page: n + 1, top: t.lineTop })));
    expect(rows).toHaveLength(10);
    expect(rows.every((r) => r.page === 2)).toBe(true);
    expect(rows[0].top).toBeCloseTo(96, 0);
    expect(result.pages.length).toBe(2);
  });
});
