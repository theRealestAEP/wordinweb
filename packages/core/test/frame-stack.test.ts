import { describe, expect, it } from "vitest";
import { DocxDocument } from "../src/docx.js";
import { layoutDocument } from "../src/layout/engine.js";
import { ApproxMeasurer } from "../src/layout/measure.js";
import type { ImageItem, LineEdgeItem, PageItem, RectItem, TextItem } from "../src/layout/types.js";
import { makeDocx, makeDocxWithMedia, wrapDocument } from "./helpers.js";

const PNG = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);

// Adjacent paragraphs with equal w:framePr are ONE frame (ECMA-376
// 17.3.1.11): Word stacks them one under another inside it. Each paragraph
// used to be drawn at the frame's own position, so their lines overlapped
// (word95 r3/r4, and the 20-paragraph bordered box of GovDocs 001939).
// Each frame paragraph here is one word: the engine emits a text item per word.

const measurer = new ApproxMeasurer();
const TW = 1 / 15; // px per twip

const SECTION =
  `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>` +
  `<w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="720" w:footer="720"/></w:sectPr>`;

const FRAME = `<w:framePr w:w="4000" w:hSpace="187" w:wrap="around" w:vAnchor="text" w:hAnchor="page" w:x="1494" w:y="200"/>`;
const FRAME_X = 1494 * TW;
const FRAME_W = 4000 * TW;

const BODY =
  `<w:p><w:r><w:t xml:space="preserve">Body text after the frame, long enough to wrap beside it on several lines of the page body here, and then some more words so that it keeps going below the frame as well.</w:t></w:r></w:p>`;

function framed(text: string, frame = FRAME, extraPPr = ""): string {
  return `<w:p><w:pPr>${frame}${extraPPr}<w:spacing w:after="0"/></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`;
}

function layoutBody(body: string, parts: Record<string, string> = {}) {
  const doc = DocxDocument.load(makeDocx({ "word/document.xml": wrapDocument(body), ...parts }));
  return layoutDocument(doc, { measurer }).pages[0].items;
}

function texts(items: PageItem[]): TextItem[] {
  return items.filter((i): i is TextItem => i.kind === "text" && i.text.trim().length > 0);
}

function line(items: PageItem[], text: string): TextItem {
  const item = texts(items).find((i) => i.text === text);
  if (!item) throw new Error(`no text item "${text}"`);
  return item;
}

describe("a frame of adjacent paragraphs with equal framePr", () => {
  it("stacks the paragraphs in order, one under another, at the frame's position", () => {
    const items = layoutBody(
      `<w:p><w:r><w:t>Intro</w:t></w:r></w:p>` +
        framed("FrameOne") +
        framed("FrameTwo") +
        framed("FrameThree") +
        BODY +
        SECTION,
    );
    const lines = ["FrameOne", "FrameTwo", "FrameThree"].map((t) => line(items, t));
    for (const l of lines) expect(l.x).toBeCloseTo(FRAME_X, 1);
    for (let k = 1; k < lines.length; k++) {
      // No vertical overlap, and the source order is kept.
      expect(lines[k].lineTop).toBeGreaterThanOrEqual(lines[k - 1].lineTop + lines[k - 1].lineHeight - 0.01);
    }
    // The frame rides with the text: its top is the anchor paragraph's top + y.
    const body = line(items, "Body");
    expect(lines[0].lineTop).toBeCloseTo(body.lineTop + 200 * TW, 0);
  });

  it("wraps the body text around the whole group, not only its first paragraph", () => {
    const items = layoutBody(
      framed("FrameOne") + framed("FrameTwo") + framed("FrameThree") + framed("FrameFour") + BODY.repeat(3) + SECTION,
    );
    const frameTop = line(items, "FrameOne").lineTop;
    const last = line(items, "FrameFour");
    const frameBottom = last.lineTop + last.lineHeight;
    const channelX = FRAME_X + FRAME_W + 187 * TW;
    const frameLines = new Set(["FrameOne", "FrameTwo", "FrameThree", "FrameFour"]);
    const bodyLines = texts(items).filter((i) => !frameLines.has(i.text));
    // Every body line beside the group starts in the channel right of it.
    const beside = bodyLines.filter((i) => i.lineTop + i.lineHeight > frameTop + 0.5 && i.lineTop < frameBottom - 0.5);
    // Four frame lines tall: the body wraps beside it on three lines or more.
    expect(new Set(beside.map((i) => Math.round(i.lineTop))).size).toBeGreaterThanOrEqual(3);
    for (const i of beside) expect(i.x).toBeGreaterThanOrEqual(channelX - 0.5);
    // Below the group the text returns to the margin.
    const below = bodyLines.filter((i) => i.lineTop >= frameBottom - 0.5);
    expect(below.length).toBeGreaterThan(0);
    expect(Math.min(...below.map((i) => i.x))).toBeCloseTo(1800 * TW, 0);
  });

  it("keeps frames with different properties apart", () => {
    const other = FRAME.replace('w:y="200"', 'w:y="2000"');
    const items = layoutBody(framed("FirstFrame") + framed("SecondFrame", other) + BODY + SECTION);
    const first = line(items, "FirstFrame");
    const second = line(items, "SecondFrame");
    // Each paragraph is its own frame at its own y, so they do not stack.
    expect(second.lineTop - first.lineTop).toBeCloseTo((2000 - 200) * TW, 0);
  });

  it("draws one border box around the group and stretches it to an exact frame height", () => {
    const bdr =
      `<w:pBdr>` +
      ["top", "left", "bottom", "right"].map((s) => `<w:${s} w:val="single" w:sz="6" w:space="1" w:color="FF0000"/>`).join("") +
      `</w:pBdr>`;
    const exact = FRAME.replace("<w:framePr ", '<w:framePr w:h="3000" w:hRule="exact" ');
    const items = layoutBody(
      framed("BoxedOne", exact, bdr) + framed("BoxedTwo", exact, bdr) + framed("BoxedThree", exact, bdr) + BODY + SECTION,
    );
    const red = items.filter((i) => i.kind === "edge" && i.border.color.toUpperCase().includes("FF0000"));
    const horizontal = red.filter((i) => i.kind === "edge" && Math.abs(i.y1 - i.y2) < 0.01);
    // Interior boundaries of the merged box paint no rule: one top, one bottom.
    expect(horizontal).toHaveLength(2);
    const ys = horizontal.map((i) => (i.kind === "edge" ? i.y1 : 0)).sort((a, b) => a - b);
    const top = line(items, "BoxedOne").lineTop;
    expect(ys[0]).toBeLessThan(top);
    // The exact height (3000tw = 200px) holds the whole group: the bottom
    // rule moves down to the frame's bottom, below all three lines.
    expect(ys[1] - ys[0]).toBeGreaterThan(3000 * TW - 12);
    const lastLine = line(items, "BoxedThree");
    expect(ys[1]).toBeGreaterThan(lastLine.lineTop + lastLine.lineHeight);
    // Body text wraps beside the frame for its full exact height.
    const bodyBeside = texts(items).filter(
      (i) => !i.text.startsWith("Boxed") && i.lineTop > ys[0] && i.lineTop + i.lineHeight < ys[1],
    );
    for (const i of bodyBeside) expect(i.x).toBeGreaterThan(FRAME_X + FRAME_W);
  });

  it("lays a table whose cell paragraphs carry the frame inside the frame, under its paragraph", () => {
    const cell = (t: string) =>
      `<w:tc><w:tcPr><w:tcW w:w="1900" w:type="dxa"/></w:tcPr>${framed(t)}</w:tc>`;
    const table =
      `<w:tbl><w:tblPr><w:tblW w:w="3800" w:type="dxa"/></w:tblPr>` +
      `<w:tblGrid><w:gridCol w:w="1900"/><w:gridCol w:w="1900"/></w:tblGrid>` +
      `<w:tr>${cell("CellA")}${cell("CellB")}</w:tr></w:tbl>`;
    const items = layoutBody(framed("Caption") + table + framed("Note") + BODY + SECTION);
    const caption = line(items, "Caption");
    const a = line(items, "CellA");
    const b = line(items, "CellB");
    const note = line(items, "Note");
    expect(a.lineTop).toBeGreaterThanOrEqual(caption.lineTop + caption.lineHeight - 0.01);
    expect(b.lineTop).toBeCloseTo(a.lineTop, 1);
    expect(note.lineTop).toBeGreaterThanOrEqual(a.lineTop + a.lineHeight - 0.01);
    // The table sits inside the frame's box, not at the body margin.
    expect(a.x).toBeCloseTo(FRAME_X, -1);
    expect(b.x).toBeLessThan(FRAME_X + FRAME_W);
    // The table takes no room in the body flow: the body starts at the
    // frame's anchor, beside the frame.
    const body = line(items, "Body");
    expect(body.lineTop).toBeCloseTo(caption.lineTop - 200 * TW, 0);
    expect(body.x).toBeGreaterThan(FRAME_X + FRAME_W);
  });

  it("sizes a frame of no w:w to its widest line and wraps the body beside it", () => {
    const auto = FRAME.replace('w:w="4000" ', "");
    const items = layoutBody(framed("Short", auto) + framed("Widest", auto) + BODY.repeat(2) + SECTION);
    const short = line(items, "Short");
    const widest = line(items, "Widest");
    expect(widest.lineTop).toBeGreaterThanOrEqual(short.lineTop + short.lineHeight - 0.01);
    expect(short.x).toBeCloseTo(FRAME_X, 1);
    // The body's channel starts one hSpace past the widest frame line.
    const body = line(items, "Body");
    expect(body.x).toBeCloseTo(widest.x + widest.width + 187 * TW, 0);
  });

  it("stacks a page-anchored frame at its page position (word95-check-b frame-three-paragraphs)", () => {
    const page = `<w:framePr w:w="2381" w:wrap="around" w:vAnchor="page" w:hAnchor="page" w:x="1475" w:y="2298"/>`;
    const items = layoutBody(framed("SideOne", page) + framed("SideTwo", page) + framed("SideThree", page) + BODY.repeat(2) + SECTION);
    const lines = ["SideOne", "SideTwo", "SideThree"].map((t) => line(items, t));
    expect(lines[0].lineTop).toBeCloseTo(2298 * TW, 0);
    for (let k = 1; k < lines.length; k++) {
      expect(lines[k].lineTop).toBeGreaterThanOrEqual(lines[k - 1].lineTop + lines[k - 1].lineHeight - 0.01);
      expect(lines[k].x).toBeCloseTo(1475 * TW, 1);
    }
  });

  it("keeps body text off the sides of a notBeside frame at the margin (word95-check-b frame-notbeside-margin)", () => {
    // The frame is anchored after three body paragraphs but sits at the
    // margin top: Word moves those paragraphs below it, and nothing beside it.
    const nb = `<w:framePr w:wrap="notBeside" w:vAnchor="margin" w:hAnchor="margin" w:x="1" w:y="1"/>`;
    const words = "Body words here. ".repeat(12);
    const para = (t: string) => `<w:p><w:r><w:t xml:space="preserve">${t}</w:t></w:r></w:p>`;
    const frameLines = ["NbOne", "NbTwo", "NbThree", "NbFour", "NbFive"];
    const items = layoutBody(
      para(`First. ${words}`) +
        para(`Second. ${words}`) +
        para(`Third. ${words}`) +
        frameLines.map((t) => framed(t, nb)).join("") +
        para(`After. ${words}`) +
        SECTION,
    );
    const frame = frameLines.map((t) => line(items, t));
    for (let k = 1; k < frame.length; k++) {
      expect(frame[k].lineTop).toBeGreaterThanOrEqual(frame[k - 1].lineTop + frame[k - 1].lineHeight - 0.01);
    }
    expect(frame[0].lineTop).toBeCloseTo(1440 * TW + 1 * TW, 0);
    const frameBottom = frame[4].lineTop + frame[4].lineHeight;
    const body = texts(items).filter((i) => !frameLines.includes(i.text));
    expect(body.length).toBeGreaterThan(0);
    for (const i of body) expect(i.lineTop).toBeGreaterThanOrEqual(frameBottom - 0.01);
    // The source order of the body is kept below the frame.
    expect(line(items, "First.").lineTop).toBeLessThan(line(items, "Second.").lineTop);
    expect(line(items, "Third.").lineTop).toBeLessThan(line(items, "After.").lineTop);
  });

  it("puts a margin-anchored frame with yAlign=bottom at the bottom of the margin box (Word 95 file 165082)", () => {
    const bottom = `<w:framePr w:w="9360" w:h="2736" w:hRule="exact" w:wrap="around" w:vAnchor="margin" w:hAnchor="margin" w:xAlign="right" w:yAlign="bottom"/>`;
    const items = layoutBody(framed("BoxLine", bottom) + BODY + SECTION);
    const box = line(items, "BoxLine");
    // Page 15840tw, bottom margin 1440tw: the 2736tw box ends at 14400tw.
    expect(box.lineTop).toBeCloseTo((15840 - 1440 - 2736) * TW, 0);
    expect(line(items, "Body").lineTop).toBeLessThan(box.lineTop);
  });

  it("leaves a drop cap as it is", () => {
    const drop =
      `<w:p><w:pPr><w:framePr w:dropCap="drop" w:lines="3" w:wrap="around" w:vAnchor="text" w:hAnchor="text"/></w:pPr>` +
      `<w:r><w:rPr><w:sz w:val="72"/></w:rPr><w:t>D</w:t></w:r></w:p>`;
    const items = layoutBody(drop + `<w:p><w:r><w:t>rop cap paragraph text.</w:t></w:r></w:p>` + SECTION);
    const letter = line(items, "D");
    expect(letter.x).toBeCloseTo(1800 * TW, 1);
    const rest = texts(items).find((i) => i.text.startsWith("rop"));
    expect(rest).toBeDefined();
    expect(rest!.x).toBeGreaterThan(letter.x + letter.width - 0.5);
    expect(rest!.lineTop).toBeCloseTo(letter.lineTop, 0);
  });

  it("stacks framed paragraphs in a header", () => {
    const R_NS = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
    const W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
    const header = `<?xml version="1.0"?><w:hdr ${W_NS}>${framed("HeaderOne")}${framed("HeaderTwo")}</w:hdr>`;
    const section = SECTION.replace("<w:sectPr>", `<w:sectPr><w:headerReference ${R_NS} w:type="default" r:id="rIdH"/>`);
    const items = layoutBody(`<w:p><w:r><w:t>Body</w:t></w:r></w:p>` + section, {
      "word/_rels/document.xml.rels":
        `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rIdH" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/></Relationships>`,
      "word/header1.xml": header,
    });
    const one = line(items, "HeaderOne");
    const two = line(items, "HeaderTwo");
    expect(two.lineTop).toBeGreaterThanOrEqual(one.lineTop + one.lineHeight - 0.01);
  });
});

describe("an exact-height frame", () => {
  it("clips its text, picture, shading and rules to its box, and paints nothing over the body", () => {
    // probe2-dropcaps-frames--v05: a 1000tw exact quote frame holds a picture
    // taller than the frame. Word cuts the picture and the shading off at the
    // frame's bottom; the body below must stay clear of the frame's paint.
    const picture =
      `<w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">` +
      `<wp:extent cx="1371600" cy="1371600"/><wp:docPr id="1" name="Pic"/>` +
      `<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
      `<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:blipFill>` +
      `<a:blip r:embed="rIdImg" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/></pic:blipFill>` +
      `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1371600" cy="1371600"/></a:xfrm></pic:spPr></pic:pic>` +
      `</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
    const pPr =
      `<w:pPr><w:framePr w:w="3000" w:h="1000" w:hRule="exact" w:wrap="around" w:vAnchor="text" w:hAnchor="margin" w:x="3000" w:y="0"/>` +
      `<w:pBdr>` +
      ["top", "left", "bottom", "right"].map((s) => `<w:${s} w:val="single" w:sz="6" w:space="4" w:color="BF8F00"/>`).join("") +
      `</w:pBdr><w:shd w:val="clear" w:color="auto" w:fill="FFF2CC"/><w:spacing w:after="0"/></w:pPr>`;
    const frame = `<w:p>${pPr}<w:r><w:t>Quote</w:t></w:r></w:p><w:p>${pPr}${picture}</w:p>`;
    const doc = DocxDocument.load(
      makeDocxWithMedia(
        {
          "word/document.xml": wrapDocument(frame + BODY.repeat(4) + SECTION),
          "word/_rels/document.xml.rels":
            `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
            `<Relationship Id="rIdImg" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/></Relationships>`,
        },
        { "word/media/image1.png": PNG },
      ),
    );
    const items = layoutDocument(doc, { measurer }).pages[0].items;
    const quote = line(items, "Quote");
    const fills = items.filter((i): i is RectItem => i.kind === "rect" && i.fill.toUpperCase().includes("FFF2CC"));
    const image = items.find((i): i is ImageItem => i.kind === "image");
    const rules = items.filter((i): i is LineEdgeItem => i.kind === "edge" && i.border.color.toUpperCase().includes("BF8F00"));
    expect(fills.length).toBeGreaterThan(0);
    expect(image).toBeDefined();
    // The frame's box: its top is the anchor paragraph's top (y=0), 1000tw tall.
    const top = Math.min(quote.lineTop, ...fills.map((f) => f.y), ...rules.map((r) => Math.min(r.y1, r.y2)));
    const bottom = top + 1000 * TW;
    const painted = [
      ...fills.map((f) => ({ top: f.y, bottom: f.y + f.height, left: f.x, right: f.x + f.width })),
      { top: image!.y, bottom: image!.y + image!.height, left: image!.x, right: image!.x + image!.width },
      ...rules.map((r) => ({
        top: Math.min(r.y1, r.y2),
        bottom: Math.max(r.y1, r.y2),
        left: Math.min(r.x1, r.x2),
        right: Math.max(r.x1, r.x2),
      })),
    ];
    // Every painted item of the frame stays inside the box.
    for (const r of painted) {
      expect(r.top).toBeGreaterThanOrEqual(top - 0.01);
      expect(r.bottom).toBeLessThanOrEqual(bottom + 0.6);
    }
    // The 1.5in picture is cut at the frame bottom, not drawn whole.
    expect(image!.height).toBeLessThan(1000 * TW);
    expect(image!.crop?.b ?? 0).toBeGreaterThan(0);
    // No body word sits on any of the frame's paint.
    const body = texts(items).filter((i) => i.text !== "Quote");
    expect(body.length).toBeGreaterThan(20);
    for (const t of body) {
      for (const r of painted) {
        const overlaps =
          t.lineTop < r.bottom && t.lineTop + t.lineHeight > r.top && t.x < r.right && t.x + t.width > r.left;
        expect(overlaps, `"${t.text}" at ${t.x.toFixed(1)},${t.lineTop.toFixed(1)}`).toBe(false);
      }
    }
  });
});
