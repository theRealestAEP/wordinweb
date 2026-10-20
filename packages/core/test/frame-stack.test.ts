import { describe, expect, it } from "vitest";
import { DocxDocument } from "../src/docx.js";
import { layoutDocument } from "../src/layout/engine.js";
import { ApproxMeasurer } from "../src/layout/measure.js";
import type { PageItem, TextItem } from "../src/layout/types.js";
import { makeDocx, wrapDocument } from "./helpers.js";

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
