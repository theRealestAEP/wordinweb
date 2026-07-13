import { describe, expect, it } from "vitest";
import { DocxDocument } from "../src/docx.js";
import { layoutDocument } from "../src/layout/engine.js";
import { ApproxMeasurer, type TextMeasurer } from "../src/layout/measure.js";
import { formatNumber } from "../src/parse/numbering.js";
import { makeDocx, wrapDocument, p, W_NS } from "./helpers.js";
import { layoutMath } from "../src/layout/math.js";
import type { MathNode } from "../src/model.js";

const measurer = new ApproxMeasurer();

function layout(parts: Record<string, string>) {
  const doc = DocxDocument.load(makeDocx(parts));
  return { doc, result: layoutDocument(doc, { measurer }) };
}

function pageText(result: ReturnType<typeof layoutDocument>, pageIdx: number): string {
  return result.pages[pageIdx].items
    .filter((i) => i.kind === "text")
    .map((i) => (i.kind === "text" ? i.text : ""))
    .join("");
}

describe("number formatting", () => {
  it("formats roman and letters", () => {
    expect(formatNumber(4, "lowerRoman")).toBe("iv");
    expect(formatNumber(1949, "upperRoman")).toBe("MCMXLIX");
    expect(formatNumber(1, "upperLetter")).toBe("A");
    expect(formatNumber(27, "lowerLetter")).toBe("aa");
  });
});

describe("layout engine", () => {
  it("produces a single page for a short document", () => {
    const { result } = layout({ "word/document.xml": wrapDocument(p("Hello world")) });
    expect(result.totalPages).toBe(1);
    expect(pageText(result, 0)).toContain("Hello");
  });

  it("run spaces give the justify packer no compression budget (typed spaces re-wrap the line)", () => {
    // Typing spaces mid-line in a justified paragraph must eventually push
    // the following words to a wrap. Pre-fix, every typed space ADDED
    // compression budget (compress = overflow / ALL spaces), so the line
    // never re-wrapped and every other space squeezed — the "space grows
    // backwards" editing bug. Run spaces (2+ adjacent) are excluded from the
    // budget and from alignment compression.
    const section =
      `<w:sectPr><w:pgSz w:w="6000" w:h="15840"/>` +
      `<w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720"/></w:sectPr>`;
    const wordsOnFirstLine = (mid: string): string[] => {
      const para =
        `<w:p><w:pPr><w:jc w:val="both"/></w:pPr><w:r><w:t xml:space="preserve">` +
        `alpha bravo charlie delta echo${mid}foxtrot golf hotel india juliet kilo lima mike november oscar` +
        `</w:t></w:r></w:p>`;
      const { result } = layout({ "word/document.xml": wrapDocument(para + section) });
      const first = Math.min(
        ...result.pages[0].items.filter((i) => i.kind === "text").map((i) => (i.kind === "text" ? i.lineTop : 1e9)),
      );
      return result.pages[0].items
        .filter((i) => i.kind === "text" && Math.abs(i.lineTop - first) < 0.5 && i.text.trim())
        .map((i) => (i.kind === "text" ? i.text : ""));
    };
    const base = wordsOnFirstLine(" ");
    const spaced = wordsOnFirstLine(" ".repeat(30));
    // 30 typed spaces (~90px) must displace words to the next line — the
    // pre-fix behavior kept the identical word set by compressing.
    expect(spaced.length).toBeLessThan(base.length);
    // And the run spaces themselves paint at natural width (uncompressed):
    // consecutive space items advance by a full space width each.
    const para =
      `<w:p><w:pPr><w:jc w:val="both"/></w:pPr><w:r><w:t xml:space="preserve">` +
      `alpha bravo charlie delta echo${" ".repeat(6)}foxtrot golf hotel india juliet kilo lima mike november oscar` +
      `</w:t></w:r></w:p>`;
    const { result } = layout({ "word/document.xml": wrapDocument(para + section) });
    const first = Math.min(
      ...result.pages[0].items.filter((i) => i.kind === "text").map((i) => (i.kind === "text" ? i.lineTop : 1e9)),
    );
    const spaces = result.pages[0].items
      .filter((i) => i.kind === "text" && i.text === " " && Math.abs(i.lineTop - first) < 0.5)
      .map((i) => (i.kind === "text" ? { x: i.x, w: i.width } : { x: 0, w: 0 }))
      .sort((a, b) => a.x - b.x);
    // Find the run: 6 consecutive spaces with equal advances.
    let runStart = -1;
    for (let i = 0; i + 5 < spaces.length; i++) {
      if (spaces[i + 5].x - spaces[i].x < 6 * spaces[i].w + 1) { runStart = i; break; }
    }
    expect(runStart).toBeGreaterThanOrEqual(0);
    const runW = spaces[runStart].w;
    for (let i = runStart; i < runStart + 6; i++) {
      expect(spaces[i].w).toBeCloseTo(runW, 1); // uniform, uncompressed
    }
  });

  it("hangs wrap-boundary spaces at the end of the wrapping line with source bindings", () => {
    const section =
      `<w:sectPr><w:pgSz w:w="3600" w:h="15840"/>` +
      `<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>`;
    const resultFor = (text: string) =>
      layout({ "word/document.xml": wrapDocument(p(text) + section) }).result;
    const item = (result: ReturnType<typeof layoutDocument>, text: string) => {
      const found = result.pages[0].items.find(
        (candidate) => candidate.kind === "text" && candidate.text === text,
      );
      if (found?.kind !== "text") throw new Error(`missing ${text}`);
      return found;
    };

    const single = resultFor("AAAA BBBB");
    const triple = resultFor("AAAA   BBBB");
    const singleA = item(single, "AAAA");
    const singleB = item(single, "BBBB");
    const tripleA = item(triple, "AAAA");
    const tripleB = item(triple, "BBBB");

    // Word never starts a wrapped line with a space: extra spaces at the
    // wrap boundary hang past the end of the upper line (zero ink, real
    // advances, caret-addressable) and the next line's first word does not
    // move.
    expect(tripleB.x).toBeCloseTo(singleB.x, 3);
    expect(tripleB.lineTop).toBeCloseTo(singleB.lineTop, 3);
    const hanging = triple.pages[0].items.filter(
      (candidate) =>
        candidate.kind === "text" &&
        candidate.text === " " &&
        candidate.lineTop === tripleA.lineTop,
    );
    expect(hanging).toHaveLength(3);
    expect(hanging.map((space) => (space.kind === "text" ? space.src?.offset : undefined))).toEqual([4, 5, 6]);
    // Sequential advances starting at the end of the wrapping word.
    let edge = tripleA.x + tripleA.width;
    for (const space of hanging) {
      if (space.kind !== "text") continue;
      expect(space.x).toBeCloseTo(edge, 3);
      edge += space.width;
    }
    // The single-space layout hangs its separator space too.
    const separator = single.pages[0].items.filter(
      (candidate) =>
        candidate.kind === "text" &&
        candidate.text === " " &&
        candidate.lineTop === singleA.lineTop,
    );
    expect(separator).toHaveLength(1);
  });

  it("paginates long content onto multiple pages", () => {
    const paras = Array.from({ length: 120 }, (_, i) => p(`Paragraph number ${i}`)).join("");
    const { result } = layout({ "word/document.xml": wrapDocument(paras) });
    expect(result.totalPages).toBeGreaterThan(1);
    // Content must stay inside the body box (1in margins on US Letter).
    for (const page of result.pages) {
      for (const item of page.items) {
        if (item.kind === "text") {
          expect(item.lineTop).toBeGreaterThanOrEqual(95); // ~96px margin
          expect(item.lineTop + item.lineHeight).toBeLessThanOrEqual(page.height - 95);
        }
      }
    }
  });

  it("honors explicit page breaks", () => {
    const body =
      p("first page") +
      `<w:p><w:r><w:br w:type="page"/></w:r><w:r><w:t>second page</w:t></w:r></w:p>`;
    const { result } = layout({ "word/document.xml": wrapDocument(body) });
    expect(result.totalPages).toBe(2);
    expect(pageText(result, 0)).toContain("first page");
    expect(pageText(result, 0)).not.toContain("second page");
    expect(pageText(result, 1)).toContain("second page");
  });

  it("does not reset following content after internal page breaks", () => {
    const body =
      p("first page") +
      `<w:p><w:r><w:br w:type="page"/><w:br w:type="page"/>` +
      `<w:t>after breaks</w:t></w:r></w:p>` +
      `<w:p><w:pPr><w:spacing w:before="240"/></w:pPr>` +
      `<w:r><w:t>following</w:t></w:r></w:p>`;
    const { result } = layout({ "word/document.xml": wrapDocument(body) });
    const page = result.pages[2];
    const after = page.items.find((item) => item.kind === "text" && item.text === "after");
    const following = page.items.find((item) => item.kind === "text" && item.text === "following");

    expect(result.totalPages).toBe(3);
    expect(after?.kind).toBe("text");
    expect(following?.kind).toBe("text");
    if (after?.kind !== "text" || following?.kind !== "text") return;
    expect(following.lineTop).toBeGreaterThan(after.lineTop + after.lineHeight);
  });

  it("honors paragraph opt-out from a section line grid", () => {
    const gridDocument = (snapToGrid = "") =>
      wrapDocument(
        `<w:p><w:pPr>${snapToGrid}</w:pPr><w:r><w:rPr><w:sz w:val="16"/></w:rPr><w:t>Title</w:t></w:r></w:p>` +
          `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>` +
          `<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>` +
          `<w:docGrid w:type="lines" w:linePitch="360"/></w:sectPr>`,
      );
    const inherited = layout({ "word/document.xml": gridDocument() }).result;
    const optedOut = layout({
      "word/document.xml": gridDocument(`<w:snapToGrid w:val="0"/>`),
    }).result;
    const inheritedTitle = inherited.pages[0].items.find((item) => item.kind === "text" && item.text === "Title");
    const optedOutTitle = optedOut.pages[0].items.find((item) => item.kind === "text" && item.text === "Title");

    expect(inheritedTitle?.kind).toBe("text");
    expect(optedOutTitle?.kind).toBe("text");
    if (inheritedTitle?.kind !== "text" || optedOutTitle?.kind !== "text") return;
    // Inherited grid paragraphs keep the measured four-pitch section reserve
    // and a pitch-sized line. An explicit opt-out uses a two-pitch reserve and
    // its natural line box.
    expect(inherited.pages[0].bodyTop - optedOut.pages[0].bodyTop).toBeCloseTo(48, 3);
    expect(inheritedTitle.lineTop - optedOutTitle.lineTop).toBeCloseTo(48, 3);
    expect(inheritedTitle.lineHeight).toBeCloseTo(24, 3);
    expect(optedOutTitle.lineHeight).toBeLessThan(24);
  });

  it("keeps a grid reserve above auto-spaced inline images", () => {
    const rels = `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdImg" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/x.png"/>
</Relationships>`;
    const inlineImage =
      `<w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">` +
      `<wp:extent cx="769620" cy="427990"/>` +
      `<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
      `<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
      `<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
      `<pic:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rIdImg"/></pic:blipFill>` +
      `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="769620" cy="427990"/></a:xfrm></pic:spPr>` +
      `</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>`;
    const documentXml = (snapToGrid = "") =>
      wrapDocument(
        `<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="330" w:lineRule="exact"/></w:pPr>` +
          `<w:r><w:t>filler</w:t></w:r></w:p>` +
          `<w:p><w:pPr>${snapToGrid}<w:spacing w:before="0" w:after="156" w:line="276" w:lineRule="auto"/></w:pPr>` +
          `<w:r><w:rPr><w:sz w:val="21"/></w:rPr>${inlineImage}</w:r>` +
          `<w:r><w:rPr><w:sz w:val="21"/></w:rPr><w:t>(0)</w:t></w:r></w:p>` +
          `<w:sectPr><w:pgSz w:w="6000" w:h="3800"/>` +
          `<w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720"/>` +
          `<w:docGrid w:type="lines" w:linePitch="312"/></w:sectPr>`,
      );
    const parts = (snapToGrid = "") => ({
      "word/document.xml": documentXml(snapToGrid),
      "word/_rels/document.xml.rels": rels,
      "word/media/x.png": "PNGDATA",
    });
    const grid = layout(parts()).result;
    const optedOut = layout(parts(`<w:snapToGrid w:val="0"/>`)).result;

    expect(pageText(grid, 0)).not.toContain("(0)");
    expect(pageText(grid, 1)).toContain("(0)");
    expect(pageText(optedOut, 0)).toContain("(0)");

    const gridLabel = grid.pages[1].items.find((item) => item.kind === "text" && item.text === "(0)");
    const gridImage = grid.pages[1].items.find((item) => item.kind === "image");
    const controlLabel = optedOut.pages[0].items.find((item) => item.kind === "text" && item.text === "(0)");
    const controlImage = optedOut.pages[0].items.find((item) => item.kind === "image");
    expect(gridLabel?.kind).toBe("text");
    expect(gridImage?.kind).toBe("image");
    expect(controlLabel?.kind).toBe("text");
    expect(controlImage?.kind).toBe("image");
    if (
      gridLabel?.kind !== "text" ||
      gridImage?.kind !== "image" ||
      controlLabel?.kind !== "text" ||
      controlImage?.kind !== "image"
    ) return;

    // A grid object line snaps to whole grid pitches with the content extent
    // centered (measured in wild2-math-eq-as-images-word.pdf): the 44.93px
    // image + 3.5px text descent = 48.43px extent exceeds the 23.92px text
    // line, so the line takes ceil(48.43/20.8) = 3 pitches = 62.4px and the
    // image top sits (62.4 - 48.43)/2 below the line top. The non-grid
    // control retains the compact image-line rule (~48.9px) and stays on
    // page 1; the 62.4px grid box does not fit and breaks to page 2.
    const extent = 44.93333333333333 + 14 * 0.25;
    const expectedGridHeight = 3 * 20.8;
    expect(gridLabel.lineHeight).toBeCloseTo(expectedGridHeight, 3);
    // (paint baselines quantize to quarter-points, so allow 0.5px)
    expect(gridImage.y - gridLabel.lineTop).toBeCloseTo((expectedGridHeight - extent) / 2, 0);
    expect(Math.abs(controlImage.y - controlLabel.lineTop)).toBeLessThan(0.2);
    expect(controlLabel.lineHeight).toBeLessThan(gridLabel.lineHeight);
  });

  it("lays a docGrid equation-image line as whole grid pitches, centered, with w:position lowering the image", () => {
    // wild2-math-eq-as-images eq(48): a 290.75x57.4pt VML pict (rounds to
    // 291x57pt) on run position -47hp (-23.5pt) in a linePitch=312 (15.6pt)
    // grid with spacing 348 atLeast. Word's PDF: the line takes 4 pitches
    // (62.4pt = 83.2px), the image spans 33.5pt above / 23.5pt below the
    // baseline, and the image top sits (62.4-57)/2 = 2.7pt below the line
    // top (shading rect 643.44 vs image top 640.75). A 50.6x31.45pt pict at
    // position -22hp rounds to 31pt and takes exactly 2 pitches (31.2pt) -
    // unrounded 31.45pt would wrongly take 3.
    const rels = `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdImg" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/eq.png"/>
</Relationships>`;
    const pict = (w: string, h: string) =>
      `<w:pict><v:shape xmlns:v="urn:schemas-microsoft-com:vml" style="width:${w}pt;height:${h}pt">` +
      `<v:imagedata xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rIdImg"/>` +
      `</v:shape></w:pict>`;
    const eqPara = (pos: number, w: string, h: string, label: string) =>
      `<w:p><w:pPr><w:spacing w:before="156" w:after="156" w:line="348" w:lineRule="atLeast"/></w:pPr>` +
      `<w:r><w:rPr><w:position w:val="${pos}"/><w:sz w:val="21"/></w:rPr>${pict(w, h)}</w:r>` +
      `<w:r><w:rPr><w:sz w:val="21"/></w:rPr><w:t>${label}</w:t></w:r></w:p>`;
    const parts = {
      "word/document.xml": wrapDocument(
        eqPara(-47, "290.75", "57.4", "(48)") +
          eqPara(-22, "50.6", "31.45", "(76)") +
          `<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>` +
          `<w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800"/>` +
          `<w:docGrid w:type="lines" w:linePitch="312"/></w:sectPr>`,
      ),
      "word/_rels/document.xml.rels": rels,
      "word/media/eq.png": "PNGDATA",
    };
    const { result } = layout(parts);
    const page = result.pages[0];
    const labels = page.items.filter((it) => it.kind === "text" && (it.text === "(48)" || it.text === "(76)"));
    const images = page.items.filter((it) => it.kind === "image");
    expect(labels).toHaveLength(2);
    expect(images).toHaveLength(2);
    if (labels[0].kind !== "text" || labels[1].kind !== "text" || images[0].kind !== "image" || images[1].kind !== "image") return;
    const pitch = 20.8; // 312tw in px
    // eq48: extent = rounded 57pt image (76px) -> 4 pitches.
    expect(images[0].height).toBeCloseTo(76, 3);
    expect(labels[0].lineHeight).toBeCloseTo(4 * pitch, 3);
    // Centered: image top = line top + (H - extent)/2 (paint baselines
    // quantize to quarter-points, so allow 0.5px).
    expect(images[0].y - labels[0].lineTop).toBeCloseTo((4 * pitch - 76) / 2, 0);
    // The lowered image hangs 23.5pt (31.33px) below the label baseline.
    const baseline48 = labels[0].baseline;
    expect(images[0].y + images[0].height - baseline48).toBeCloseTo(23.5 * (4 / 3), 1);
    // eq76: 31.45pt rounds to 31pt (41.33px) -> exactly 2 pitches.
    expect(images[1].height).toBeCloseTo(31 * (4 / 3), 3);
    expect(labels[1].lineHeight).toBeCloseTo(2 * pitch, 3);
  });

  it("preserves a section line-grid opt-out through column balancing", () => {
    const firstSection =
      `<w:p><w:pPr><w:snapToGrid w:val="0"/></w:pPr>` +
      `<w:r><w:rPr><w:sz w:val="16"/></w:rPr><w:t>Title</w:t></w:r></w:p>` +
      `<w:p><w:pPr><w:sectPr><w:pgSz w:w="12240" w:h="15840"/>` +
      `<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>` +
      `<w:cols w:num="2" w:space="720"/>` +
      `<w:docGrid w:type="lines" w:linePitch="360"/>` +
      `</w:sectPr></w:pPr></w:p>`;
    const nextSection =
      p("Next") +
      `<w:sectPr><w:type w:val="continuous"/><w:pgSz w:w="12240" w:h="15840"/>` +
      `<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>`;
    const { result } = layout({
      "word/document.xml": wrapDocument(firstSection + nextSection),
    });
    const title = result.pages[0].items.find(
      (item) => item.kind === "text" && item.text === "Title",
    );

    expect(result.pages[0].bodyTop).toBeCloseTo(144, 3);
    expect(title?.kind).toBe("text");
    if (title?.kind !== "text") return;
    expect(title.lineTop).toBeCloseTo(144, 3);
  });

  it("draws a vertical rule between columns for w:cols w:sep and honours per-column widths", () => {
    // w:cols w:sep="1" paints a rule centered in each inter-column gap; explicit
    // unequal w:col widths/spaces are honoured raw (probe3-columns-unequal).
    const body =
      p("Left column body text here") +
      `<w:p><w:r><w:t>Second column body text</w:t></w:r><w:r><w:br w:type="column"/></w:r></w:p>` +
      p("Tail") +
      `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>` +
      `<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>` +
      `<w:cols w:num="2" w:sep="1" w:equalWidth="0">` +
      `<w:col w:w="5040" w:space="360"/><w:col w:w="3960"/></w:cols></w:sectPr>`;
    const { result } = layout({ "word/document.xml": wrapDocument(body) });
    const edges = result.pages[0].items.filter(
      (i) => i.kind === "edge" && i.x1 === i.x2,
    );
    expect(edges.length).toBeGreaterThanOrEqual(1);
    // The rule sits in the gap between col 1 (left margin 1440tw = 96px, width
    // 5040tw = 336px) and col 2 (starts after a 360tw = 24px space): centered at
    // 96 + 336 + 24/2 = 444px.
    const rule = edges.find((e) => e.kind === "edge" && Math.abs(e.x1 - 444) < 1);
    expect(rule).toBeDefined();
  });

  it("bottom-aligns a legacy doc-grid keepNext chain before a leading page break", () => {
    // Word PDF control (wild2-med-nccih-protocol): the first title bbox starts
    // at 614.92pt. Disabling keepNext on either heading, removing docGrid, or
    // changing compatibilityMode 14 to 15 moves it to 329.07pt. Word 2010
    // keeps the leading break line on page 1 and aligns the kept chain above it.
    const body =
      `<w:p><w:pPr><w:pStyle w:val="Heading1"/><w:spacing w:before="5120"/>` +
      `<w:jc w:val="center"/><w:rPr><w:sz w:val="52"/></w:rPr></w:pPr>` +
      `<w:r><w:rPr><w:sz w:val="52"/></w:rPr><w:t>Cover</w:t></w:r>` +
      `<w:r><w:rPr><w:sz w:val="52"/></w:rPr><w:br/></w:r>` +
      `<w:r><w:rPr><w:sz w:val="52"/></w:rPr><w:t>Title</w:t></w:r></w:p>` +
      `<w:p><w:pPr><w:pStyle w:val="Heading1"/><w:jc w:val="center"/></w:pPr>` +
      `<w:r><w:rPr><w:sz w:val="52"/></w:rPr><w:br w:type="page"/></w:r>` +
      `<w:r><w:rPr><w:sz w:val="32"/></w:rPr><w:t>PAGE TWO</w:t></w:r></w:p>` +
      p("body") +
      `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>` +
      `<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>` +
      `<w:titlePg/><w:docGrid w:linePitch="360"/></w:sectPr>`;
    const { result } = layout({
      "word/document.xml": wrapDocument(body),
      "word/styles.xml": `<?xml version="1.0"?>
        <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
            <w:rPr><w:sz w:val="24"/></w:rPr>
          </w:style>
          <w:style w:type="paragraph" w:styleId="Heading1">
            <w:basedOn w:val="Normal"/>
            <w:pPr><w:keepNext/><w:spacing w:before="360" w:after="120" w:line="240" w:lineRule="atLeast"/></w:pPr>
            <w:rPr><w:rFonts w:ascii="Arial Bold" w:hAnsi="Arial Bold"/><w:b/></w:rPr>
          </w:style>
        </w:styles>`,
      "word/settings.xml": `<?xml version="1.0"?>
        <w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:compat><w:compatSetting w:name="compatibilityMode" w:val="14"/></w:compat>
        </w:settings>`,
    });

    const title = result.pages[0].items.find((item) => item.kind === "text" && item.text === "Cover");
    expect(title?.kind).toBe("text");
    if (title?.kind !== "text") return;
    expect(title.lineTop).toBeCloseTo(816.4, 1);
    expect(pageText(result, 1)).toContain("PAGE TWO");
  });

  it("collapses preceding after-spacing across a legacy leading page break", () => {
    const render = (mode: number, after: number) => {
      const body =
        `<w:p><w:pPr><w:spacing w:after="${after}"/></w:pPr>` +
        `<w:r><w:t>OLD</w:t></w:r></w:p>` +
        `<w:p><w:pPr><w:spacing w:before="240"/></w:pPr>` +
        `<w:r><w:br w:type="page"/></w:r><w:r><w:t>OPEN</w:t></w:r></w:p>` +
        `<w:sectPr><w:pgSz w:w="6000" w:h="6000"/>` +
        `<w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720"/></w:sectPr>`;
      return layout({
        "word/document.xml": wrapDocument(body),
        "word/settings.xml": `<?xml version="1.0"?>
          <w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:compat><w:compatSetting w:name="compatibilityMode" w:val="${mode}"/></w:compat>
          </w:settings>`,
      }).result;
    };
    const openerOffset = (result: ReturnType<typeof layoutDocument>) => {
      const page = result.pages[1];
      const opener = page.items.find((item) => item.kind === "text" && item.text === "OPEN");
      if (opener?.kind !== "text") throw new Error("opener missing");
      return opener.lineTop - page.bodyTop;
    };

    expect(openerOffset(render(14, 120))).toBeCloseTo(8, 3);
    expect(openerOffset(render(14, 0))).toBeCloseTo(16, 3);
    expect(openerOffset(render(15, 120))).toBeCloseTo(0, 3);
  });

  it("collapses an empty section closer's after into a legacy leading-break opener", () => {
    const pageGeometry =
      `<w:pgSz w:w="6000" w:h="6000"/>` +
      `<w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720"/>`;
    const render = (mode: number, leadingBreak: boolean) => {
      const closer =
        `<w:p><w:pPr><w:spacing w:after="120"/></w:pPr></w:p>` +
        `<w:p><w:pPr><w:spacing w:after="120"/><w:sectPr>` +
        `<w:type w:val="continuous"/>${pageGeometry}</w:sectPr></w:pPr></w:p>`;
      const opener =
        `<w:p><w:pPr><w:spacing w:before="360"/></w:pPr>` +
        (leadingBreak ? `<w:r><w:br w:type="page"/></w:r>` : "") +
        `<w:r><w:t>OPEN</w:t></w:r></w:p>`;
      return layout({
        "word/document.xml": wrapDocument(
          closer + opener + `<w:sectPr>${leadingBreak ? `<w:type w:val="continuous"/>` : ""}${pageGeometry}</w:sectPr>`,
        ),
        "word/settings.xml": `<?xml version="1.0"?>
          <w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:compat><w:compatSetting w:name="compatibilityMode" w:val="${mode}"/></w:compat>
          </w:settings>`,
      }).result;
    };
    const openerOffset = (result: ReturnType<typeof layoutDocument>) => {
      const page = result.pages[1];
      const opener = page.items.find((item) => item.kind === "text" && item.text === "OPEN");
      if (opener?.kind !== "text") throw new Error("section opener missing");
      return opener.lineTop - page.bodyTop;
    };

    expect(openerOffset(render(14, true))).toBeCloseTo(16, 3);
    expect(openerOffset(render(14, false))).toBeCloseTo(24, 3);
    expect(openerOffset(render(15, true))).toBeCloseTo(0, 3);
  });

  it("keeps a trailing-break section closer's mark before a continuous section", () => {
    const pageGeometry =
      `<w:pgSz w:w="6000" w:h="6000"/>` +
      `<w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720"/>`;
    const render = (type: "continuous" | "nextPage") => {
      const closer =
        `<w:p><w:pPr><w:pStyle w:val="Heading2"/><w:sectPr>${pageGeometry}</w:sectPr></w:pPr>` +
        `<w:r><w:br w:type="page"/></w:r></w:p>`;
      const opener =
        `<w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr>` +
        `<w:r><w:t>OPEN</w:t></w:r></w:p>`;
      return layout({
        "word/document.xml": wrapDocument(
          closer + opener + `<w:sectPr><w:type w:val="${type}"/>${pageGeometry}</w:sectPr>`,
        ),
        "word/styles.xml": `<?xml version="1.0"?>
          <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
              <w:rPr><w:sz w:val="24"/></w:rPr>
            </w:style>
            <w:style w:type="paragraph" w:styleId="Heading2">
              <w:basedOn w:val="Normal"/>
              <w:pPr><w:spacing w:before="240" w:after="120" w:line="240" w:lineRule="atLeast"/></w:pPr>
              <w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:b/></w:rPr>
            </w:style>
          </w:styles>`,
        "word/settings.xml": `<?xml version="1.0"?>
          <w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:compat><w:compatSetting w:name="compatibilityMode" w:val="14"/></w:compat>
          </w:settings>`,
      }).result;
    };
    const opener = (result: ReturnType<typeof layoutDocument>) => {
      const page = result.pages[1];
      const item = page.items.find((candidate) => candidate.kind === "text" && candidate.text === "OPEN");
      if (item?.kind !== "text") throw new Error("section opener missing");
      return { page, item };
    };

    const continuousResult = render("continuous");
    const nextPageResult = render("nextPage");
    const continuous = opener(continuousResult);
    const nextPage = opener(nextPageResult);
    expect(continuousResult.totalPages).toBe(2);
    expect(nextPageResult.totalPages).toBe(2);
    // NCCIH's Word PDF measures this hidden line at 13.75pt; its 6pt after
    // remains in addition to the visible Heading2's own 12pt before.
    expect(continuous.item.lineTop - continuous.page.bodyTop).toBeCloseTo(
      16 + continuous.item.lineHeight + 8,
      3,
    );
    expect(nextPage.item.lineTop - nextPage.page.bodyTop).toBeCloseTo(16, 3);
  });

  it("lets an ordinary break-only paragraph overflow before applying its page break", () => {
    const body =
      `<w:p><w:pPr><w:spacing w:after="1200"/></w:pPr><w:r><w:t>first page</w:t></w:r></w:p>` +
      `<w:p><w:r><w:br w:type="page"/></w:r></w:p>` +
      `<w:p><w:r><w:lastRenderedPageBreak/><w:t>second page</w:t></w:r></w:p>` +
      `<w:sectPr><w:pgSz w:w="12240" w:h="3000"/>` +
      `<w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720"/></w:sectPr>`;
    const { result } = layout({ "word/document.xml": wrapDocument(body) });

    expect(result.totalPages).toBe(3);
    expect(pageText(result, 0)).toContain("first page");
    expect(pageText(result, 1)).toBe("");
    expect(pageText(result, 2)).toContain("second page");
  });

  it("does not add another empty line after a table's mandatory empty paragraph", () => {
    const table = `<w:tbl>
      <w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>
      <w:tr><w:trPr><w:trHeight w:val="900" w:hRule="exact"/></w:trPr>
        <w:tc><w:p><w:r><w:t>table</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>`;
    const body =
      `<w:p><w:r><w:t>first page</w:t></w:r></w:p>` +
      table +
      `<w:p/>` +
      `<w:p><w:r><w:br w:type="page"/></w:r></w:p>` +
      `<w:p><w:r><w:lastRenderedPageBreak/><w:t>second page</w:t></w:r></w:p>` +
      `<w:sectPr><w:pgSz w:w="12240" w:h="3000"/>` +
      `<w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720"/></w:sectPr>`;
    const { result } = layout({ "word/document.xml": wrapDocument(body) });

    expect(result.totalPages).toBe(2);
    expect(pageText(result, 0)).toContain("first page");
    expect(pageText(result, 1)).toContain("second page");
  });

  it("lets an authored empty paragraph after a table leave a break on a blank page", () => {
    const table = `<w:tbl>
      <w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>
      <w:tr><w:trPr><w:trHeight w:val="900" w:hRule="exact"/></w:trPr>
        <w:tc><w:p><w:r><w:t>table</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>`;
    const body =
      `<w:p><w:r><w:t>first page</w:t></w:r></w:p>` +
      table +
      `<w:p><w:pPr><w:spacing w:before="0" w:after="0"/><w:rPr><w:b/></w:rPr></w:pPr></w:p>` +
      `<w:p><w:r><w:br w:type="page"/></w:r></w:p>` +
      `<w:p><w:r><w:lastRenderedPageBreak/><w:t>third page</w:t></w:r></w:p>` +
      `<w:sectPr><w:pgSz w:w="12240" w:h="3000"/>` +
      `<w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720"/></w:sectPr>`;
    const { result } = layout({ "word/document.xml": wrapDocument(body) });

    expect(result.totalPages).toBe(3);
    expect(pageText(result, 0)).toContain("first page");
    expect(pageText(result, 1)).toBe("");
    expect(pageText(result, 2)).toContain("third page");
  });

  it("flows a table's continuation row into the next column at that column's x", () => {
    // staging-breaks p4: a 2-row table in a multi-column section whose second
    // row does not fit in column 1 flows into column 2. The continuation row
    // must paint at COLUMN 2's x, not the original column's - a table split
    // across columns used to keep the first column's x0 and overlap row 1 on
    // top of row 0.
    const filler = Array.from({ length: 9 }, (_, i) => p(`Filler line ${i}`)).join("");
    const table =
      `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>` +
      `<w:tblGrid><w:gridCol w:w="1400"/><w:gridCol w:w="1400"/></w:tblGrid>` +
      `<w:tr><w:tc><w:tcPr><w:tcW w:w="2800" w:type="dxa"/><w:gridSpan w:val="2"/></w:tcPr>` +
      `<w:p><w:r><w:t>ROWZERO</w:t></w:r></w:p></w:tc></w:tr>` +
      `<w:tr><w:tc><w:tcPr><w:tcW w:w="1400" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>AA</w:t></w:r></w:p></w:tc>` +
      `<w:tc><w:tcPr><w:tcW w:w="1400" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>BB</w:t></w:r></w:p></w:tc></w:tr>` +
      `</w:tbl>`;
    const section =
      `<w:p><w:pPr><w:sectPr><w:cols w:num="2" w:space="720"/>` +
      `<w:pgSz w:w="12240" w:h="4000"/>` +
      `<w:pgMar w:top="720" w:right="1440" w:bottom="720" w:left="1440"/></w:sectPr></w:pPr></w:p>`;
    const { result } = layout({ "word/document.xml": wrapDocument(filler + table + section) });
    const rowZero = result.pages[0].items.find((i) => i.kind === "text" && i.text === "ROWZERO");
    const aa = result.pages[0].items.find((i) => i.kind === "text" && i.text === "AA");
    expect(rowZero?.kind).toBe("text");
    expect(aa?.kind).toBe("text");
    if (rowZero?.kind !== "text" || aa?.kind !== "text") return;
    // Row 1 ("AA") lands in a later column, well to the right of row 0.
    expect(aa.x).toBeGreaterThan(rowZero.x + 200);
  });

  it("drops space-before when a keepLines paragraph is moved to a column top", () => {
    // A multi-line keepLines paragraph with a large space-before that cannot fit
    // in the remaining space at the bottom of a filled column is moved whole to
    // the top of the next column. Word collapses its space-before against the
    // column top, so its first line must sit AT the band top - not one
    // before-height (400 twips = 20pt ~ 27px) lower. Regression guard for the
    // wild-multicolumn sliver-heading drift: the keepLines/keepNext move used to
    // keep the before and shifted the whole one-glyph column down.
    const longHeading =
      "Heading text long enough to wrap across two lines so keepLines keeps them " +
      "together when it moves to the next column top of this section body";
    const filler = Array.from({ length: 48 }, (_, i) => p(`Filler line number ${i} here`)).join("");
    const headingPara =
      `<w:p><w:pPr><w:keepLines/><w:spacing w:before="400"/></w:pPr>` +
      `<w:r><w:t xml:space="preserve">${longHeading}</w:t></w:r></w:p>`;
    const twoCol =
      `<w:p><w:pPr><w:sectPr><w:cols w:num="2" w:space="720"/>` +
      `<w:pgSz w:w="12240" w:h="15840"/>` +
      `<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:pPr></w:p>`;
    const { result } = layout({ "word/document.xml": wrapDocument(filler + headingPara + twoCol) });
    const heading = result.pages[0].items.find(
      (i) => i.kind === "text" && i.text.includes("Heading"),
    );
    expect(heading?.kind).toBe("text");
    if (heading?.kind !== "text") return;
    // Landed in the SECOND column (x well past the left margin), at its top.
    expect(heading.x).toBeGreaterThan(300);
    expect(heading.lineTop).toBeCloseTo(96, 0); // band top, before collapsed
  });

  it("does not charge drop-cap height to a keepNext chain", () => {
    const filler = Array.from(
      { length: 14 },
      (_, i) =>
        `<w:p><w:pPr><w:spacing w:line="240" w:lineRule="exact"/></w:pPr>` +
        `<w:r><w:t>FILL${i}</w:t></w:r></w:p>`,
    ).join("");
    const heading =
      `<w:p><w:pPr><w:keepNext/><w:spacing w:after="80"/></w:pPr>` +
      `<w:r><w:t>HEADING</w:t></w:r></w:p>`;
    const dropCap =
      `<w:p><w:pPr><w:keepNext/>` +
      `<w:framePr w:dropCap="drop" w:lines="2" w:wrap="auto"/>` +
      `<w:spacing w:line="480" w:lineRule="exact"/>` +
      `<w:rPr><w:sz w:val="56"/></w:rPr></w:pPr>` +
      `<w:r><w:rPr><w:sz w:val="56"/></w:rPr><w:t>D</w:t></w:r></w:p>`;
    const body =
      `<w:p><w:r><w:t xml:space="preserve">BODY one two three four five six seven eight nine ten eleven twelve</w:t></w:r></w:p>`;
    const section =
      `<w:sectPr><w:pgSz w:w="6000" w:h="6000"/>` +
      `<w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720"/>` +
      `<w:cols w:num="2" w:space="240"/></w:sectPr>`;
    const { result } = layout({
      "word/document.xml": wrapDocument(filler + heading + dropCap + body + section),
    });
    const page = result.pages[0];
    const headingText = page.items.find(
      (item) => item.kind === "text" && item.text === "HEADING",
    );
    const bodyText = page.items.find(
      (item) => item.kind === "text" && item.text.includes("BODY"),
    );

    expect(headingText?.kind).toBe("text");
    expect(bodyText?.kind).toBe("text");
    if (headingText?.kind !== "text" || bodyText?.kind !== "text") return;
    expect(headingText.x).toBeLessThan(208);
    expect(bodyText.x).toBeLessThan(208);
  });

  it("hangs a dropCap=margin letter into the margin with full-width body text", () => {
    // probe2-dropcaps-frames p1: Word renders w:dropCap="margin" with the big
    // letter OUT in the left margin (advance edge at the text margin) and the
    // following paragraph flowing at the FULL column width — unlike "drop",
    // which sinks the letter and indents the text around it.
    const marginCap =
      `<w:p><w:pPr>` +
      `<w:framePr w:dropCap="margin" w:lines="3" w:wrap="around" w:vAnchor="text" w:hAnchor="page"/>` +
      `<w:spacing w:after="0" w:line="240" w:lineRule="auto"/>` +
      `<w:rPr><w:sz w:val="84"/></w:rPr></w:pPr>` +
      `<w:r><w:rPr><w:sz w:val="84"/></w:rPr><w:t>M</w:t></w:r></w:p>`;
    const body =
      `<w:p><w:r><w:t xml:space="preserve">argin drop caps hang out into the left margin instead of sinking</w:t></w:r></w:p>`;
    const section =
      `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>` +
      `<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>`;
    const { result } = layout({
      "word/document.xml": wrapDocument(marginCap + body + section),
    });
    const page = result.pages[0];
    const cap = page.items.find((i) => i.kind === "text" && i.text === "M");
    const bodyText = page.items.find(
      (i) => i.kind === "text" && i.text.includes("argin"),
    );
    expect(cap?.kind).toBe("text");
    expect(bodyText?.kind).toBe("text");
    if (cap?.kind !== "text" || bodyText?.kind !== "text") return;
    const marginLeft = 1440 / 15; // 96px text margin (1440 twips = 1in = 96px)
    // The letter hangs into the margin: its left edge is LEFT of the text
    // margin, and its advance right edge sits AT the text margin.
    expect(cap.x).toBeLessThan(marginLeft - 1);
    expect(cap.x + cap.width).toBeCloseTo(marginLeft, 0);
    // Body text is NOT indented around the letter — it starts at the margin.
    expect(bodyText.x).toBeCloseTo(marginLeft, 0);
  });

  it("keeps a TEXT-anchored dropCap=margin letter at the column with drop-style wrap", () => {
    // parity2-dropcap p1: with hAnchor="text" (not "page"), Word keeps the
    // margin drop cap AT the column edge and wraps the body around it, exactly
    // like dropCap="drop" — the hang-into-margin treatment is page-anchor only.
    // (This fixture scored 0.00 for days, then regressed to 1.78 when the hang
    // path first shipped ungated.)
    const marginCap =
      `<w:p><w:pPr>` +
      `<w:framePr w:dropCap="margin" w:lines="3" w:wrap="around" w:vAnchor="text" w:hAnchor="text"/>` +
      `<w:spacing w:after="0" w:line="240" w:lineRule="auto"/>` +
      `<w:rPr><w:sz w:val="84"/></w:rPr></w:pPr>` +
      `<w:r><w:rPr><w:sz w:val="84"/></w:rPr><w:t>M</w:t></w:r></w:p>`;
    const body =
      `<w:p><w:r><w:t xml:space="preserve">argin drop caps with a text anchor sink into the column like drop caps do</w:t></w:r></w:p>`;
    const section =
      `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>` +
      `<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>`;
    const { result } = layout({
      "word/document.xml": wrapDocument(marginCap + body + section),
    });
    const page = result.pages[0];
    const cap = page.items.find((i) => i.kind === "text" && i.text === "M");
    const bodyText = page.items.find(
      (i) => i.kind === "text" && i.text.includes("argin"),
    );
    expect(cap?.kind).toBe("text");
    expect(bodyText?.kind).toBe("text");
    if (cap?.kind !== "text" || bodyText?.kind !== "text") return;
    const marginLeft = 1440 / 15;
    // Letter sits AT the text margin (in the column), body indented past it.
    expect(cap.x).toBeCloseTo(marginLeft, 0);
    expect(bodyText.x).toBeGreaterThan(marginLeft + cap.width - 1);
  });

  it("insets the wrap channel around a positioned frame by Word's default 6pt hSpace", () => {
    // probe2-dropcaps-frames p1: body text wrapping beside a w:framePr callout
    // starts 6pt (=8px) past the frame edge, not flush against it. Word applies
    // this default horizontal wrap distance when w:hSpace is absent; getting it
    // wrong widens the channel and lets a trailing word fit that Word wraps.
    const marginLeft = 1440 / 15; // 96px
    const frameWidthPx = 2000 / 15; // 2000 twips = 133.33px
    const render = (hSpaceAttr: string) => {
      const frame =
        `<w:p><w:pPr>` +
        `<w:framePr w:w="2000" w:h="1000" w:hRule="exact" w:wrap="around"` +
        ` w:vAnchor="paragraph" w:hAnchor="margin" w:x="0" w:y="0"${hSpaceAttr}/>` +
        `</w:pPr><w:r><w:t>Frame</w:t></w:r></w:p>`;
      const body =
        `<w:p><w:r><w:t xml:space="preserve">Body text that flows to the right of the positioned frame and keeps going for a while so it wraps onto several lines beside it.</w:t></w:r></w:p>`;
      const section =
        `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>` +
        `<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>`;
      const { result } = layout({
        "word/document.xml": wrapDocument(frame + body + section),
      });
      const bodyText = result.pages[0].items.find(
        (i) => i.kind === "text" && i.text === "Body",
      );
      expect(bodyText?.kind).toBe("text");
      return bodyText?.kind === "text" ? bodyText.x : NaN;
    };
    // Default (no hSpace): channel starts at frameRight + 6pt.
    const defaultGapPx = 6 * (96 / 72); // 8px
    expect(render("")).toBeCloseTo(marginLeft + frameWidthPx + defaultGapPx, 0);
    // Explicit hSpace overrides the default (240 twips = 16px).
    expect(render(` w:hSpace="240"`)).toBeCloseTo(marginLeft + frameWidthPx + 16, 0);
  });

  it("paints an exact-height frame's paragraph box to the frame height, not the content height", () => {
    // probe2-dropcaps-frames p1: the pull-quote frame (w:h=1000 w:hRule=exact,
    // two lines of text, pBdr box + shd fill) paints its fill and rules across
    // the whole 1000-twip frame, leaving empty shaded space below the text.
    // Web previously boxed only the laid content (~2 lines), losing the bottom
    // third of the frame's ink.
    const frameHPx = 1000 / 15; // 66.67px
    const frame =
      `<w:p><w:pPr>` +
      `<w:framePr w:w="3000" w:h="1000" w:hRule="exact" w:wrap="around"` +
      ` w:vAnchor="paragraph" w:hAnchor="margin" w:x="0" w:y="0"/>` +
      `<w:pBdr>` +
      `<w:top w:val="single" w:sz="6" w:space="4" w:color="BF8F00"/>` +
      `<w:left w:val="single" w:sz="6" w:space="4" w:color="BF8F00"/>` +
      `<w:bottom w:val="single" w:sz="6" w:space="4" w:color="BF8F00"/>` +
      `<w:right w:val="single" w:sz="6" w:space="4" w:color="BF8F00"/>` +
      `</w:pBdr>` +
      `<w:shd w:val="clear" w:color="auto" w:fill="FFF2CC"/>` +
      `<w:spacing w:after="0" w:line="240" w:lineRule="auto"/>` +
      `</w:pPr><w:r><w:t>Boxed frame</w:t></w:r></w:p>`;
    const body = `<w:p><w:r><w:t>Body paragraph after the frame.</w:t></w:r></w:p>`;
    const section =
      `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>` +
      `<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>`;
    const { result } = layout({
      "word/document.xml": wrapDocument(frame + body + section),
    });
    const items = result.pages[0].items;
    const fill = items.find((i) => i.kind === "rect" && i.fill?.toUpperCase().includes("FFF2CC"));
    expect(fill?.kind).toBe("rect");
    if (fill?.kind !== "rect") return;
    // The shading (border-box interior) grows far beyond one line (~19px):
    // frame height minus the two border pads (space 4pt + rule ~= 6.33px each).
    expect(fill.height).toBeGreaterThan(frameHPx - 16);
    // Border-to-border the painted box spans the full frame height: the bottom
    // rule sits ~frameH below the top rule, and the side rules span the fill.
    const edges = items.filter((i) => i.kind === "edge" && i.border?.color?.toUpperCase().includes("BF8F00"));
    expect(edges.length).toBe(4);
    const rules = edges.filter((e) => e.kind === "edge" && Math.abs(e.y1 - e.y2) < 0.01);
    expect(rules.length).toBe(2);
    if (rules[0]?.kind === "edge" && rules[1]?.kind === "edge") {
      expect(Math.abs(Math.abs(rules[1].y1 - rules[0].y1) - frameHPx)).toBeLessThan(3);
    }
    const sides = edges.filter((e) => e.kind === "edge" && Math.abs(e.x1 - e.x2) < 0.01);
    expect(sides.length).toBe(2);
    for (const s of sides) {
      if (s.kind === "edge") expect(s.y2 - s.y1).toBeGreaterThan(frameHPx - 16);
    }
  });

  it("insets a bordered run (w:bdr) by rule width + space at segment and line starts", () => {
    // probe2-run-borders p1: Word reserves the painted rule width + w:space
    // horizontally at each bordered-segment boundary and re-insets EVERY
    // wrapped line of the run (each visual line starts 1.00pt past the margin
    // for sz=8). Skipping the inset let one extra word fit per line, shifting
    // every subsequent wrap point.
    const bordered =
      `<w:p><w:r><w:rPr><w:bdr w:val="single" w:sz="8" w:space="0" w:color="7030A0"/></w:rPr>` +
      `<w:t xml:space="preserve">This is a single bordered run whose text is long enough that it must wrap across two or three lines and every line re-insets from the margin by the border pad.</w:t></w:r></w:p>`;
    const section =
      `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>` +
      `<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>`;
    const { result } = layout({
      "word/document.xml": wrapDocument(bordered + section),
    });
    const items = result.pages[0].items.filter((i) => i.kind === "text" && i.text.trim().length > 0);
    const marginLeft = 1440 / 15; // 96px
    const padPx = (8 / 8) * (96 / 72); // sz=8 -> 1pt rule, space=0 -> 1.333px
    // Rows by y: every line's first glyph is inset by the pad, not at the margin.
    const rows = new Map<number, number>();
    for (const it of items) {
      if (it.kind !== "text") continue;
      const y = Math.round(it.baseline);
      rows.set(y, Math.min(rows.get(y) ?? Infinity, it.x));
    }
    const starts = [...rows.entries()].sort((a, b) => a[0] - b[0]).map(([, x]) => x);
    expect(starts.length).toBeGreaterThan(1); // it wraps
    for (const s of starts) {
      expect(s).toBeCloseTo(marginLeft + padPx, 0);
    }
  });

  it("renders U+00AD soft hyphens as visible non-breaking hyphens", () => {
    // probe2-hyphenation p1: Word paints an optional hyphen as a hyphen glyph
    // in EVERY position and never breaks a line at it — the soft-hyphenated
    // word moves whole. Mapping to "-" (a hyphenBreaks opportunity) split the
    // word where Word kept it together, reflowing the paragraph.
    const para =
      `<w:p><w:r><w:t xml:space="preserve">start super­cali­fragilistic end</w:t></w:r></w:p>`;
    const section =
      `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>` +
      `<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>`;
    const { result } = layout({ "word/document.xml": wrapDocument(para + section) });
    const texts = result.pages[0].items.filter((i) => i.kind === "text");
    const joined = texts.map((t) => (t.kind === "text" ? t.text : "")).join("");
    // Soft hyphens paint as visible U+2011 hyphens, not invisible U+00AD.
    expect(joined).toContain("super‑cali‑fragilistic");
    expect(joined).not.toContain("­");
  });

  it("fits a w:fitText run into exactly its target width", () => {
    // probe3-text-effects: <w:fitText w:val="1440"/> compresses "SQUEEZE ME
    // INTO ONE INCH" so the run spans exactly 1in (Word PDF: glyph extent
    // 71.2pt inside the 72pt box). Unimplemented, the run painted at natural
    // width and overhung Word's by ~70px.
    const para =
      `<w:p><w:r><w:t xml:space="preserve">Label: </w:t></w:r>` +
      `<w:r><w:rPr><w:fitText w:val="1440" w:id="90"/></w:rPr>` +
      `<w:t xml:space="preserve">SQUEEZE ME INTO ONE INCH</w:t></w:r></w:p>`;
    const section =
      `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>` +
      `<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>`;
    const { result } = layout({ "word/document.xml": wrapDocument(para + section) });
    const texts = result.pages[0].items.filter(
      (i): i is Extract<(typeof result.pages)[0]["items"][0], { kind: "text" }> => i.kind === "text",
    );
    const first = texts.find((t) => t.text === "SQUEEZE");
    const last = texts.find((t) => t.text === "INCH");
    expect(first).toBeDefined();
    expect(last).toBeDefined();
    if (!first || !last) return;
    const span = last.x + last.width - first.x;
    expect(span).toBeCloseTo(1440 / 15, 0); // 96px = 1in
  });

  it("applies the final full-width banner's spacing-after before column body text", () => {
    const render = (authorsAfter: number) => {
      const frame = `<w:framePr w:w="9360" w:wrap="notBeside" w:hAnchor="text" w:vAnchor="text" w:xAlign="center"/>`;
      const filler = Array.from({ length: 70 }, (_, i) => p(`Body line ${i}`)).join("");
      const body =
        `<w:p><w:pPr><w:pStyle w:val="Title"/><w:framePr w:h="360" w:hRule="exact"/></w:pPr>` +
        `<w:r><w:t>Title</w:t></w:r></w:p>` +
        `<w:p><w:pPr><w:pStyle w:val="Authors"/><w:framePr w:h="781" w:hRule="exact" w:y="-213"/></w:pPr>` +
        `<w:r><w:t>Authors</w:t></w:r></w:p>` +
        `<w:p><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr>` +
        `<w:r><w:t>Abstract</w:t></w:r></w:p>` +
        filler +
        `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>` +
        `<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>` +
        `<w:cols w:num="2" w:space="720"/></w:sectPr>`;
      return layout({
        "word/document.xml": wrapDocument(body),
        "word/styles.xml": `<?xml version="1.0"?>
          <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
              <w:rPr><w:sz w:val="24"/></w:rPr>
            </w:style>
            <w:style w:type="paragraph" w:styleId="Title">
              <w:basedOn w:val="Normal"/>
              <w:pPr>${frame}<w:spacing w:after="480"/></w:pPr>
            </w:style>
            <w:style w:type="paragraph" w:styleId="Authors">
              <w:basedOn w:val="Normal"/>
              <w:pPr>${frame}<w:spacing w:after="${authorsAfter}"/></w:pPr>
            </w:style>
          </w:styles>`,
      }).result;
    };
    const item = (result: ReturnType<typeof layoutDocument>, text: string) => {
      const found = result.pages[0].items.find(
        (candidate) => candidate.kind === "text" && candidate.text === text,
      );
      if (found?.kind !== "text") throw new Error(`${text} missing`);
      return found;
    };

    const withoutAfter = render(0);
    const withAfter = render(320);
    const titleWithout = item(withoutAfter, "Title");
    const authorsWithout = item(withoutAfter, "Authors");
    const secondColumnTop = (result: ReturnType<typeof layoutDocument>) =>
      Math.min(
        ...result.pages[0].items
          .filter(
            (candidate) =>
              candidate.kind === "text" && candidate.x > 400,
          )
          .map((candidate) => (candidate.kind === "text" ? candidate.lineTop : Infinity)),
      );

    // Consecutive frames remain one stacked banner: the title's 480-twip
    // spacing-after does not open a gap before Authors.
    expect(authorsWithout.lineTop - titleWithout.lineTop).toBeCloseTo(360 / 15, 3);
    expect(item(withAfter, "Title").lineTop).toBeCloseTo(titleWithout.lineTop, 3);
    expect(item(withAfter, "Authors").lineTop).toBeCloseTo(authorsWithout.lineTop, 3);
    // Only the final banner paragraph's 320-twip after-spacing advances the
    // following normal-flow cursor, once.
    expect(item(withAfter, "Abstract").lineTop - item(withoutAfter, "Abstract").lineTop).toBeCloseTo(
      320 / 15,
      3,
    );
    // The shared column band excludes paragraph spacing: column two restarts
    // immediately below Authors and does not pay the first cursor's gap again.
    expect(secondColumnTop(withAfter)).toBeCloseTo(secondColumnTop(withoutAfter), 3);
    expect(secondColumnTop(withAfter) - authorsWithout.lineTop).toBeCloseTo(781 / 15, 3);
  });

  it("uses only whole-line space above a full-width banner in later columns", () => {
    const render = (prefixHeight: number) => {
      const line = (text: string) =>
        `<w:r><w:t>${text}</w:t><w:br/></w:r>`;
      const body =
        `<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="${prefixHeight}" w:lineRule="exact"/></w:pPr>` +
        `<w:r><w:t>Prefix</w:t></w:r></w:p>` +
        `<w:p><w:pPr><w:framePr w:w="3600" w:h="600" w:hRule="exact" w:wrap="notBeside" ` +
        `w:hAnchor="text" w:vAnchor="text" w:xAlign="center"/></w:pPr>` +
        `<w:r><w:t>Banner</w:t></w:r></w:p>` +
        `<w:p><w:pPr><w:widowControl w:val="0"/><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="exact"/></w:pPr>` +
        Array.from({ length: 14 }, (_, i) => line(`L${i + 1}`)).join("") +
        `<w:r><w:t>L15</w:t></w:r></w:p>` +
        `<w:sectPr><w:pgSz w:w="6000" w:h="4000"/>` +
        `<w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720"/>` +
        `<w:cols w:num="2" w:space="240"/></w:sectPr>`;
      return layout({ "word/document.xml": wrapDocument(body) }).result;
    };
    const item = (result: ReturnType<typeof layoutDocument>, text: string) => {
      const found = result.pages[0].items.find(
        (candidate) => candidate.kind === "text" && candidate.text === text,
      );
      if (found?.kind !== "text") throw new Error(`${text} missing`);
      return found;
    };

    const oneLineGap = render(240);
    const shortGap = render(180);
    const oneLineL8 = item(oneLineGap, "L8");
    const oneLineL9 = item(oneLineGap, "L9");
    const shortL8 = item(shortGap, "L8");

    expect(oneLineL8.x).toBeGreaterThan(200);
    expect(oneLineL8.lineTop).toBeCloseTo(oneLineGap.pages[0].bodyTop, 3);
    expect(oneLineL9.lineTop).toBeCloseTo(item(oneLineGap, "Banner").lineTop + 600 / 15, 3);
    // The pre-banner line consumes column-flow capacity; it does not buy an
    // extra line at the page bottom.
    expect(pageText(oneLineGap, 0)).toContain("L14");
    expect(pageText(oneLineGap, 0)).not.toContain("L15");
    expect(pageText(oneLineGap, 1)).toContain("L15");
    // Twelve pixels remain before the banner, less than one 16px line, so the
    // first continuation line jumps directly below it.
    expect(shortL8.x).toBeGreaterThan(200);
    expect(shortL8.lineTop).toBeCloseTo(item(shortGap, "Banner").lineTop + 600 / 15, 3);
    expect(pageText(shortGap, 0)).toContain("L14");
    expect(pageText(shortGap, 0)).not.toContain("L15");
    expect(pageText(shortGap, 1)).toContain("L15");
  });

  it("keeps ordinary positioned-frame spacing out of the normal spacing chain", () => {
    const render = (after: number) => {
      const body =
        `<w:p><w:pPr><w:spacing w:after="${after}"/>` +
        `<w:framePr w:w="3000" w:h="600" w:hRule="exact" w:wrap="notBeside"/></w:pPr>` +
        `<w:r><w:t>Frame</w:t></w:r></w:p>` +
        `<w:p><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr>` +
        `<w:r><w:t>Body</w:t></w:r></w:p>` +
        `<w:sectPr><w:pgSz w:w="6000" w:h="6000"/>` +
        `<w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720"/></w:sectPr>`;
      return layout({ "word/document.xml": wrapDocument(body) }).result;
    };
    const bodyTop = (after: number) => {
      const found = render(after).pages[0].items.find(
        (candidate) => candidate.kind === "text" && candidate.text === "Body",
      );
      if (found?.kind !== "text") throw new Error("Body missing");
      return found.lineTop;
    };

    expect(bodyTop(320)).toBeCloseTo(bodyTop(0), 3);
  });

  it("resolves PAGE and NUMPAGES fields in footers per page", () => {
    const paras = Array.from({ length: 120 }, (_, i) => p(`Paragraph ${i}`)).join("");
    const { result } = layout({
      "word/document.xml": wrapDocument(
        paras +
          `<w:sectPr>
            <w:footerReference xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" w:type="default" r:id="rIdF"/>
            <w:pgSz w:w="12240" w:h="15840"/>
            <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/>
          </w:sectPr>`,
      ),
      "word/_rels/document.xml.rels": `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdF" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>
</Relationships>`,
      "word/footer1.xml": `<?xml version="1.0"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p>
    <w:r><w:t xml:space="preserve">Page </w:t></w:r>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
    <w:r><w:t xml:space="preserve"> of </w:t></w:r>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText xml:space="preserve"> NUMPAGES </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
  </w:p>
</w:ftr>`,
    });
    const total = result.totalPages;
    expect(total).toBeGreaterThan(1);
    for (let i = 0; i < total; i++) {
      const text = pageText(result, i);
      expect(text).toContain(`Page ${i + 1} of ${total}`);
    }
  });

  it("overlays a widthless centered PAGE footer frame only for single-digit pages", () => {
    const body = Array.from({ length: 10 }, (_, i) =>
      `<w:p><w:r><w:t>Body ${i + 1}</w:t>${i < 9 ? '<w:br w:type="page"/>' : ""}</w:r></w:p>`,
    ).join("");
    const { result } = layout({
      "word/document.xml": wrapDocument(
        body +
          `<w:sectPr>
            <w:footerReference xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" w:type="default" r:id="rIdF"/>
            <w:pgSz w:w="12240" w:h="15840"/>
            <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:footer="720"/>
          </w:sectPr>`,
      ),
      "word/_rels/document.xml.rels": `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdF" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>
</Relationships>`,
      "word/footer1.xml": `<?xml version="1.0"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p><w:pPr><w:framePr w:wrap="around" w:vAnchor="text" w:hAnchor="margin" w:xAlign="center" w:y="1"/></w:pPr>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>PAGE</w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
  </w:p>
  <w:p><w:r><w:ptab w:relativeTo="margin" w:alignment="left" w:leader="none"/></w:r></w:p>
  <w:p><w:r><w:t>Footer label</w:t></w:r></w:p>
  <w:p><w:r><w:tab/></w:r></w:p>
</w:ftr>`,
    });
    expect(result.totalPages).toBe(10);
    const pageNumber = (page: number) => {
      const laidPage = result.pages[page - 1];
      const item = laidPage.items.slice(laidPage.hfStart).find(
        (candidate) => candidate.kind === "text" && candidate.text === String(page),
      );
      if (item?.kind !== "text") throw new Error(`missing page ${page} footer`);
      return item;
    };
    const footerLabel = (page: number) => {
      const laidPage = result.pages[page - 1];
      const item = laidPage.items.slice(laidPage.hfStart).find(
        (candidate) => candidate.kind === "text" && candidate.text === "Footer",
      );
      if (item?.kind !== "text") throw new Error(`missing page ${page} label`);
      return item;
    };
    // The frame is out of flow: the empty ptab follower shares its band and
    // the label lays one line below the number on EVERY page. But the frame
    // still reserves its own line in the footer HEIGHT once its painted text
    // is wider than its glyph box (two or more digits) — measured from the
    // NIH reference PDF over all 419 pages: footer top = pageBottom −
    // footerDist − 3 lines on the single-digit pages 1-9 and − 4 lines from
    // page 10 on, while the painted stack (number, admin line directly
    // below) never changes shape.
    for (let page = 2; page <= 9; page++) {
      expect(result.pages[page - 1].bodyBottom).toBeCloseTo(result.pages[0].bodyBottom, 3);
      expect(pageNumber(page).lineTop).toBeCloseTo(pageNumber(1).lineTop, 3);
      expect(footerLabel(page).lineTop).toBeCloseTo(footerLabel(1).lineTop, 3);
    }
    // Page 10 ("10" is wider than the glyph box): one extra line of footer
    // height — the whole footer (and the body bottom) moves up by one line,
    // but the label stays exactly one line below the number.
    const lineH = footerLabel(1).lineTop - pageNumber(1).lineTop;
    expect(lineH).toBeGreaterThan(8);
    expect(result.pages[9].bodyBottom).toBeCloseTo(result.pages[0].bodyBottom - lineH, 3);
    expect(pageNumber(10).lineTop).toBeCloseTo(pageNumber(1).lineTop - lineH, 3);
    expect(footerLabel(10).lineTop - pageNumber(10).lineTop).toBeCloseTo(lineH, 3);
    // xAlign=center positions the number at the margin-box center.
    const n1 = pageNumber(1);
    const contentWidth = (12240 - 1440 - 1440) / 15;
    expect(n1.x + n1.width / 2).toBeCloseTo(96 + contentWidth / 2, 0);
  });

  it("stacks a PAGE footer frame above a following line whose extent collides", () => {
    const { result } = layout({
      "word/document.xml": wrapDocument(
        `<w:p><w:r><w:t>Body</w:t></w:r></w:p>` +
          `<w:sectPr>
            <w:footerReference xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" w:type="default" r:id="rIdF"/>
            <w:pgSz w:w="12240" w:h="15840"/>
            <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:footer="720"/>
          </w:sectPr>`,
      ),
      "word/_rels/document.xml.rels": `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdF" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>
</Relationships>`,
      // NIH pattern: centered PAGE frame + centered admin line — the text's
      // natural extent crosses the frame box, so it wraps BELOW the number.
      "word/footer1.xml": `<?xml version="1.0"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p><w:pPr><w:framePr w:wrap="around" w:vAnchor="text" w:hAnchor="margin" w:xAlign="center" w:y="1"/></w:pPr>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>PAGE</w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
  </w:p>
  <w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>Centered admin footer line 01-23-4567</w:t></w:r></w:p>
</w:ftr>`,
    });
    const page = result.pages[0];
    const hf = page.items.slice(page.hfStart);
    const num = hf.find((it) => it.kind === "text" && it.text === "1");
    const admin = hf.find((it) => it.kind === "text" && it.text.includes("Centered"));
    if (num?.kind !== "text" || admin?.kind !== "text") throw new Error("missing footer items");
    expect(admin.lineTop).toBeGreaterThan(num.lineTop + 1);
  });

  it("glues NBSP-flanked spaces into one wrap unit; plain multi-spaces still break", () => {
    // Word does not break inside a whitespace cluster that TOUCHES an NBSP:
    // the flanking words wrap together (NIH p106 Hunogigu+NBSP+space+Durirone,
    // p383 "gedubid the"+NBSP+space+[underlined fill-in] -- both NBSP+space in
    // the XML, measured against the Word PDF). PLAIN runs of spaces remain
    // ordinary break opportunities -- gluing them regressed interactive
    // typing (a space typed at a wrap boundary dragged the previous word and
    // the caret to the next line).
    const narrow =
      `<w:sectPr><w:pgSz w:w="4000" w:h="15840"/>` +
      `<w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720"/></w:sectPr>`;
    const lineTexts = (xml: string): string[][] => {
      const { result } = layout({ "word/document.xml": wrapDocument(xml + narrow) });
      const byLine = new Map<number, string[]>();
      for (const it of result.pages[0].items) {
        if (it.kind !== "text" || !it.text.trim()) continue;
        const key = Math.round(it.baseline);
        byLine.set(key, [...(byLine.get(key) ?? []), it.text]);
      }
      return [...byLine.entries()].sort((a, b) => a[0] - b[0]).map(([, t]) => t);
    };
    // Plain double space: an ordinary break opportunity -- "zz." stays on the
    // upper line while the long word wraps alone.
    const plain = lineTexts(
      `<w:p><w:r><w:t xml:space="preserve">aa bb cc dd ee ff gg hh iii." Xxxx yy zz.  Qqqqqqqqqqqqqq rr</w:t></w:r></w:p>`,
    );
    const zzLine = plain.find((l) => l.some((t) => t.includes("zz.")));
    expect(zzLine ?? []).not.toContain("Qqqqqqqqqqqqqq");
    // NBSP at the end of the word glues the following plain space too: "zz."
    // moves down WITH the long word as one unit.
    const nbsp = lineTexts(
      `<w:p><w:r><w:t xml:space="preserve">aa bb cc dd ee ff gg hh iii." Xxxx yy zz.  Qqqqqqqqqqqqqq rr</w:t></w:r></w:p>`,
    );
    const zzNbsp = nbsp.find((l) => l.some((t) => t.includes("zz.")));
    expect(zzNbsp ?? []).toContain("Qqqqqqqqqqqqqq");
  });

  it("moves a keepNext chain whole when its 3-line terminator cannot split", () => {
    // widow+orphan make a 2-3 line terminator unsplittable, so a keepNext
    // heading whose successor must move wholesale moves with it (NIH
    // p416/417: '537' keepNext + heading + 3-line URL paragraph relocate
    // together although the first two terminator lines would fit).
    const filler = Array.from({ length: 47 }, (_, i) => p(`filler line ${i + 1}`)).join("");
    const term = Array.from({ length: 19 }, () => "wordy content").join(" ");
    const { result } = layout({
      "word/document.xml": wrapDocument(
        filler +
          `<w:p><w:pPr><w:keepNext/></w:pPr><w:r><w:t>537</w:t></w:r></w:p>` +
          `<w:p><w:pPr><w:keepNext/></w:pPr><w:r><w:t>Heading line</w:t></w:r></w:p>` +
          `<w:p><w:r><w:t xml:space="preserve">${term}</w:t></w:r></w:p>` +
          `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>` +
          `<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>`,
      ),
    });
    const pageOf = (needle: string): number =>
      result.pages.findIndex((pg) => pg.items.some((it) => it.kind === "text" && it.text.includes(needle)));
    const termPage = pageOf("wordy");
    expect(termPage).toBe(1); // the terminator itself moved wholesale
    // The chain head must sit on the same page as the terminator paragraph.
    expect(pageOf("537")).toBe(termPage);
    expect(pageOf("Heading")).toBe(termPage);
  });

  it("computes numbering labels with restarts", () => {
    const numberingXml = `<?xml version="1.0"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="0">
    <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/></w:lvl>
    <w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%2)"/><w:lvlJc w:val="left"/></w:lvl>
  </w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
</w:numbering>`;
    const numPara = (text: string, ilvl: number) =>
      `<w:p><w:pPr><w:numPr><w:ilvl w:val="${ilvl}"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`;
    const { result } = layout({
      "word/document.xml": wrapDocument(
        numPara("one", 0) + numPara("two", 0) + numPara("sub-a", 1) + numPara("sub-b", 1) + numPara("three", 0) + numPara("sub-restart", 1),
      ),
      "word/numbering.xml": numberingXml,
    });
    const text = pageText(result, 0);
    expect(text).toContain("1.");
    expect(text).toContain("2.");
    expect(text).toContain("a)");
    expect(text).toContain("b)");
    expect(text).toContain("3.");
    // After returning to level 0, level 1 restarts at "a)"
    const idxThree = text.indexOf("3.");
    expect(text.slice(idxThree)).toContain("a)");
  });

  it("shares one counter per abstractNum across instances; startOverride restarts once", () => {
    // Word keys numbering state by ABSTRACT definition: an instance whose
    // lvlOverride merely redefines the level (no startOverride) continues the
    // running counter (phase23: Heading1 hops numId 71 -> 77 -> 74 and Word
    // numbers straight through 1..11). Only a startOverride restarts, the
    // first time that instance is referenced.
    const numberingXml = `<?xml version="1.0"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="0">
    <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/></w:lvl>
  </w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
  <w:num w:numId="2"><w:abstractNumId w:val="0"/>
    <w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride>
  </w:num>
  <w:num w:numId="3"><w:abstractNumId w:val="0"/>
    <w:lvlOverride w:ilvl="0">
      <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/></w:lvl>
    </w:lvlOverride>
  </w:num>
</w:numbering>`;
    const numPara = (text: string, numId: number) =>
      `<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="${numId}"/></w:numPr></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`;
    const { result } = layout({
      "word/document.xml": wrapDocument(
        numPara("a", 1) + // 1.
          numPara("b", 1) + // 2.
          numPara("c", 2) + // startOverride: restart -> 1.
          numPara("d", 1) + // shared counter continues -> 2.
          numPara("e", 3) + // lvl redefinition, NO startOverride -> 3.
          numPara("f", 2), // numId 2 already referenced: no second restart -> 4.
      ),
      "word/numbering.xml": numberingXml,
    });
    const labels = result.pages[0].items
      .filter((i) => i.kind === "text" && /^\d+\.$/.test(i.text))
      .map((i) => (i.kind === "text" ? i.text : ""));
    expect(labels).toEqual(["1.", "2.", "1.", "2.", "3.", "4."]);
  });

  it("extends paragraph shading/borders over the hanging-indent numbering label", () => {
    // Word anchors paragraph decoration at the paragraph's leftmost text
    // extent: with ind left=432 hanging=432 the numbering label sits at the
    // margin INSIDE the shaded box (phase23 Heading1's blue banner shows
    // "4<tab>TITLE" inside the full-width fill).
    const numberingXml = `<?xml version="1.0"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="0">
    <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1"/><w:lvlJc w:val="left"/>
      <w:pPr><w:ind w:left="432" w:hanging="432"/></w:pPr>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
</w:numbering>`;
    const heading =
      `<w:p><w:pPr>` +
      `<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>` +
      `<w:shd w:val="clear" w:fill="4F81BD"/>` +
      `</w:pPr><w:r><w:t>BANNER</w:t></w:r></w:p>`;
    const { result } = layout({
      "word/document.xml": wrapDocument(heading + p("body")),
      "word/numbering.xml": numberingXml,
    });
    const items = result.pages[0].items;
    const rect = items.find((i) => i.kind === "rect" && i.fill.toUpperCase() === "#4F81BD");
    const label = items.find((i) => i.kind === "text" && i.text === "1");
    const body = items.find((i) => i.kind === "text" && i.text === "body");
    if (rect?.kind !== "rect" || label?.kind !== "text" || body?.kind !== "text") {
      throw new Error("items not found");
    }
    // The label starts at the hanging outdent = the plain body margin, and
    // the shading box reaches back to enclose it.
    expect(label.x).toBeCloseTo(body.x, 1);
    expect(rect.x).toBeLessThanOrEqual(label.x + 0.01);
  });

  it("lets a style's own ind beat the numbering level's for style-sourced numbering", () => {
    // phase23 Heading3: the style carries ind left=720 while the linked
    // numbering level says left=4410 hanging=720. Word paints the number at
    // the margin (style left wins attribute-wise; the level's hanging
    // survives because the style sets none). A DIRECT numPr keeps the
    // opposite precedence (level ind beats the style chain).
    const numberingXml = `<?xml version="1.0"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="0">
    <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1"/><w:lvlJc w:val="left"/>
      <w:pPr><w:ind w:left="4410" w:hanging="720"/></w:pPr>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
</w:numbering>`;
    const stylesXml = `<?xml version="1.0"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:styleId="H3">
    <w:name w:val="H3"/>
    <w:pPr>
      <w:numPr><w:numId w:val="1"/></w:numPr>
      <w:ind w:left="720"/>
    </w:pPr>
  </w:style>
</w:styles>`;
    const heading = `<w:p><w:pPr><w:pStyle w:val="H3"/></w:pPr><w:r><w:t>TITLE</w:t></w:r></w:p>`;
    const { result } = layout({
      "word/document.xml": wrapDocument(heading + p("body")),
      "word/numbering.xml": numberingXml,
      "word/styles.xml": stylesXml,
    });
    const items = result.pages[0].items;
    const label = items.find((i) => i.kind === "text" && i.text === "1");
    const body = items.find((i) => i.kind === "text" && i.text === "body");
    if (label?.kind !== "text" || body?.kind !== "text") throw new Error("items not found");
    // left=720 (style) with hanging=720 (level): the number sits at the margin.
    expect(label.x).toBeCloseTo(body.x, 1);
  });

  it("lays out table rows and repeats content within page", () => {
    const rows = Array.from(
      { length: 3 },
      (_, i) => `<w:tr>
        <w:tc><w:p><w:r><w:t>R${i}C0</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>R${i}C1</w:t></w:r></w:p></w:tc>
      </w:tr>`,
    ).join("");
    const { result } = layout({
      "word/document.xml": wrapDocument(
        `<w:tbl><w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>${rows}</w:tbl>` + p("after table"),
      ),
    });
    const text = pageText(result, 0);
    for (let i = 0; i < 3; i++) {
      expect(text).toContain(`R${i}C0`);
      expect(text).toContain(`R${i}C1`);
    }
    // Cells in the same row share a top; columns are offset horizontally.
    const c0 = result.pages[0].items.find((it) => it.kind === "text" && it.text.includes("R0C0"));
    const c1 = result.pages[0].items.find((it) => it.kind === "text" && it.text.includes("R0C1"));
    if (c0?.kind !== "text" || c1?.kind !== "text") throw new Error("cells not found");
    expect(Math.abs(c0.lineTop - c1.lineTop)).toBeLessThan(0.5);
    // No tcW anywhere: Word ignores the authored grid and autofits columns
    // to content (probe-tablegrid), so the second column hugs the first.
    expect(c1.x - c0.x).toBeGreaterThan(20);
    expect(c1.x - c0.x).toBeLessThan(90);
  });

  it("expands an autofit column to hold a tab layout (tabs are not shrink points)", () => {
    // staging-tblextreme: Word widens the L...R column from its authored
    // 2800tw to the full tab run (right stop 3200tw + end-of-cell mark);
    // treating the tab as a break opportunity instead left the column at
    // 2800tw and dropped "R" to its own line.
    const table = `<w:tbl>
      <w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>
      <w:tblGrid><w:gridCol w:w="2800"/><w:gridCol w:w="2800"/></w:tblGrid>
      <w:tr>
        <w:tc><w:tcPr><w:tcW w:w="2800" w:type="dxa"/></w:tcPr>
          <w:p><w:pPr><w:tabs><w:tab w:val="right" w:pos="3200"/></w:tabs></w:pPr>
            <w:r><w:t>L</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>R</w:t></w:r></w:p>
        </w:tc>
        <w:tc><w:tcPr><w:tcW w:w="2800" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>C2</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>`;
    const { result } = layout({ "word/document.xml": wrapDocument(table) });
    const items = result.pages[0].items;
    const l = items.find((it) => it.kind === "text" && it.text === "L");
    const r = items.find((it) => it.kind === "text" && it.text === "R");
    const c2 = items.find((it) => it.kind === "text" && it.text === "C2");
    if (l?.kind !== "text" || r?.kind !== "text" || c2?.kind !== "text") throw new Error("missing spans");
    // R stays on the same line as L, right-aligned at the 3200tw stop.
    expect(Math.abs(r.lineTop - l.lineTop)).toBeLessThan(0.5);
    // Column 1 grew past its authored 2800tw (186.7px) to hold the tab run:
    // C2 starts at ~3200tw + mark, not at 186.7px.
    expect(c2.x - l.x).toBeGreaterThan(205);
    expect(c2.x - l.x).toBeLessThan(230);
  });

  it("skips decimal tab stops for explicit tabs inside table cells", () => {
    // Measured in staging-tblextreme: tab + "12.5" with a decimal stop at
    // 2600tw lands LEFT-aligned on the next default stop (2880tw = 192px),
    // not decimal-aligned at 2600 - Word reserves in-cell decimal stops for
    // its automatic numeric alignment.
    const cellPara = `<w:p><w:pPr><w:tabs><w:tab w:val="decimal" w:pos="2600"/></w:tabs></w:pPr>
      <w:r><w:tab/></w:r><w:r><w:t>12.5</w:t></w:r></w:p>`;
    const table = `<w:tbl>
      <w:tblPr><w:tblLayout w:type="fixed"/></w:tblPr>
      <w:tblGrid><w:gridCol w:w="6000"/></w:tblGrid>
      <w:tr><w:tc><w:tcPr><w:tcW w:w="6000" w:type="dxa"/></w:tcPr>${cellPara}</w:tc></w:tr>
    </w:tbl>`;
    const { result } = layout({ "word/document.xml": wrapDocument(table) });
    const num = result.pages[0].items.find((it) => it.kind === "text" && it.text === "12.5");
    if (num?.kind !== "text") throw new Error("missing 12.5");
    // Page margin 96 + default stop 2880tw (192px) = 288, left-aligned.
    expect(num.x).toBeGreaterThan(286);
    expect(num.x).toBeLessThan(291);
  });

  it("confines an overrunning nested grid to its host cell at the grid's own ratio", () => {
    // staging-tblextreme's footnote table: a trusted [1400,1400] nested grid
    // in a narrower cell renders at the CELL width split 50/50 (Word scales
    // the authored grid), not autofit to each column's content.
    const nested = `<w:tbl>
      <w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>
      <w:tblGrid><w:gridCol w:w="2100"/><w:gridCol w:w="2100"/></w:tblGrid>
      <w:tr>
        <w:tc><w:tcPr><w:tcW w:w="2100" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>aa</w:t></w:r></w:p></w:tc>
        <w:tc><w:tcPr><w:tcW w:w="2100" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>bb</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>`;
    const outer = `<w:tbl>
      <w:tblPr><w:tblLayout w:type="fixed"/></w:tblPr>
      <w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid>
      <w:tr><w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>${nested}<w:p/></w:tc></w:tr>
    </w:tbl>`;
    const { result } = layout({ "word/document.xml": wrapDocument(outer) });
    const items = result.pages[0].items;
    const a = items.find((it) => it.kind === "text" && it.text === "aa");
    const b = items.find((it) => it.kind === "text" && it.text === "bb");
    if (a?.kind !== "text" || b?.kind !== "text") throw new Error("missing nested cells");
    // Host cell is 3000tw = 200px; the 4200tw grid is clamped to it and the
    // two equal grid columns stay equal: boundary at ~100px, not at "aa"'s
    // content width.
    expect(b.x - a.x).toBeGreaterThan(95);
    expect(b.x - a.x).toBeLessThan(105);
  });

  it("keeps a doubly-nested table's minimum when confining its parent (hard min)", () => {
    // staging-grid4 L2: the column holding the deeper table is pinned at that
    // table's minimum and the flexible text column absorbs the whole loss.
    const inner2 = `<w:tbl>
      <w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>
      <w:tblGrid><w:gridCol w:w="1500"/></w:tblGrid>
      <w:tr><w:tc><w:tcPr><w:tcW w:w="1500" w:type="dxa"/></w:tcPr>
        <w:p><w:r><w:t>mmmmmmmmmm</w:t></w:r></w:p></w:tc></w:tr>
    </w:tbl>`;
    const inner1 = `<w:tbl>
      <w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>
      <w:tblGrid><w:gridCol w:w="1800"/><w:gridCol w:w="1800"/></w:tblGrid>
      <w:tr>
        <w:tc><w:tcPr><w:tcW w:w="1800" w:type="dxa"/></w:tcPr>${inner2}<w:p/></w:tc>
        <w:tc><w:tcPr><w:tcW w:w="1800" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>side text here</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>`;
    const outer = `<w:tbl>
      <w:tblPr><w:tblLayout w:type="fixed"/></w:tblPr>
      <w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid>
      <w:tr><w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>${inner1}<w:p/></w:tc></w:tr>
    </w:tbl>`;
    const { result } = layout({ "word/document.xml": wrapDocument(outer) });
    const items = result.pages[0].items;
    const deep = items.find((it) => it.kind === "text" && it.text.startsWith("mmmm"));
    const side = items.find((it) => it.kind === "text" && it.text.startsWith("side"));
    if (deep?.kind !== "text" || side?.kind !== "text") throw new Error("missing cells");
    // Proportional confinement alone would put the boundary at 100px; the
    // wide unbreakable word (10 x 0.85em x 14.67px ~= 125px) pins col1 at the
    // deep table's minimum instead.
    expect(side.x - deep.x).toBeGreaterThan(115);
  });

  it("marks only table fills and rules with table paint roles", () => {
    const table = `<w:tbl>
      <w:tblPr><w:tblBorders>
        <w:top w:val="single"/><w:left w:val="single"/>
        <w:bottom w:val="single"/><w:right w:val="single"/>
      </w:tblBorders></w:tblPr>
      <w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>
      <w:tr><w:tc><w:tcPr><w:shd w:val="clear" w:fill="D9E2F3"/></w:tcPr>
        <w:p><w:r><w:t>cell</w:t></w:r></w:p>
      </w:tc></w:tr>
    </w:tbl>`;
    const paragraphBorder =
      `<w:p><w:pPr><w:pBdr><w:bottom w:val="single"/></w:pBdr></w:pPr>` +
      `<w:r><w:t>paragraph</w:t></w:r></w:p>`;
    const { result } = layout({
      "word/document.xml": wrapDocument(table + paragraphBorder),
    });
    const fills = result.pages[0].items.filter(
      (item) => item.kind === "rect" && item.fill.toUpperCase() === "#D9E2F3",
    );
    const edges = result.pages[0].items.filter((item) => item.kind === "edge");

    expect(fills).toHaveLength(1);
    expect(fills[0]?.kind === "rect" ? fills[0].role : undefined).toBe("table-fill");
    expect(edges.filter((edge) => edge.role === "table-rule")).toHaveLength(4);
    expect(edges.filter((edge) => edge.role === undefined)).toHaveLength(1);
  });

  it("paints style banding only when the style chain declares tblStyleRowBandSize", () => {
    // Word skips band1Horz/band2Horz shading when no style in the chain
    // declares an explicit w:tblStyleRowBandSize (every built-in banded style
    // writes w:val="1"), despite ECMA-376's nominal default of 1 — verified
    // against Word's output for staging-styles' CondGrid.
    const stylesXml = (tblPrExtra: string) => `<?xml version="1.0"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="table" w:styleId="Banded">
    <w:tblPr>${tblPrExtra}</w:tblPr>
    <w:tblStylePr w:type="band1Horz"><w:tcPr><w:shd w:val="clear" w:fill="D9E2F3"/></w:tcPr></w:tblStylePr>
  </w:style>
</w:styles>`;
    const row = (t: string) => `<w:tr><w:tc><w:p><w:r><w:t>${t}</w:t></w:r></w:p></w:tc></w:tr>`;
    const documentXml = wrapDocument(
      `<w:tbl><w:tblPr><w:tblStyle w:val="Banded"/>` +
        `<w:tblLook w:firstRow="1" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/></w:tblPr>` +
        `<w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>` +
        row("head") + row("r1") + row("r2") + row("r3") +
        `</w:tbl><w:p/>`,
    );
    const bandFills = (tblPrExtra: string) => {
      const { result } = layout({
        "word/document.xml": documentXml,
        "word/styles.xml": stylesXml(tblPrExtra),
      });
      return result.pages[0].items.filter(
        (item) => item.kind === "rect" && item.fill.toUpperCase() === "#D9E2F3",
      );
    };
    // No band size declared anywhere in the chain: no banding at all.
    expect(bandFills("")).toHaveLength(0);
    // Explicit band size 1: band1Horz hits body rows 1 and 3 (row 0 is the
    // header per tblLook firstRow, so band counting starts at row 1).
    expect(bandFills(`<w:tblStyleRowBandSize w:val="1"/>`)).toHaveLength(2);
  });

  it("rotates btLr cell text against declared and measured row heights", () => {
    const table = (rowHeight: string, ordinary: string) =>
      `<w:tbl><w:tblPr><w:tblLayout w:type="fixed"/><w:tblBorders>
          <w:top w:val="single" w:sz="4"/><w:left w:val="single" w:sz="4"/>
          <w:bottom w:val="single" w:sz="4"/><w:right w:val="single" w:sz="4"/>
          <w:insideV w:val="single" w:sz="4"/>
        </w:tblBorders></w:tblPr>
        <w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="300"/></w:tblGrid>
        <w:tr>${rowHeight}
          <w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>${ordinary}</w:tc>
          <w:tc><w:tcPr><w:tcW w:w="300" w:type="dxa"/><w:textDirection w:val="btLr"/><w:vAlign w:val="center"/></w:tcPr>
            <w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:sz w:val="16"/></w:rPr><w:t>VERTICAL</w:t></w:r></w:p>
          </w:tc>
        </w:tr>
      </w:tbl>`;
    const render = (rowHeight: string, ordinary: string) =>
      layout({
        "word/document.xml": wrapDocument(table(rowHeight, ordinary) + p("AFTER")),
      }).result;
    const declared = render(
      `<w:trPr><w:trHeight w:val="900"/></w:trPr>`,
      p("ORDINARY"),
    );
    const measured = render("", p("A") + p("B") + p("C") + p("D"));
    const textItem = (result: ReturnType<typeof layoutDocument>, text: string) => {
      const item = result.pages[0].items.find(
        (candidate) => candidate.kind === "text" && candidate.text === text,
      );
      if (item?.kind !== "text") throw new Error(`missing ${text}`);
      return item;
    };
    const rotatedBounds = (item: ReturnType<typeof textItem>) => {
      if (!item.rotate) throw new Error("missing rotation");
      const top = item.glyphTop ?? item.lineTop;
      const height = item.glyphBoxH ?? item.lineHeight;
      const pivotX = item.x + item.rotate.ox;
      const pivotY = top + item.rotate.oy;
      const radians = item.rotate.deg * Math.PI / 180;
      const points = [
        [item.x, top],
        [item.x + item.width, top],
        [item.x, top + height],
        [item.x + item.width, top + height],
      ].map(([x, y]) => ({
        x: pivotX + (x - pivotX) * Math.cos(radians) - (y - pivotY) * Math.sin(radians),
        y: pivotY + (x - pivotX) * Math.sin(radians) + (y - pivotY) * Math.cos(radians),
      }));
      const xs = points.map((point) => point.x);
      const ys = points.map((point) => point.y);
      const left = Math.min(...xs);
      const right = Math.max(...xs);
      const minTop = Math.min(...ys);
      const maxBottom = Math.max(...ys);
      return {
        left,
        right,
        top: minTop,
        bottom: maxBottom,
        width: right - left,
        height: maxBottom - minTop,
      };
    };
    const rowBounds = (result: ReturnType<typeof layoutDocument>) => {
      const edges = result.pages[0].items.filter((item) => item.kind === "edge");
      const xs = [...new Set(edges.filter((edge) => edge.x1 === edge.x2).map((edge) => edge.x1))].sort((a, b) => a - b);
      const ys = [...new Set(edges.filter((edge) => edge.y1 === edge.y2).map((edge) => edge.y1))].sort((a, b) => a - b);
      return { left: xs[xs.length - 2], right: xs[xs.length - 1], top: ys[0], bottom: ys[ys.length - 1] };
    };

    const declaredLabel = textItem(declared, "VERTICAL");
    const measuredLabel = textItem(measured, "VERTICAL");
    const declaredBox = rotatedBounds(declaredLabel);
    const measuredBox = rotatedBounds(measuredLabel);
    const declaredCell = rowBounds(declared);
    const measuredCell = rowBounds(measured);
    expect(declaredLabel.rotate?.deg).toBe(-90);
    expect(measuredLabel.rotate?.deg).toBe(-90);
    expect(declaredLabel.src?.t?.text).toBe("VERTICAL");
    expect(measuredLabel.src?.t?.text).toBe("VERTICAL");
    expect(textItem(declared, "ORDINARY").rotate).toBeUndefined();
    expect(declaredBox.width).toBeCloseTo(declaredLabel.lineHeight, 3);
    expect(declaredBox.height).toBeCloseTo(declaredLabel.width, 3);
    expect(declaredBox.height).toBeLessThan(60);
    expect(measuredBox.height).toBeCloseTo(measuredLabel.width, 3);
    for (const [box, cell] of [[declaredBox, declaredCell], [measuredBox, measuredCell]] as const) {
      expect(box.left).toBeGreaterThanOrEqual(cell.left);
      expect(box.right).toBeLessThanOrEqual(cell.right);
      expect(box.top).toBeGreaterThanOrEqual(cell.top);
      expect(box.bottom).toBeLessThanOrEqual(cell.bottom);
      expect(Math.abs((box.left + box.right - cell.left - cell.right) / 2)).toBeLessThan(0.5);
      expect(Math.abs((box.top + box.bottom - cell.top - cell.bottom) / 2)).toBeLessThan(0.5);
    }
    expect(textItem(declared, "AFTER").lineTop).toBeGreaterThanOrEqual(declaredCell.bottom);
    expect(textItem(measured, "AFTER").lineTop).toBeGreaterThanOrEqual(measuredCell.bottom);
  });

  it("autofits columns to content when the grid is a placeholder", () => {
    // Junk grid (100 twips per column, like generator output) + 100% width:
    // Word ignores the grid and sizes columns by content.
    const tbl = `<w:tbl>
      <w:tblPr><w:tblW w:type="pct" w:w="100%"/></w:tblPr>
      <w:tblGrid><w:gridCol w:w="100"/><w:gridCol w:w="100"/><w:gridCol w:w="100"/></w:tblGrid>
      <w:tr>
        <w:tc><w:p><w:r><w:t>Tiny</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>Mid col</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>A distinctly longer notes column with plenty of words</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>`;
    const { result } = layout({ "word/document.xml": wrapDocument(tbl + p("after")) });
    const cell = (t: string) => {
      const it = result.pages[0].items.find((i) => i.kind === "text" && i.text.includes(t));
      if (it?.kind !== "text") throw new Error(`cell ${t} not found`);
      return it;
    };
    const x0 = cell("Tiny").x;
    const x1 = cell("Mid").x;
    const x2 = cell("distinctly").x;
    const w0 = x1 - x0;
    const w1 = x2 - x1;
    // An even split would give each column ~1/3 of the content width
    // (~200px). Content-proportional sizing keeps the short columns narrow,
    // leaving the notes column with the bulk of the width.
    expect(w0).toBeLessThan(120);
    expect(w1).toBeLessThan(160);
    expect(w0 + w1).toBeLessThan(300);
    expect(w0).toBeLessThan(w1 + 40); // both stay near content size
  });

  it("autofits a tcW-less grid to content like Word", () => {
    // Word only trusts a grid it wrote itself, and it writes tcW on every
    // cell. A plausible-looking grid with no tcW is ignored: Word sizes the
    // "x" column to its content (5.75pt in the probe-tablegrid export).
    const tbl = `<w:tbl>
      <w:tblGrid><w:gridCol w:w="4680"/><w:gridCol w:w="4680"/></w:tblGrid>
      <w:tr>
        <w:tc><w:p><w:r><w:t>x</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>a much much much longer cell body here</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>`;
    const { result } = layout({ "word/document.xml": wrapDocument(tbl + p("after")) });
    const a = result.pages[0].items.find((i) => i.kind === "text" && i.text.includes("x"));
    const b = result.pages[0].items.find((i) => i.kind === "text" && i.text === "a");
    if (a?.kind !== "text" || b?.kind !== "text") throw new Error("cells not found");
    expect(b.x - a.x).toBeLessThan(40);
  });

  it("vertical (tbRl) section starts a fresh column when East Asian text resumes after a Western run", () => {
    // probe2-ruby-vertical p2: the body column ends right after the embedded
    // Latin "textDirection=tbRl" and the following CJK opens a new column,
    // even though the column is far from full. The short "日本ABC字" would
    // otherwise sit in one column; the grid-resync break splits 字 off.
    const body =
      `<w:p><w:r><w:t>日本ABC字</w:t></w:r></w:p>` +
      `<w:sectPr><w:textDirection w:val="tbRl"/>` +
      `<w:pgSz w:w="12240" w:h="15840"/>` +
      `<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>`;
    const { result } = layout({ "word/document.xml": wrapDocument(body) });
    const nichi = result.pages[0].items.find((i) => i.kind === "text" && i.text === "日");
    const ji = result.pages[0].items.find((i) => i.kind === "text" && i.text === "字");
    if (nichi?.kind !== "text" || ji?.kind !== "text") throw new Error("vertical CJK not found");
    // The section frame is laid horizontally then rotated +90° into columns, so
    // each pre-rotation LINE becomes a vertical column. Both chars start their
    // line (equal x); the resync put 字 on a second line, so it carries a
    // distinct baseline — a fresh column after rotation. Without the resync the
    // short run would share one line.
    expect(nichi.rotate?.deg).toBe(90);
    expect(Math.abs((ji.baseline ?? 0) - (nichi.baseline ?? 0))).toBeGreaterThan(10);
  });

  it("ruby distributeSpace spreads the annotation across the base width", () => {
    const rubyRun =
      `<w:r><w:ruby>` +
      `<w:rubyPr><w:hpsRaise w:val="11"/><w:rubyAlign w:val="distributeSpace"/></w:rubyPr>` +
      `<w:rt><w:r><w:rPr><w:sz w:val="8"/></w:rPr><w:t>かんじ</w:t></w:r></w:rt>` +
      `<w:rubyBase><w:r><w:rPr><w:sz w:val="22"/></w:rPr><w:t>漢字</w:t></w:r></w:rubyBase>` +
      `</w:ruby></w:r>`;
    const { result } = layout({ "word/document.xml": wrapDocument(`<w:p>${rubyRun}</w:p>`) });
    const rt = result.pages[0].items.find((i) => i.kind === "text" && i.text === "かんじ");
    if (rt?.kind !== "text") throw new Error("ruby annotation not found");
    // The 3-glyph annotation is spread over the 2-glyph base with a positive
    // inter-glyph gap (distributeSpace) rather than painted as a tight cluster.
    expect(rt.props.letterSpacing ?? 0).toBeGreaterThan(0);
  });

  it("honors a Word-authored grid (cells carry tcW)", () => {
    const tbl = `<w:tbl>
      <w:tblGrid><w:gridCol w:w="4680"/><w:gridCol w:w="4680"/></w:tblGrid>
      <w:tr>
        <w:tc><w:tcPr><w:tcW w:type="dxa" w:w="4680"/></w:tcPr><w:p><w:r><w:t>x</w:t></w:r></w:p></w:tc>
        <w:tc><w:tcPr><w:tcW w:type="dxa" w:w="4680"/></w:tcPr><w:p><w:r><w:t>a much much much longer cell body here</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>`;
    const { result } = layout({ "word/document.xml": wrapDocument(tbl + p("after")) });
    const a = result.pages[0].items.find((i) => i.kind === "text" && i.text.includes("x"));
    const b = result.pages[0].items.find((i) => i.kind === "text" && i.text === "a");
    if (a?.kind !== "text" || b?.kind !== "text") throw new Error("cells not found");
    // Equal 4680-twip columns: the second cell starts at the halfway point.
    expect(b.x - a.x).toBeGreaterThan(280);
  });

  it("expands an auto-width grid column to an indented paragraph's minimum", () => {
    const border = '<w:tblBorders><w:top w:val="single"/><w:left w:val="single"/>' +
      '<w:bottom w:val="single"/><w:right w:val="single"/><w:insideV w:val="single"/></w:tblBorders>';
    const tbl = `<w:tbl>
      <w:tblPr><w:tblW w:type="auto" w:w="0"/>${border}</w:tblPr>
      <w:tblGrid><w:gridCol w:w="1500"/><w:gridCol w:w="1500"/></w:tblGrid>
      <w:tr>
        <w:tc><w:tcPr><w:tcW w:type="dxa" w:w="1500"/></w:tcPr>
          <w:p><w:pPr><w:ind w:left="1440"/></w:pPr><w:r><w:t>inner</w:t></w:r></w:p></w:tc>
        <w:tc><w:tcPr><w:tcW w:type="dxa" w:w="1500"/></w:tcPr><w:p><w:r><w:t>adjacent</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>`;
    const { result } = layout({ "word/document.xml": wrapDocument(tbl) });
    const verticalXs = [...new Set(result.pages[0].items.flatMap((item) =>
      item.kind === "edge" && item.x1 === item.x2 ? [item.x1] : [],
    ))].sort((a, b) => a - b);

    expect(verticalXs).toHaveLength(3);
    expect(verticalXs[1] - verticalXs[0]).toBeGreaterThan(130);
    expect(verticalXs[2] - verticalXs[1]).toBeCloseTo(100, 1);
  });

  it("applies pgNumType start to display numbers", () => {
    const { result } = layout({
      "word/document.xml": wrapDocument(
        p("content") + `<w:sectPr><w:pgNumType w:start="5"/><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>`,
      ),
    });
    expect(result.pages[0].number).toBe(5);
  });

  it("adds a top border reserve OUTSIDE the before/after spacing collapse", () => {
    // Word collapses plain before/after (larger wins) but the bordered
    // paragraph's rule + space always push its text further down so the box
    // top clears the gap (wild-doerfp p31: H1 after=360 -> boxed Heading1
    // sits 18pt + 1.5pt below, not max(18pt, 1.5pt)).
    const { result } = layout({
      "word/document.xml": wrapDocument(
        `<w:p><w:pPr><w:spacing w:after="360"/></w:pPr><w:r><w:t>above</w:t></w:r></w:p>` +
          `<w:p><w:pPr><w:pBdr><w:top w:val="single" w:sz="4" w:space="1"/></w:pBdr></w:pPr><w:r><w:t>boxed</w:t></w:r></w:p>`,
      ),
    });
    const items = result.pages[0].items.filter((i) => i.kind === "text");
    const a = items.find((i) => i.kind === "text" && i.text === "above");
    const b = items.find((i) => i.kind === "text" && i.text === "boxed");
    if (a?.kind !== "text" || b?.kind !== "text") throw new Error("items missing");
    const gap = b.lineTop - (a.lineTop + a.lineHeight);
    // 360tw = 24px after + border reserve (1pt space = 4/3px + 0.75px rule).
    expect(gap).toBeCloseTo(24 + 4 / 3 + 0.75, 1);
  });

  it("adds a bottom border reserve outside the collapse against the next before", () => {
    const { result } = layout({
      "word/document.xml": wrapDocument(
        `<w:p><w:pPr><w:spacing w:after="240"/><w:pBdr><w:bottom w:val="single" w:sz="4" w:space="1"/></w:pBdr></w:pPr><w:r><w:t>boxed</w:t></w:r></w:p>` +
          `<w:p><w:pPr><w:spacing w:before="300"/></w:pPr><w:r><w:t>below</w:t></w:r></w:p>`,
      ),
    });
    const items = result.pages[0].items.filter((i) => i.kind === "text");
    const a = items.find((i) => i.kind === "text" && i.text === "boxed");
    const b = items.find((i) => i.kind === "text" && i.text === "below");
    if (a?.kind !== "text" || b?.kind !== "text") throw new Error("items missing");
    const gap = b.lineTop - (a.lineTop + a.lineHeight);
    // max(after 16px, before 20px) collapsed + the box's bottom reserve.
    expect(gap).toBeCloseTo(20 + 4 / 3 + 0.75, 1);
  });

  it("keeps the single collapsed reserve between merged identical-border paragraphs", () => {
    const bordered = (t: string) =>
      `<w:p><w:pPr><w:spacing w:before="0" w:after="0"/><w:pBdr><w:top w:val="single" w:sz="4" w:space="1"/><w:bottom w:val="single" w:sz="4" w:space="1"/></w:pBdr></w:pPr><w:r><w:t>${t}</w:t></w:r></w:p>`;
    const { result } = layout({
      "word/document.xml": wrapDocument(bordered("first") + bordered("second")),
    });
    const items = result.pages[0].items.filter((i) => i.kind === "text");
    const a = items.find((i) => i.kind === "text" && i.text === "first");
    const b = items.find((i) => i.kind === "text" && i.text === "second");
    if (a?.kind !== "text" || b?.kind !== "text") throw new Error("items missing");
    // Inside a merged box no rule paints between the paragraphs, but Word
    // still keeps ONE space+rule reserve of room (the top and bottom pads
    // collapse against each other, not add) — pre-existing calibrated
    // behavior (Alex Pickett cover RECIPIENT/ADDRESS block), preserved by
    // the outside-the-collapse reserve rule.
    expect(b.lineTop - (a.lineTop + a.lineHeight)).toBeCloseTo(4 / 3 + 0.75, 1);
  });

  it("anchors paint-routed CJK glyph boxes by the browser strut box", () => {
    // A measurer that reports the paint face's own (small) font box: the
    // renderer centers glyphs by that strut, so glyphTop must compensate to
    // land the baseline exactly (staging-eastasian: MS Mincho box is 1.0em
    // while the Hiragino line profile is 1.643em).
    const pb = { ascent: 10, descent: 2 };
    const m: TextMeasurer = {
      width: (t, f, ls) => measurer.width(t, f, ls),
      metrics: (f) => measurer.metrics(f),
      paintBox: () => pb,
    };
    const doc = DocxDocument.load(
      makeDocx({
        "word/document.xml": wrapDocument(
          `<w:p><w:r><w:rPr><w:rFonts w:eastAsia="MS Mincho"/></w:rPr><w:t>水は</w:t></w:r></w:p>`,
        ),
      }),
    );
    const result = layoutDocument(doc, { measurer: m });
    const cjk = result.pages[0].items.find(
      (i) => i.kind === "text" && i.text.includes("水") && i.font.paintFamily,
    );
    if (cjk?.kind !== "text") throw new Error("CJK item missing");
    const boxH = cjk.glyphBoxH!;
    // Browser centering: baseline = glyphTop + (boxH - pbBox)/2 + pbAsc = b.
    expect(cjk.glyphTop! + (boxH - pb.ascent - pb.descent) / 2 + pb.ascent).toBeCloseTo(
      cjk.baseline,
      3,
    );
  });

  it("draws a paragraph bottom border as a divider line", () => {
    const { result } = layout({
      "word/document.xml": wrapDocument(
        `<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="auto"/></w:pBdr></w:pPr><w:r><w:t>above the line</w:t></w:r></w:p>`,
      ),
    });
    const edges = result.pages[0].items.filter((i) => i.kind === "edge");
    expect(edges.length).toBe(1);
    if (edges[0].kind !== "edge") return;
    expect(Math.abs(edges[0].y1 - edges[0].y2)).toBeLessThan(0.01); // horizontal
  });

  it("keeps header text in the header zone and pushes body below it", () => {
    const { result, doc } = layout({
      "word/document.xml": wrapDocument(
        p("body text") +
          `<w:sectPr>
            <w:headerReference xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" w:type="default" r:id="rIdH"/>
            <w:pgSz w:w="12240" w:h="15840"/>
            <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/>
          </w:sectPr>`,
      ),
      "word/_rels/document.xml.rels": `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdH" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
</Relationships>`,
      "word/header1.xml": `<?xml version="1.0"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${p("HEADER")}</w:hdr>`,
    });
    const items = result.pages[0].items.filter((i) => i.kind === "text");
    const headerItem = items.find((i) => i.kind === "text" && i.text === "HEADER");
    const bodyItem = items.find((i) => i.kind === "text" && i.text.includes("body"));
    if (headerItem?.kind !== "text" || bodyItem?.kind !== "text") throw new Error("items missing");
    // Header sits at headerDistance (720 twips = 48px), above the body top (96px).
    expect(headerItem.lineTop).toBeGreaterThanOrEqual(47);
    expect(headerItem.lineTop).toBeLessThan(96);
    expect(bodyItem.lineTop).toBeGreaterThanOrEqual(95);
  });

  it("resolves margin-relative header anchors from the grown body top", () => {
    // wild2-med-phase23 (PDF-pinned): the tall logo header grows past the
    // 72pt top margin, and Word resolves the anchor's vRel="margin"
    // posOffset from the EFFECTIVE body top (header bottom, 129.77pt), not
    // the nominal margin — logo top 20.21pt = 129.77 - 109.5, on every page.
    const headerParas = ["H1", "H2", "H3", "H4"].map((t) => p(t)).join("");
    const anchor = `<w:p><w:r><w:drawing>
      <wp:anchor xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" behindDoc="0">
        <wp:positionH relativeFrom="margin"><wp:posOffset>-457200</wp:posOffset></wp:positionH>
        <wp:positionV relativeFrom="margin"><wp:posOffset>-457200</wp:posOffset></wp:positionV>
        <wp:extent cx="914400" cy="457200"/>
        <wp:wrapNone/>
        <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
            <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
              <pic:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rIdImg"/></pic:blipFill>
              <pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="457200"/></a:xfrm></pic:spPr>
            </pic:pic>
          </a:graphicData>
        </a:graphic>
      </wp:anchor>
    </w:drawing></w:r></w:p>`;
    const { result } = layout({
      "word/document.xml": wrapDocument(
        p("body text") +
          `<w:sectPr>
            <w:headerReference xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" w:type="default" r:id="rIdH"/>
            <w:pgSz w:w="12240" w:h="15840"/>
            <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/>
          </w:sectPr>`,
      ),
      "word/_rels/document.xml.rels": `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdH" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
</Relationships>`,
      "word/header1.xml": `<?xml version="1.0"?>
<w:hdr ${W_NS}>${anchor}${headerParas}</w:hdr>`,
      "word/_rels/header1.xml.rels": `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdImg" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/x.png"/>
</Relationships>`,
      "word/media/x.png": "PNGDATA",
    });
    const items = result.pages[0].items;
    const bodyItem = items.find((i) => i.kind === "text" && i.text.includes("body"));
    const img = items.find((i) => i.kind === "image");
    if (bodyItem?.kind !== "text" || img?.kind !== "image") throw new Error("items missing");
    // The 5-paragraph header grows the body top well past the 96px margin...
    expect(bodyItem.lineTop).toBeGreaterThan(110);
    // ...and the anchor's -36pt (-48px) margin offset resolves from THAT
    // grown top, not from the nominal 96px margin (which would put it at 48).
    expect(img.y).toBeCloseTo(bodyItem.lineTop - 48, 1);
  });

  it("gives the document-opening empty paragraph two slots under a grown header", () => {
    // PDF-pinned: wild2-med-phase23 p1 - under a header that outgrew the top
    // margin, the empty opener before a paragraph takes 2 x (line + after)
    // (first body baseline 179.05 = grown bodyTop 129.77 + 2 x 19.4 +
    // ascent), while every continuation page starts exactly at bodyTop.
    // Under a NORMAL header the same construct takes ONE slot
    // (wild-athabasca p1). wild2-legal p1 pins the sibling before-a-table
    // case (two mark lines, no after).
    const spacing = `<w:spacing w:before="0" w:after="120" w:line="240" w:lineRule="auto"/>`;
    const emptyOpener = `<w:p><w:pPr>${spacing}<w:rPr><w:sz w:val="22"/></w:rPr></w:pPr></w:p>`;
    const marker = `<w:p><w:pPr>${spacing}</w:pPr><w:r><w:rPr><w:sz w:val="22"/></w:rPr><w:t>MARKER</w:t></w:r></w:p>`;
    const sect = `<w:sectPr>
        <w:headerReference xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" w:type="default" r:id="rIdH"/>
        <w:pgSz w:w="12240" w:h="15840"/>
        <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/>
      </w:sectPr>`;
    const tallHeader = {
      "word/_rels/document.xml.rels": `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdH" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
</Relationships>`,
      "word/header1.xml": `<?xml version="1.0"?>
<w:hdr ${W_NS}>${["H1", "H2", "H3", "H4"].map((t) => p(t)).join("")}</w:hdr>`,
    };
    const find = (result: ReturnType<typeof layoutDocument>) => {
      const item = result.pages[0].items.find((i) => i.kind === "text" && i.text === "MARKER");
      if (item?.kind !== "text") throw new Error("missing MARKER");
      return item;
    };
    const withOpener = find(
      layout({ "word/document.xml": wrapDocument(emptyOpener + marker + sect), ...tallHeader }).result,
    );
    const bare = find(layout({ "word/document.xml": wrapDocument(marker + sect), ...tallHeader }).result);
    // after=120tw = 6pt = 8px; the empty opener contributes (line + after) TWICE.
    const slot = withOpener.lineHeight + 8;
    expect(withOpener.lineTop - bare.lineTop).toBeCloseTo(2 * slot, 1);
    // Without the grown header the opener takes a single slot (athabasca).
    const plainOpener = find(layout({ "word/document.xml": wrapDocument(emptyOpener + marker) }).result);
    const plainBare = find(layout({ "word/document.xml": wrapDocument(marker) }).result);
    expect(plainOpener.lineTop - plainBare.lineTop).toBeCloseTo(slot, 1);
  });
});

describe("pagination robustness", () => {
  it("terminates when orphan control cannot fit two lines on any page", () => {
    // Two lines of 400pt exact spacing (533px) on a US Letter body (~848px):
    // a lone first line at the bottom triggers the orphan push; after the
    // push the pair still cannot share a page. Must not loop forever.
    const body =
      p("filler before") +
      `<w:p><w:pPr><w:spacing w:line="8000" w:lineRule="exact"/></w:pPr>
        <w:r><w:t>first tall line</w:t><w:br/><w:t>second tall line</w:t></w:r>
      </w:p>` +
      p("after");
    const { result } = layout({ "word/document.xml": wrapDocument(body) });
    expect(result.totalPages).toBeGreaterThanOrEqual(2);
    const all = result.pages.map((_, i) => pageText(result, i)).join("");
    expect(all).toContain("first tall line");
    expect(all).toContain("second tall line");
    expect(all).toContain("after");
  });

  it("lays out the chronology-style header (exact small line + colored border)", () => {
    const { result } = layout({
      "word/document.xml": wrapDocument(
        `<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:color="7A6E73" w:sz="4" w:space="2"/></w:pBdr><w:spacing w:after="0" w:before="0" w:line="170" w:lineRule="exact"/></w:pPr><w:r><w:rPr><w:color w:val="7A6E73"/><w:sz w:val="17"/></w:rPr><w:t>Created by Cobbery</w:t></w:r></w:p>` +
          p("body"),
      ),
    });
    const txt = result.pages[0].items.find((i) => i.kind === "text" && i.text.includes("Created"));
    if (txt?.kind !== "text") throw new Error("missing header text");
    expect(txt.props.color).toBe("#7A6E73");
    const edge = result.pages[0].items.find((i) => i.kind === "edge");
    if (edge?.kind !== "edge") throw new Error("missing border edge");
    expect(edge.border.color).toBe("#7A6E73");
  });
});

describe("floating images", () => {
  it("wraps text beside a square-anchored image", () => {
    const anchor = `<w:p><w:r><w:drawing>
      <wp:anchor xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">
        <wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>
        <wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>
        <wp:extent cx="1905000" cy="1905000"/>
        <wp:wrapSquare wrapText="bothSides"/>
        <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
            <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
              <pic:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rIdImg"/></pic:blipFill>
              <pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1905000" cy="1905000"/></a:xfrm></pic:spPr>
            </pic:pic>
          </a:graphicData>
        </a:graphic>
      </wp:anchor>
    </w:drawing></w:r>
    <w:r><w:t>${"words flow beside the floating image ".repeat(40)}</w:t></w:r></w:p>`;
    const { result } = layout({
      "word/document.xml": wrapDocument(anchor),
      "word/_rels/document.xml.rels": `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdImg" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/x.png"/>
</Relationships>`,
      "word/media/x.png": "PNGDATA",
    });
    const items = result.pages[0].items;
    const img = items.find((i) => i.kind === "image");
    expect(img).toBeDefined();
    if (img?.kind !== "image") return;
    expect(Math.round(img.width)).toBe(200); // 1905000 EMU = 200px
    // Text beside the image starts to its right; text below spans full width.
    const beside = items.filter(
      (i) => i.kind === "text" && i.text.trim() && i.lineTop < img.y + img.height,
    );
    const below = items.filter(
      (i) => i.kind === "text" && i.text.trim() && i.lineTop > img.y + img.height + 2,
    );
    expect(beside.length).toBeGreaterThan(0);
    expect(below.length).toBeGreaterThan(0);
    for (const t of beside) {
      if (t.kind !== "text") continue;
      expect(t.x).toBeGreaterThanOrEqual(img.x + img.width);
    }
    expect(below.some((t) => t.kind === "text" && t.x < img.x + img.width)).toBe(true);
  });
});

describe("tab leaders", () => {
  it("fills dotted leaders up to the tab stop (TOC pattern)", () => {
    const { result } = layout({
      "word/document.xml": wrapDocument(
        `<w:p><w:pPr><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9000"/></w:tabs></w:pPr>
          <w:r><w:t>Chapter One</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>4</w:t></w:r>
        </w:p>`,
      ),
    });
    const dots = result.pages[0].items.find(
      (i) => i.kind === "text" && /^\.{10,}$/.test(i.text),
    );
    expect(dots).toBeDefined();
    if (dots?.kind !== "text") return;
    const num = result.pages[0].items.find((i) => i.kind === "text" && i.text === "4");
    if (num?.kind !== "text") throw new Error("page number missing");
    expect(dots.x + dots.width).toBeLessThanOrEqual(num.x + 4);
  });

  it("uses unstyled TOC leader and page-number fonts for line metrics", () => {
    const doc = DocxDocument.load(
      makeDocx({
        "word/document.xml": wrapDocument(
          `<w:p><w:pPr><w:pStyle w:val="TOC1"/><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="5000"/></w:tabs></w:pPr>
            <w:hyperlink w:anchor="_Toc1">
              <w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t>Entry</w:t></w:r>
              <w:r><w:tab/></w:r>
              <w:r><w:fldChar w:fldCharType="begin"/></w:r>
              <w:r><w:instrText xml:space="preserve"> PAGEREF _Toc1 \\h </w:instrText></w:r>
              <w:r><w:fldChar w:fldCharType="separate"/></w:r>
              <w:r><w:t>4</w:t></w:r>
              <w:r><w:fldChar w:fldCharType="end"/></w:r>
            </w:hyperlink>
          </w:p>`,
        ),
        "word/styles.xml": `<?xml version="1.0"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri"/><w:sz w:val="20"/></w:rPr></w:rPrDefault></w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"/>
  <w:style w:type="paragraph" w:styleId="TOC1"><w:basedOn w:val="Normal"/></w:style>
  <w:style w:type="character" w:styleId="Hyperlink"><w:rPr><w:rFonts w:ascii="Arial"/></w:rPr></w:style>
</w:styles>`,
      }),
    );
    const familyMeasurer: TextMeasurer = {
      width: (text, font, letterSpacing) => measurer.width(text, font, letterSpacing),
      metrics: (font) => {
        const lineHeight = font.size * (font.family === "Calibri" ? 1.5 : 0.75);
        return { ascent: lineHeight * 0.8, descent: lineHeight * 0.2, lineHeight };
      },
    };
    const result = layoutDocument(doc, { measurer: familyMeasurer });
    const textItem = (text: string) => {
      const item = result.pages[0].items.find(
        (candidate) => candidate.kind === "text" && candidate.text === text,
      );
      if (item?.kind !== "text") throw new Error(`missing ${text}`);
      return item;
    };
    const title = textItem("Entry");
    const pageNumber = textItem("4");
    const leader = result.pages[0].items.find(
      (candidate) => candidate.kind === "text" && /^\.{10,}$/.test(candidate.text),
    );
    if (leader?.kind !== "text") throw new Error("missing leader");

    expect(title.font.family).toBe("Arial");
    expect(leader.font.family).toBe("Calibri");
    expect(pageNumber.font.family).toBe("Calibri");
    expect(title.lineHeight).toBeCloseTo((10 * 4 / 3) * 1.5, 3);
  });

  it("ignores nonleader tabs for line metrics but retains leader and mark metrics", () => {
    const paragraph = (leader = "", includeTab = true) =>
      `<w:p><w:pPr><w:spacing w:line="276" w:lineRule="auto"/>` +
      `<w:tabs><w:tab w:val="left" ${leader ? `w:leader="${leader}" ` : ""}w:pos="720"/></w:tabs>` +
      `<w:rPr><w:sz w:val="22"/></w:rPr></w:pPr>` +
      `<w:r><w:rPr><w:sz w:val="20"/></w:rPr><w:t>A</w:t></w:r>` +
      (includeTab ? `<w:r><w:rPr><w:sz w:val="22"/></w:rPr><w:tab/></w:r>` : "") +
      `<w:r><w:rPr><w:sz w:val="20"/></w:rPr><w:t>B</w:t></w:r></w:p>`;
    const resultFor = (body: string) =>
      layout({ "word/document.xml": wrapDocument(body) }).result;
    const textItem = (result: ReturnType<typeof layoutDocument>, text: string) => {
      const item = result.pages[0].items.find(
        (candidate) => candidate.kind === "text" && candidate.text === text,
      );
      if (item?.kind !== "text") throw new Error(`missing ${text}`);
      return item;
    };

    const control = resultFor(paragraph("", false));
    const nonleader = resultFor(paragraph());
    const leader = resultFor(paragraph("dot"));
    const controlA = textItem(control, "A");
    const nonleaderA = textItem(nonleader, "A");
    const leaderA = textItem(leader, "A");

    expect(nonleaderA.lineHeight).toBeCloseTo(controlA.lineHeight, 3);
    expect(leaderA.lineHeight).toBeGreaterThan(nonleaderA.lineHeight);
    expect(textItem(nonleader, "B").x).toBeCloseTo(textItem(leader, "B").x, 3);
    expect(textItem(nonleader, "B").x).toBeGreaterThan(textItem(control, "B").x);

    const markProps = `<w:pPr><w:rPr><w:sz w:val="18"/></w:rPr></w:pPr>`;
    const tabOnly = resultFor(
      `<w:p>${markProps}<w:r><w:rPr><w:sz w:val="40"/></w:rPr><w:tab/></w:r></w:p>` +
      p("MARKER"),
    );
    const empty = resultFor(`<w:p>${markProps}</w:p>` + p("MARKER"));
    expect(textItem(tabOnly, "MARKER").lineTop).toBeCloseTo(
      textItem(empty, "MARKER").lineTop,
      3,
    );
  });
});

describe("footnotes and endnotes", () => {
  const FN_RELS = `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdFn" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>
  <Relationship Id="rIdEn" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" Target="endnotes.xml"/>
</Relationships>`;

  const footnotesXml = (notes: string) => `<?xml version="1.0"?>
<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>
  <w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>
  ${notes}
</w:footnotes>`;

  const endnotesXml = (notes: string) => `<?xml version="1.0"?>
<w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:endnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:endnote>
  ${notes}
</w:endnotes>`;

  const note = (tag: "footnote" | "endnote", id: number, text: string) =>
    `<w:${tag} w:id="${id}"><w:p><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:${tag}Ref/></w:r><w:r><w:t xml:space="preserve"> ${text}</w:t></w:r></w:p></w:${tag}>`;

  it("renders a footnote at the bottom of the referencing page with a separator", () => {
    const { result } = layout({
      "word/document.xml": wrapDocument(
        `<w:p><w:r><w:t>Body text</w:t></w:r><w:r><w:footnoteReference w:id="1"/></w:r></w:p>`,
      ),
      "word/_rels/document.xml.rels": FN_RELS,
      "word/footnotes.xml": footnotesXml(note("footnote", 1, "alpha note text")),
    });
    expect(result.totalPages).toBe(1);
    const page = result.pages[0];
    expect(pageText(result, 0)).toContain("alpha note text");

    // Reference mark and the note's own mark are both superscript "1"s.
    const marks = page.items.filter(
      (i) => i.kind === "text" && i.text === "1" && i.props.verticalAlign === "superscript",
    );
    expect(marks.length).toBe(2);

    // Separator rule: a short (2in) line above the note, below the body text.
    const body = page.items.find((i) => i.kind === "text" && i.text.includes("Body"));
    if (body?.kind !== "text") throw new Error("body text missing");
    const sep = page.items.find(
      (i) => i.kind === "edge" && Math.abs(i.x2 - i.x1 - 192) < 1 && i.y1 === i.y2,
    );
    expect(sep).toBeDefined();
    if (sep?.kind !== "edge") return;
    expect(sep.y1).toBeGreaterThan(body.lineTop);

    // Note text sits at the bottom of the body box (1in margins on US Letter).
    const noteText = page.items.find((i) => i.kind === "text" && i.text.includes("alpha"));
    if (noteText?.kind !== "text") throw new Error("note text missing");
    expect(noteText.lineTop).toBeGreaterThan(sep.y1);
    expect(noteText.lineTop).toBeGreaterThan(880);
    expect(noteText.lineTop + noteText.lineHeight).toBeLessThanOrEqual(page.bodyBottom + 0.5);
  });

  it("reserves footnote space only in the referencing column", () => {
    const line = (text: string, footnote = false) =>
      `<w:p><w:pPr><w:spacing w:line="240" w:lineRule="exact"/></w:pPr>` +
      `<w:r><w:t>${text}</w:t></w:r>` +
      (footnote ? `<w:r><w:footnoteReference w:id="1"/></w:r>` : "") +
      `</w:p>`;
    const body =
      line("FIRST", true) +
      Array.from({ length: 34 }, (_, i) => line(`BODY${i + 1}`)).join("") +
      `<w:sectPr><w:pgSz w:w="6000" w:h="6000"/>` +
      `<w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720"/>` +
      `<w:cols w:num="2" w:space="240"/></w:sectPr>`;
    const { result } = layout({
      "word/document.xml": wrapDocument(body),
      "word/_rels/document.xml.rels": FN_RELS,
      "word/footnotes.xml": footnotesXml(note("footnote", 1, "column note")),
    });

    expect(result.totalPages).toBe(1);
    const page = result.pages[0];
    const first = page.items.find((item) => item.kind === "text" && item.text === "FIRST");
    const firstColumnLast = page.items.find(
      (item) => item.kind === "text" && item.text === "BODY15",
    );
    const sentinel = page.items.find((item) => item.kind === "text" && item.text === "BODY34");
    const noteText = page.items.find(
      (item) => item.kind === "text" && item.text.includes("column"),
    );
    if (
      first?.kind !== "text" ||
      firstColumnLast?.kind !== "text" ||
      sentinel?.kind !== "text" ||
      noteText?.kind !== "text"
    ) {
      throw new Error("expected text is missing");
    }
    // Multi-column Word leaves a 26px body reserve above the note area; the
    // sixteenth fixed-height line therefore remains in the noted column.
    expect(firstColumnLast.x).toBe(first.x);
    expect(sentinel.x).toBeGreaterThan(first.x);
    expect(sentinel.lineTop).toBeGreaterThan(noteText.lineTop);
    expect(noteText.x).toBeLessThan(sentinel.x);
  });

  it("binds a footnote in a split table row to the page painting its partition", () => {
    // A single row taller than a page: the split's rest-partition carries a
    // footnote reference and must register the note on page 2.
    const cellParas =
      Array.from({ length: 70 }, (_, i) => `<w:p><w:r><w:t>Row line ${i}</w:t></w:r></w:p>`).join("") +
      `<w:p><w:r><w:t xml:space="preserve">noted line</w:t></w:r><w:r><w:footnoteReference w:id="1"/></w:r></w:p>`;
    const tbl =
      `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="9026"/></w:tblGrid>` +
      `<w:tr><w:tc><w:tcPr><w:tcW w:w="9026" w:type="dxa"/></w:tcPr>${cellParas}</w:tc></w:tr></w:tbl>`;
    const { result } = layout({
      "word/document.xml": wrapDocument(tbl + `<w:p><w:r><w:t>after table</w:t></w:r></w:p>`),
      "word/_rels/document.xml.rels": FN_RELS,
      "word/footnotes.xml": footnotesXml(note("footnote", 1, "split row note")),
    });
    expect(result.totalPages).toBeGreaterThan(1);
    // The reference line and the note text must share a page.
    const refPage = result.pages.findIndex((pg) => pg.items.some((i) => i.kind === "text" && i.text.includes("noted")));
    const notePage = result.pages.findIndex((pg) => pg.items.some((i) => i.kind === "text" && i.text.includes("split")));
    expect(refPage).toBeGreaterThan(0);
    expect(notePage).toBe(refPage);
  });

  it("keeps each footnote on the same page as its reference", () => {
    const filler = (n: number, from = 0) =>
      Array.from({ length: n }, (_, i) => p(`Filler paragraph ${from + i}`)).join("");
    const refPara = (word: string, id: number) =>
      `<w:p><w:r><w:t xml:space="preserve">${word}</w:t></w:r><w:r><w:footnoteReference w:id="${id}"/></w:r></w:p>`;
    const { result } = layout({
      "word/document.xml": wrapDocument(
        refPara("HEADREF", 1) + filler(70) + refPara("TAILREF", 2) + filler(5, 70),
      ),
      "word/_rels/document.xml.rels": FN_RELS,
      "word/footnotes.xml": footnotesXml(
        note("footnote", 1, "alphanote") + note("footnote", 2, "betanote"),
      ),
    });
    expect(result.totalPages).toBeGreaterThan(1);
    const pageOf = (needle: string) =>
      result.pages.findIndex((pg) =>
        pg.items.some((i) => i.kind === "text" && i.text.includes(needle)),
      );
    expect(pageOf("HEADREF")).toBe(0);
    expect(pageOf("alphanote")).toBe(0);
    const tailPage = pageOf("TAILREF");
    expect(tailPage).toBeGreaterThan(0);
    expect(pageOf("betanote")).toBe(tailPage);
    // Numbering follows document order.
    expect(pageText(result, 0)).toContain("alphanote");
    const tailMarks = result.pages[tailPage].items.filter(
      (i) => i.kind === "text" && i.text === "2" && i.props.verticalAlign === "superscript",
    );
    expect(tailMarks.length).toBe(2);
  });

  it("shrinks the body so footnote content never overlaps the footer margin", () => {
    const longNote = Array.from({ length: 4 }, (_, i) => p(`NOTETOKEN${i} with a good amount of text to wrap around`)).join("");
    const filler = Array.from({ length: 60 }, (_, i) => p(`Body ${i}`)).join("");
    const { result } = layout({
      "word/document.xml": wrapDocument(
        filler + `<w:p><w:r><w:t>REFHERE</w:t></w:r><w:r><w:footnoteReference w:id="1"/></w:r></w:p>`,
      ),
      "word/_rels/document.xml.rels": FN_RELS,
      "word/footnotes.xml": footnotesXml(
        `<w:footnote w:id="1">${longNote}</w:footnote>`,
      ),
    });
    for (const pg of result.pages) {
      for (const item of pg.items.slice(0, pg.hfStart)) {
        if (item.kind === "text") {
          expect(item.lineTop + item.lineHeight).toBeLessThanOrEqual(pg.bodyBottom + 0.5);
        }
      }
    }
    const refPage = result.pages.findIndex((pg) =>
      pg.items.some((i) => i.kind === "text" && i.text.includes("REFHERE")),
    );
    const notePage = result.pages.findIndex((pg) =>
      pg.items.some((i) => i.kind === "text" && i.text.includes("NOTETOKEN0")),
    );
    expect(notePage).toBe(refPage);
  });

  it("flows endnotes after the last body block with lowerRoman marks", () => {
    const { result } = layout({
      "word/document.xml": wrapDocument(
        `<w:p><w:r><w:t>Body start</w:t></w:r><w:r><w:endnoteReference w:id="1"/></w:r></w:p>` +
          p("Body end"),
      ),
      "word/_rels/document.xml.rels": FN_RELS,
      "word/endnotes.xml": endnotesXml(note("endnote", 1, "closing remark")),
    });
    expect(result.totalPages).toBe(1);
    const page = result.pages[0];
    expect(pageText(result, 0)).toContain("closing remark");
    const marks = page.items.filter(
      (i) => i.kind === "text" && i.text === "i" && i.props.verticalAlign === "superscript",
    );
    expect(marks.length).toBe(2);
    // Endnote content comes after the last body paragraph, not at page bottom.
    const bodyEnd = page.items.find((i) => i.kind === "text" && i.text === "end");
    const noteText = page.items.find((i) => i.kind === "text" && i.text.includes("closing"));
    if (bodyEnd?.kind !== "text" || noteText?.kind !== "text") throw new Error("items missing");
    expect(noteText.lineTop).toBeGreaterThan(bodyEnd.lineTop);
    expect(noteText.lineTop).toBeLessThan(400);
  });

  it("honors sectPr footnote number format", () => {
    const { result } = layout({
      "word/document.xml": wrapDocument(
        `<w:p><w:r><w:t>Text</w:t></w:r><w:r><w:footnoteReference w:id="1"/></w:r></w:p>` +
          `<w:sectPr><w:footnotePr><w:numFmt w:val="chicago"/></w:footnotePr>
            <w:pgSz w:w="12240" w:h="15840"/>
            <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/>
          </w:sectPr>`,
      ),
      "word/_rels/document.xml.rels": FN_RELS,
      "word/footnotes.xml": footnotesXml(note("footnote", 1, "starred note")),
    });
    const marks = result.pages[0].items.filter((i) => i.kind === "text" && i.text === "*");
    expect(marks.length).toBe(2);
  });
});

describe("justified line breaking (Word pack-vs-break rule)", () => {
  // ApproxMeasurer: n/o = 0.5em, i = 0.28em, m = 0.85em, space = 0.25em,
  // em = 14.666px. Ten "no" fillers + 9 spaces = 179.6585px of line.
  const fillers = Array(10).fill("no").join(" ");
  const jp = (text: string, bidi = false) =>
    `<w:p><w:pPr>${bidi ? "<w:bidi/>" : ""}<w:jc w:val="both"/></w:pPr><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
  const sect = (pgW: number) =>
    `<w:sectPr><w:pgSz w:w="${pgW}" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>`;
  const linesOf = (result: ReturnType<typeof layoutDocument>) => {
    const tops = new Map<number, string[]>();
    for (const i of result.pages[0].items) {
      if (i.kind !== "text") continue;
      const key = Math.round(i.lineTop);
      if (!tops.has(key)) tops.set(key, []);
      tops.get(key)!.push(i.text);
    }
    return [...tops.entries()].sort((a, b) => a[0] - b[0]).map(([, t]) => t.join("").replace(/\s+/g, " ").trim());
  };

  it("packs a wide final word at ~20% compression (break would leave a gaping line)", () => {
    // content 225.87px; "mmmm" needs 20% space compression, the break
    // alternative a 140% stretch: Word packs (compress <= stretch/2, <= 25%).
    const { result } = layout({
      "word/document.xml": wrapDocument(jp(`${fillers} mmmm`) + sect(6268)),
    });
    const lines = linesOf(result);
    expect(lines[0].endsWith("mmmm")).toBe(true);
  });

  it("breaks before a narrow final word at ~16% compression (stretch is cheap)", () => {
    // content 188.87px; "in" needs only 16% compression but breaking costs a
    // mere 28% stretch: Word breaks (16% > 28%/2).
    const { result } = layout({
      "word/document.xml": wrapDocument(jp(`${fillers} in`) + sect(5713)),
    });
    const lines = linesOf(result);
    expect(lines[0].endsWith("no")).toBe(true);
    expect(lines[1]).toBe("in");
  });

  it("does not overpack a bidi paragraph at the LTR compression boundary", () => {
    const text = `${fillers} mmmm`;
    const { result: ltr } = layout({
      "word/document.xml": wrapDocument(jp(text) + sect(6268)),
    });
    const { result: bidi } = layout({
      "word/document.xml": wrapDocument(jp(text, true) + sect(6268)),
    });

    expect(linesOf(ltr)[0].endsWith("mmmm")).toBe(true);
    expect(linesOf(bidi)[0].endsWith("no")).toBe(true);
    expect(linesOf(bidi)[1]).toBe("mmmm");
  });

  it("only overpacks justified lines in compatibility mode 15", () => {
    const text = `${fillers} mmmm`;
    const settings = (mode: number) =>
      `<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
      `<w:compat><w:compatSetting w:name="compatibilityMode" w:val="${mode}"/></w:compat>` +
      `</w:settings>`;
    const resultFor = (mode: number) =>
      layout({
        "word/document.xml": wrapDocument(jp(text) + sect(6268)),
        "word/settings.xml": settings(mode),
      }).result;

    expect(linesOf(resultFor(15))[0].endsWith("mmmm")).toBe(true);
    expect(linesOf(resultFor(14))[0].endsWith("no")).toBe(true);
    expect(linesOf(resultFor(14))[1]).toBe("mmmm");
  });

  it("never splits a word at a formatting-run boundary", () => {
    // "c" (run 1) + "cc" (bold run 2) form one word; "c" alone fits on line 1
    // but the word must move down as a unit.
    const para =
      '<w:p><w:r><w:t xml:space="preserve">aa bb c</w:t></w:r>' +
      '<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">cc</w:t></w:r></w:p>';
    const { result } = layout({ "word/document.xml": wrapDocument(para + sect(3650)) });
    const lines = linesOf(result);
    expect(lines[0]).toBe("aa bb");
    expect(lines[1]).toBe("ccc");
  });
});

describe("superscript / subscript", () => {
  it("raises superscript and drops subscript from the baseline like Word", () => {
    const para =
      '<w:p><w:r><w:t xml:space="preserve">base </w:t></w:r>' +
      '<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:t>sup</w:t></w:r>' +
      '<w:r><w:rPr><w:vertAlign w:val="subscript"/></w:rPr><w:t>sub</w:t></w:r></w:p>';
    const { result } = layout({ "word/document.xml": wrapDocument(para) });
    const items = result.pages[0].items.filter((i) => i.kind === "text");
    const base = items.find((i) => i.kind === "text" && i.text === "base")!;
    const sup = items.find((i) => i.kind === "text" && i.text === "sup")!;
    const sub = items.find((i) => i.kind === "text" && i.text === "sub")!;
    if (base.kind !== "text" || sup.kind !== "text" || sub.kind !== "text") throw new Error();
    // Word (probe-vertalign): raise 7/22 and drop 1/11 of the UNSCALED size;
    // scaled size is 65% rounded to half-points (14.666px -> 9.333px).
    expect(sup.font.size).toBeCloseTo(9.3333, 3);
    expect(base.baseline - sup.baseline).toBeCloseTo(14.666 * (7 / 22), 2);
    expect(sub.baseline - base.baseline).toBeCloseTo(14.666 / 11, 2);
    // Renderer anchor: explicit glyph box, baseline-aligned. Every span
    // anchors to the engine baseline (spaced-line leading hangs below).
    expect(sup.glyphTop).toBeCloseTo(sup.baseline - 0.9 * sup.font.size, 2);
    expect(sup.glyphBoxH).toBeCloseTo(1.15 * sup.font.size, 2);
    expect(base.glyphTop).toBeCloseTo(base.baseline - 0.9 * base.font.size, 2);
  });

  it("keys positioned-run line growth to that run's own size", () => {
    const body = (position = "") =>
      `<w:p>` +
      `<w:r><w:rPr><w:sz w:val="40"/></w:rPr><w:t>TALL</w:t></w:r>` +
      `<w:r><w:rPr><w:sz w:val="20"/><w:vertAlign w:val="superscript"/>${position}</w:rPr><w:t>sup</w:t></w:r>` +
      `</w:p>` +
      p("MARKER");
    const control = layout({ "word/document.xml": wrapDocument(body()) }).result;
    const positioned = layout({
      "word/document.xml": wrapDocument(body(`<w:position w:val="20"/>`)),
    }).result;
    const text = (result: ReturnType<typeof layoutDocument>, value: string) => {
      const item = result.pages[0].items.find((candidate) => candidate.kind === "text" && candidate.text === value);
      if (item?.kind !== "text") throw new Error(`missing ${value}`);
      return item;
    };
    const controlTall = text(control, "TALL");
    const controlSup = text(control, "sup");
    const positionedTall = text(positioned, "TALL");
    const positionedSup = text(positioned, "sup");
    const markerDelta = text(positioned, "MARKER").lineTop - text(control, "MARKER").lineTop;

    // The small raised run barely protrudes past the 20pt neighbor; its 10pt
    // position must not be charged against that unrelated tall run's ascent.
    expect(markerDelta).toBeLessThan(2);
    // Paint still applies the full 10pt baseline shift to the target run.
    expect(
      (positionedTall.baseline - positionedSup.baseline) -
        (controlTall.baseline - controlSup.baseline),
    ).toBeCloseTo(10 * (4 / 3), 3);
  });

  it("reuses the descent when every text run on a line is raised", () => {
    const paragraph = (mixed = false, exact = false) =>
      `<w:p><w:pPr>${exact ? '<w:spacing w:line="600" w:lineRule="exact"/>' : ""}</w:pPr>` +
      `<w:r><w:rPr><w:sz w:val="40"/><w:position w:val="32"/></w:rPr><w:t>raised</w:t></w:r>` +
      `<w:r><w:rPr><w:sz w:val="40"/>${mixed ? "" : '<w:position w:val="32"/>'}</w:rPr><w:t> peer</w:t></w:r>` +
      `</w:p>` +
      p("MARKER");
    const line = (mixed = false, exact = false) => {
      const result = layout({
        "word/document.xml": wrapDocument(paragraph(mixed, exact)),
      }).result;
      const raised = result.pages[0].items.find(
        (item) => item.kind === "text" && item.text === "raised",
      );
      const marker = result.pages[0].items.find(
        (item) => item.kind === "text" && item.text === "MARKER",
      );
      if (raised?.kind !== "text" || marker?.kind !== "text") throw new Error();
      return { raised, marker };
    };

    const uniform = line();
    const mixed = line(true);
    const exact = line(false, true);
    const size = 20 * (4 / 3);
    const natural = size * 1.15;
    const raise = 16 * (4 / 3);
    const descent = size * 0.25;

    // The dense legacy-equation control has consecutive 20pt Times lines,
    // all at +16pt. Word advances them by 34.78pt: natural height + raise,
    // less the font descent. ApproxMeasurer uses a 0.25em descent. A baseline
    // run prevents that descent reuse.
    expect(uniform.marker.lineTop - uniform.raised.lineTop).toBeCloseTo(
      natural + raise - descent,
      3,
    );
    expect(mixed.marker.lineTop - mixed.raised.lineTop).toBeCloseTo(natural + raise, 3);
    expect(uniform.raised.lineHeight).toBeCloseTo(natural + raise - descent, 3);
    expect(mixed.raised.lineHeight).toBeCloseTo(natural + raise, 3);
    // Exact line spacing already suppresses positioned-run growth entirely.
    expect(exact.raised.lineHeight).toBeCloseTo(40, 3);
  });

  it("a raised label beside a tall inline image does not inflate the line", () => {
    // dense figure: a small "V1" label raised +160pt (w:position 320) shares
    // the line with a ~248px inline picture. The raised label stays inside the
    // image extent, so the figure line must keep the image's own height, not
    // image + raise (which doubled the block and desynced pagination).
    const rels = `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdImg" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/x.png"/>
</Relationships>`;
    const inlineImg =
      `<w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">` +
      `<wp:extent cx="2700000" cy="2361681"/>` +
      `<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
      `<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
      `<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
      `<pic:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rIdImg"/></pic:blipFill>` +
      `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2700000" cy="2361681"/></a:xfrm></pic:spPr>` +
      `</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>`;
    const figPara = (label: string) =>
      `<w:p><w:r>${inlineImg}</w:r>${label}</w:p>`;
    const raisedLabel = `<w:r><w:rPr><w:position w:val="320"/><w:sz w:val="32"/></w:rPr><w:t>V1</w:t></w:r>`;
    const parts = (label: string) => ({
      "word/document.xml": wrapDocument(figPara(label) + p("MARKER")),
      "word/_rels/document.xml.rels": rels,
      "word/media/x.png": "PNGDATA",
    });
    const withRaise = layout(parts(raisedLabel)).result;
    const noRaise = layout(parts(`<w:r><w:t>V1</w:t></w:r>`)).result;
    const markerTop = (r: ReturnType<typeof layoutDocument>) => {
      const it = r.pages[0].items.find((i) => i.kind === "text" && i.text === "MARKER");
      if (it?.kind !== "text") throw new Error("missing MARKER");
      return it.lineTop;
    };
    // The +160pt (213px) raise must not push the MARKER paragraph down by
    // anything like the full shift: the figure line keeps the ~248px image
    // height (the label sits inside it), so the delta is at most a few px of
    // genuine protrusion, never the ~213px the old additive rule added.
    expect(markerTop(withRaise) - markerTop(noRaise)).toBeLessThan(5);
    // And it fits on one page (image ~248px << body), not two.
    expect(withRaise.totalPages).toBe(1);
  });

  it("an image-only line under a multiple lays only the spacing leading below it, no glyph descent", () => {
    // msa's signature rows: a lone inline group (no text run) in a paragraph
    // with a line multiple. Word clears such a text-less image line with only
    // the (k-1)x line-spacing leading below it - there is no glyph descent to
    // reserve. A trailing text run DOES add its below-share, so the same line
    // with a "." at the end must sit taller (pushing MARKER lower). Pinning
    // the two apart guards the image-only descent rule.
    const rels = `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdImg" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/x.png"/>
</Relationships>`;
    const inlineImg =
      `<w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">` +
      `<wp:extent cx="2095500" cy="400050"/>` +
      `<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
      `<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
      `<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
      `<pic:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rIdImg"/></pic:blipFill>` +
      `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2095500" cy="400050"/></a:xfrm></pic:spPr>` +
      `</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>`;
    const imgPara = (tail: string) =>
      `<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="276" w:lineRule="auto"/></w:pPr>` +
      `<w:r>${inlineImg}</w:r>${tail}</w:p>`;
    const parts = (tail: string) => ({
      "word/document.xml": wrapDocument(imgPara(tail) + p("MARKER")),
      "word/_rels/document.xml.rels": rels,
      "word/media/x.png": "PNGDATA",
    });
    const imageOnly = layout(parts("")).result;
    const withText = layout(parts(`<w:r><w:t>.</w:t></w:r>`)).result;
    const markerTop = (r: ReturnType<typeof layoutDocument>) => {
      const it = r.pages[0].items.find((i) => i.kind === "text" && i.text === "MARKER");
      if (it?.kind !== "text") throw new Error("missing MARKER");
      return it.lineTop;
    };
    // The text glyph's below-share (its quantized descent, ~2-3px) lifts MARKER
    // strictly lower than the image-only line does.
    expect(markerTop(withText)).toBeGreaterThan(markerTop(imageOnly) + 0.5);
  });
});

describe("sections & page borders", () => {
  it("continuous section shares the page; new-page section does not", () => {
    const cont = wrapDocument(
      p("first section text") +
        `<w:p><w:pPr><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:pPr></w:p>` +
        p("second section text") +
        `<w:sectPr><w:type w:val="continuous"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>`,
    );
    const { result } = layout({ "word/document.xml": cont });
    expect(result.totalPages).toBe(1);
    // and without continuous: two pages
    const hard = cont.replace('<w:type w:val="continuous"/>', "");
    const { result: r2 } = layout({ "word/document.xml": hard });
    expect(r2.totalPages).toBe(2);
  });

  it("renders page borders inset from the text margins", () => {
    const docXml = wrapDocument(
      p("bordered page") +
        `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>` +
        `<w:pgBorders w:offsetFrom="text"><w:top w:val="single" w:sz="8" w:space="24" w:color="FF0000"/>` +
        `<w:left w:val="single" w:sz="8" w:space="24"/><w:bottom w:val="single" w:sz="8" w:space="24"/>` +
        `<w:right w:val="single" w:sz="8" w:space="24"/></w:pgBorders></w:sectPr>`,
    );
    const { result } = layout({ "word/document.xml": docXml });
    const edges = result.pages[0].items.filter((i) => i.kind === "edge");
    expect(edges.length).toBe(4);
    const top = edges.find((e) => e.kind === "edge" && e.y1 === e.y2 && e.y1 < 100);
    if (!top || top.kind !== "edge") throw new Error();
    expect(top.border.color).toBe("#FF0000");
    // 1in margin (96px) minus 24pt (32px) offset, then half of a 1pt rule
    // because w:space measures to the border edge.
    expect(top.y1).toBeCloseTo(63.33, 1);
  });
});

describe("table autofit + tblInd (wild2-sci-chem-omml p9 Word PDF)", () => {
  // Word's rendered autofit columns for a table that paints NO vertical
  // rules are content + cell margins EXACTLY (chem p9: 31.8pt = "3.81" at
  // 21pt + 10.8pt margins), and in compatibilityMode <= 14 w:tblInd measures
  // to the first cell's TEXT edge (the grid begins a cell left-margin
  // further left; mode 15 measures to the border).
  const SECT =
    `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>` +
    `<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>`;
  const tblXml =
    `<w:tbl><w:tblPr>` +
    `<w:tblW w:w="0" w:type="auto"/><w:tblInd w:w="240" w:type="dxa"/>` +
    `<w:tblBorders><w:top w:val="single" w:sz="4"/><w:bottom w:val="single" w:sz="4"/></w:tblBorders>` +
    `<w:tblCellMar><w:left w:w="108" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar>` +
    `</w:tblPr><w:tblGrid><w:gridCol w:w="600"/><w:gridCol w:w="700"/></w:tblGrid>` +
    `<w:tr>` +
    `<w:tc><w:tcPr><w:tcW w:w="0" w:type="auto"/></w:tcPr><w:p><w:r><w:t>3.81</w:t></w:r></w:p></w:tc>` +
    `<w:tc><w:tcPr><w:tcW w:w="0" w:type="auto"/></w:tcPr><w:p><w:r><w:t>3.529</w:t></w:r></w:p></w:tc>` +
    `</w:tr></w:tbl>`;
  const settingsXml = (mode: number) =>
    `<?xml version="1.0"?><w:settings ${W_NS}><w:compat>` +
    `<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="${mode}"/>` +
    `</w:compat></w:settings>`;
  const textX = (result: ReturnType<typeof layoutDocument>, text: string): number => {
    const it = result.pages[0].items.find((i) => i.kind === "text" && i.text === text);
    if (it?.kind !== "text") throw new Error(`missing ${text}`);
    return it.x;
  };

  it("no-vertical-rules autofit column = content + margins, and the sizing token never char-wraps", () => {
    const { result } = layout({ "word/document.xml": wrapDocument(tblXml + SECT) });
    // The token stays whole (an exact-fit column must not hard-wrap "3.81").
    const x1 = textX(result, "3.81");
    const x2 = textX(result, "3.529");
    const col1 = measurer.width("3.81", { family: "Calibri", size: 44 / 3, bold: false, italic: false }) + 14.4;
    // cell2 text = cell1 text + col1 width (content + 7.2 + 7.2, no +2 rule fudge)
    expect(x2 - x1).toBeCloseTo(col1, 1);
  });

  it("bordered zero-margin autofit column = content + 2×declared rule width (parity-tables Word PDF)", () => {
    // Word sizes a content-fit column of a BORDERED table at text + the
    // declared vertical-rule width on each side — 1.33px for the sz-4 grid —
    // not a flat 2px: parity-tables' "Left 2in" column renders 46.04px for
    // 44.71px of text, and the text ends 0.33px before the next rule.
    const tbl =
      `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>` +
      `<w:tblBorders>` +
      `<w:top w:val="single" w:sz="4"/><w:left w:val="single" w:sz="4"/>` +
      `<w:bottom w:val="single" w:sz="4"/><w:right w:val="single" w:sz="4"/>` +
      `<w:insideH w:val="single" w:sz="4"/><w:insideV w:val="single" w:sz="4"/>` +
      `</w:tblBorders></w:tblPr>` +
      `<w:tblGrid><w:gridCol w:w="2880"/><w:gridCol w:w="5760"/></w:tblGrid>` +
      `<w:tr>` +
      `<w:tc><w:p><w:r><w:t>Left 2in</w:t></w:r></w:p></w:tc>` +
      `<w:tc><w:p><w:r><w:t>Right side content</w:t></w:r></w:p></w:tc>` +
      `</w:tr></w:tbl>`;
    const { result } = layout({ "word/document.xml": wrapDocument(tbl + SECT) });
    const x1 = textX(result, "Left");
    const x2 = textX(result, "Right");
    const text1 = measurer.width("Left 2in", { family: "Calibri", size: 44 / 3, bold: false, italic: false });
    // col1 = text + 2×(sz-4 rule = 0.5pt); both cells inset 1px from their
    // left grid edge, so the text gap equals the column width.
    expect(x2 - x1).toBeCloseTo(text1 + 2 * ((0.5 * 96) / 72), 1);
    // And the exact-fit line must not wrap: "2in" stays on the first line.
    const left = result.pages[0].items.find((i) => i.kind === "text" && i.text === "Left");
    const tail = result.pages[0].items.find((i) => i.kind === "text" && i.text === "2in");
    if (left?.kind !== "text" || tail?.kind !== "text") throw new Error("cells not found");
    expect(tail.baseline).toBeCloseTo(left.baseline, 3);
  });

  it("BORDERLESS autofit column = content + margins exactly, no phantom rule allowance", () => {
    // A table with no tblBorders anywhere in its style chain paints no
    // vertical rules, so its measured column widths are content + margins
    // EXACTLY — not content + margins + an assumed sz-4 rule. This is the
    // NIH clause-matrix constraint (wild2-legal-nih-contract, tblW 4800 pct,
    // TableNormal margins 108tw): Word's p228 columns are [76.02, 59.28,
    // 365.82]pt and its one-line col3 title measures 357.61pt, so a 0.5pt
    // phantom rule in the Word-exact mins re-runs the pct-raise
    // redistribution, takes ~2px from col3 and wraps 195 pages (+3 total).
    // Reconciled with parity-tables p1 (bordered, rule-aware pad, above):
    // the allowance is the DECLARED rule width, and borderless means 0.
    const tbl =
      `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>` +
      `<w:tblCellMar><w:left w:w="108" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar>` +
      `</w:tblPr>` +
      `<w:tblGrid><w:gridCol w:w="1394"/><w:gridCol w:w="7435"/></w:tblGrid>` +
      `<w:tr>` +
      `<w:tc><w:p><w:r><w:t>Vamom</w:t></w:r></w:p></w:tc>` +
      `<w:tc><w:p><w:r><w:t>Figican by Pikuhuzoke</w:t></w:r></w:p></w:tc>` +
      `</w:tr></w:tbl>`;
    const { result } = layout({ "word/document.xml": wrapDocument(tbl + SECT) });
    const x1 = textX(result, "Vamom");
    const x2 = textX(result, "Figican");
    const text1 = measurer.width("Vamom", { family: "Calibri", size: 44 / 3, bold: false, italic: false });
    // col1 = text + 7.2px + 7.2px (108tw sides), no rule term: both cells
    // inset their text by the same 7.2px left margin, so the x gap equals
    // the column width.
    expect(x2 - x1).toBeCloseTo(text1 + 14.4, 1);
  });

  it("BORDERED margin-ful autofit column = content + margins + capped 2px allowance", () => {
    // Same table WITH an sz-4 grid and 108tw margins: the rule allowance
    // bridges from the declared rule width (zero-margin, parity-tables) up
    // to the legacy 2px cap once the margins exceed it — min(2, rule +
    // margin). NIH p358-360's status tables (tblW 4000/4200 pct, sz-4,
    // 108tw) reproduce Word's scale-down columns ([187.83, 90.80, 57.27,
    // 102.28]pt on p359) only at the 2px cap: their two-line " Rugehini
    // doluguseqesu qapabipe" wrap sits 0.4px from col1's edge, and a
    // rule-only allowance (0.667px) tips it the wrong way.
    const tbl =
      `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>` +
      `<w:tblBorders>` +
      `<w:top w:val="single" w:sz="4"/><w:left w:val="single" w:sz="4"/>` +
      `<w:bottom w:val="single" w:sz="4"/><w:right w:val="single" w:sz="4"/>` +
      `<w:insideH w:val="single" w:sz="4"/><w:insideV w:val="single" w:sz="4"/>` +
      `</w:tblBorders>` +
      `<w:tblCellMar><w:left w:w="108" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar>` +
      `</w:tblPr>` +
      `<w:tblGrid><w:gridCol w:w="1394"/><w:gridCol w:w="7435"/></w:tblGrid>` +
      `<w:tr>` +
      `<w:tc><w:p><w:r><w:t>Vamom</w:t></w:r></w:p></w:tc>` +
      `<w:tc><w:p><w:r><w:t>Figican by Pikuhuzoke</w:t></w:r></w:p></w:tc>` +
      `</w:tr></w:tbl>`;
    const { result } = layout({ "word/document.xml": wrapDocument(tbl + SECT) });
    const x1 = textX(result, "Vamom");
    const x2 = textX(result, "Figican");
    const text1 = measurer.width("Vamom", { family: "Calibri", size: 44 / 3, bold: false, italic: false });
    // col1 = text + 7.2 + 7.2 margins + min(2, 0.667 + 7.2) = +2px.
    expect(x2 - x1).toBeCloseTo(text1 + 14.4 + 2, 1);
  });

  it("compatibilityMode 14 shifts a tblInd table left by the first cell margin", () => {
    const parts15 = { "word/document.xml": wrapDocument(tblXml + SECT), "word/settings.xml": settingsXml(15) };
    const parts14 = { "word/document.xml": wrapDocument(tblXml + SECT), "word/settings.xml": settingsXml(14) };
    const x15 = textX(layout(parts15).result, "3.81");
    const x14 = textX(layout(parts14).result, "3.81");
    expect(x15).toBeCloseTo(96 + 16 + 7.2, 1); // margin + tblInd + cellMarLeft
    expect(x14).toBeCloseTo(x15 - 7.2, 1);
  });
});

describe("table row splitting", () => {
  const bigRow = (n: number, extra = "") => {
    const paras = Array.from({ length: n }, (_, i) => `<w:p><w:r><w:t>cell line ${i}</w:t></w:r></w:p>`).join("");
    return `<w:tbl><w:tblGrid><w:gridCol w:w="8000"/></w:tblGrid>
      <w:tr>${extra}<w:tc>${paras}</w:tc></w:tr></w:tbl>`;
  };

  it("splits a too-tall row across pages by default", () => {
    // ~90 lines at ~17px in a ~700px body: must span pages, content intact.
    const { result } = layout({ "word/document.xml": wrapDocument(bigRow(90)) });
    expect(result.totalPages).toBeGreaterThan(1);
    const all = result.pages.flatMap((pg) => pg.items.filter((i) => i.kind === "text"));
    expect(all.some((i) => i.kind === "text" && i.text.includes("0"))).toBe(true);
    const last = result.pages[result.pages.length - 1].items.filter((i) => i.kind === "text");
    expect(last.length).toBeGreaterThan(0);
    // every page's text stays inside the body box
    for (const pg of result.pages) {
      for (const it of pg.items) {
        if (it.kind === "text") expect(it.lineTop + it.lineHeight).toBeLessThanOrEqual(pg.height - 90);
      }
    }
  });

  it("splits a row that fits a fresh page when useful content fits before the break", () => {
    const filler = Array.from({ length: 7 }, (_, i) => p(`filler ${i}`)).join("");
    const section =
      `<w:sectPr><w:pgSz w:w="12240" w:h="3600"/>` +
      `<w:pgMar w:top="360" w:right="720" w:bottom="360" w:left="720"/></w:sectPr>`;
    const { result } = layout({
      "word/document.xml": wrapDocument(filler + bigRow(8) + section),
    });

    const firstPage = pageText(result, 0);
    const secondPage = pageText(result, 1);
    expect(result.totalPages).toBe(2);
    expect(firstPage).toContain("cell line 0");
    expect(secondPage).toContain("cell line 7");
  });

  it("hugs a nested-table split cut to the last whole nested row, not the body bottom", () => {
    // staging-grid4 p2: when a wrapper cell's kept content ends at a nested-row
    // rule, Word draws the page-slice bottom border right below that rule (last
    // dotted rule + cell trailing inset), NOT at the page body bottom — the
    // leftover band that cannot hold another whole nested row stays outside the
    // box (Word p2: outer border 19px above the body bottom).
    const innerRows = Array.from({ length: 30 }, (_, i) =>
      `<w:tr><w:tc><w:tcPr><w:tcW w:w="6800" w:type="dxa"/></w:tcPr>` +
      `<w:p><w:r><w:t>deep row ${i}</w:t></w:r></w:p></w:tc></w:tr>`,
    ).join("");
    const inner =
      `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>` +
      `<w:tblBorders><w:top w:val="dotted" w:sz="4" w:color="7030A0"/><w:bottom w:val="dotted" w:sz="4" w:color="7030A0"/><w:insideH w:val="dotted" w:sz="4" w:color="7030A0"/></w:tblBorders>` +
      `</w:tblPr><w:tblGrid><w:gridCol w:w="6800"/></w:tblGrid>${innerRows}</w:tbl>`;
    const outer =
      `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>` +
      `<w:tblBorders><w:top w:val="single" w:sz="12" w:color="1F3864"/><w:bottom w:val="single" w:sz="12" w:color="1F3864"/><w:left w:val="single" w:sz="12" w:color="1F3864"/><w:right w:val="single" w:sz="12" w:color="1F3864"/></w:tblBorders>` +
      `</w:tblPr><w:tblGrid><w:gridCol w:w="7000"/></w:tblGrid>` +
      `<w:tr><w:tc><w:tcPr><w:tcW w:w="7000" w:type="dxa"/></w:tcPr>` +
      `<w:p><w:r><w:t>Wrapper cell:</w:t></w:r></w:p>${inner}<w:p></w:p></w:tc></w:tr></w:tbl>`;
    const section =
      `<w:sectPr><w:pgSz w:w="12240" w:h="7200"/>` +
      `<w:pgMar w:top="720" w:right="1440" w:bottom="720" w:left="1440"/></w:sectPr>`;
    const { result } = layout({
      "word/document.xml": wrapDocument(outer + `<w:p></w:p>` + section),
    });
    expect(result.totalPages).toBeGreaterThan(1);
    const p1 = result.pages[0];
    const bodyBottom = p1.height - 720 / 15; // 480 - 48
    const hEdges = p1.items.filter(
      (i) => i.kind === "edge" && Math.abs(i.y1 - i.y2) < 0.01 && Math.abs(i.x2 - i.x1) > 4,
    );
    const solidBottom = Math.max(
      ...hEdges.filter((e) => e.kind === "edge" && e.border.color?.includes("1F3864")).map((e) => e.y1),
    );
    const lastDotted = Math.max(
      ...hEdges.filter((e) => e.kind === "edge" && e.border.color?.includes("7030A0")).map((e) => e.y1),
    );
    // The slice's outer bottom border hugs the last nested rule (within the
    // cell trailing inset), instead of sitting at the page body bottom.
    expect(solidBottom - lastDotted).toBeLessThan(10);
    expect(solidBottom).toBeLessThan(bodyBottom - 2);
  });

  it("reserves trailing paragraph space when choosing the row split line", () => {
    const filler = Array.from({ length: 7 }, (_, i) => p(`filler ${i}`)).join("");
    const lines = Array.from({ length: 5 }, (_, i) =>
      `<w:r><w:t>row line ${i}</w:t>${i < 4 ? "<w:br/>" : ""}</w:r>`,
    ).join("");
    const row = `<w:tbl><w:tblGrid><w:gridCol w:w="8000"/></w:tblGrid><w:tr><w:tc>
      <w:p><w:pPr><w:spacing w:after="400"/></w:pPr>${lines}</w:p>
    </w:tc></w:tr></w:tbl>`;
    const section =
      `<w:sectPr><w:pgSz w:w="12240" w:h="3600"/>` +
      `<w:pgMar w:top="360" w:right="720" w:bottom="360" w:left="720"/></w:sectPr>`;
    const { result } = layout({
      "word/document.xml": wrapDocument(filler + row + section),
    });

    expect(pageText(result, 0)).toContain("row line 0");
    expect(pageText(result, 1)).toContain("row line 3");
    expect(pageText(result, 1)).toContain("row line 4");
  });

  it("moves a three-line row whole instead of leaving a one-line continuation", () => {
    const filler = Array.from({ length: 10 }, (_, i) => p(`filler ${i}`)).join("");
    const section =
      `<w:sectPr><w:pgSz w:w="12240" w:h="3600"/>` +
      `<w:pgMar w:top="360" w:right="720" w:bottom="360" w:left="720"/></w:sectPr>`;
    const { result } = layout({
      "word/document.xml": wrapDocument(filler + bigRow(3) + section),
    });

    expect(pageText(result, 0)).not.toContain("cell line");
    expect(pageText(result, 1)).toContain("cell line 0");
    const rowLineTops = result.pages[1].items.flatMap((item) =>
      item.kind === "text" && item.text.startsWith("cell") ? [item.lineTop] : [],
    );
    expect(new Set(rowLineTops).size).toBe(3);
  });

  it("honors cantSplit by moving the row whole", () => {
    const xml = wrapDocument(p("lead") + bigRow(20, '<w:trPr><w:cantSplit/></w:trPr>') );
    // Fill most of page 1 first so the row doesn't fit
    const filler = Array.from({ length: 42 }, (_, i) => p(`filler ${i}`)).join("");
    const { result } = layout({ "word/document.xml": wrapDocument(filler + bigRow(25, "<w:trPr><w:cantSplit/></w:trPr>")) });
    void xml;
    // The row starts on page 2 (its first cell line is not on page 1).
    const p1 = result.pages[0].items.filter((i) => i.kind === "text").map((i) => (i.kind === "text" ? i.text : ""));
    expect(p1.some((t) => t.includes("cell"))).toBe(false);
  });
});

describe("inline drawing groups", () => {
  const GROUP =
    '<w:p><w:r><w:drawing xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"' +
    ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"' +
    ' xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"' +
    ' xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
    '<wp:inline><wp:extent cx="329184" cy="329184"/>' +
    '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup">' +
    "<wpg:wgp><wpg:grpSpPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"329184\" cy=\"329184\"/>" +
    '<a:chOff x="0" y="0"/><a:chExt cx="208" cy="208"/></a:xfrm></wpg:grpSpPr>' +
    "<wps:wsp><wps:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"208\" cy=\"208\"/></a:xfrm>" +
    '<a:solidFill><a:srgbClr val="37B6AE"/></a:solidFill>' +
    '<a:custGeom><a:pathLst><a:path w="208" h="208">' +
    '<a:moveTo><a:pt x="0" y="0"/></a:moveTo><a:lnTo><a:pt x="208" y="208"/></a:lnTo>' +
    "</a:path></a:pathLst></a:custGeom></wps:spPr></wps:wsp>" +
    "</wpg:wgp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>";

  it("keeps the drawingHit target on the group when it sits inside a table cell", () => {
    const tbl =
      "<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w=\"4000\"/></w:tblGrid>" +
      `<w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>${GROUP}</w:tc></w:tr></w:tbl>`;
    const { result } = layout({ "word/document.xml": wrapDocument(tbl + p("after")) });
    const items = result.pages[0].items;
    const hit = items.find((i) => i.kind === "drawingHit");
    const path = items.find((i) => i.kind === "path");
    expect(hit).toBeDefined();
    expect(path).toBeDefined();
    if (hit?.kind !== "drawingHit" || path?.kind !== "path") throw new Error();
    // The transparent hit target must sit exactly over the painted group
    // even after the cell-frame offset (offsetItem must shift drawingHit).
    expect(hit.x).toBeCloseTo(path.x, 3);
    expect(hit.y).toBeCloseTo(path.y, 3);
    expect(hit.y).toBeGreaterThan(50); // moved with the cell, not stuck at origin
  });

});

describe("East Asian (CJK) layout", () => {
  // A run whose eastAsia font is set; ideographs are laid one em wide and
  // every inter-character boundary is a break opportunity.
  const cjk = (text: string) =>
    `<w:p><w:pPr><w:jc w:val="both"/></w:pPr>` +
    `<w:r><w:rPr><w:rFonts w:eastAsia="MS Mincho" w:ascii="Calibri"/></w:rPr>` +
    `<w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;

  it("emits one text item per CJK character (inter-character breaking)", () => {
    const { result } = layout({ "word/document.xml": wrapDocument(cjk("水は方円")) });
    const texts = result.pages[0].items.filter((i) => i.kind === "text");
    // Every ideograph/kana is its own breakable item.
    expect(texts.length).toBe(4);
    for (const t of texts) if (t.kind === "text") expect(t.text.length).toBe(1);
  });

  it("lays CJK glyphs one em (font size) wide and wraps between characters", () => {
    // 40 CJK chars at 11pt = 40em; the default text column is far narrower, so
    // the line must wrap between characters into multiple rows.
    const { result } = layout({ "word/document.xml": wrapDocument(cjk("水".repeat(90))) });
    const texts = result.pages[0].items.filter((i) => i.kind === "text");
    if (texts[0].kind !== "text") throw new Error();
    // 11pt -> 14.667px; one em wide.
    expect(texts[0].width).toBeCloseTo(texts[0].font.size, 1);
    const rows = new Set(texts.map((t) => (t.kind === "text" ? Math.round(t.baseline) : 0)));
    expect(rows.size).toBeGreaterThan(1);
  });

  it("Chinese fallback is by MS Mincho cmap COVERAGE, not kana presence", () => {
    // staging-eastasian 年号 run: a kana-less segment whose every code point
    // MS Mincho covers KEEPS the Japanese face's line profile (Word lays that
    // line at the 26px Mincho pitch); only a segment containing a
    // simplified-only form (时) drops to the Chinese fallback profile.
    const famsOf = (text: string) => {
      const { result } = layout({ "word/document.xml": wrapDocument(cjk(text)) });
      return new Set(
        result.pages[0].items.filter((i) => i.kind === "text").map((i) => (i.kind === "text" ? i.font.family : "")),
      );
    };
    const covered = famsOf("年号");
    expect(covered.has("Hiragino Mincho ProN")).toBe(true);
    expect(covered.has("PingFang TC")).toBe(false);
    const fallback = famsOf("学时习");
    expect(fallback.has("PingFang TC")).toBe(true);
  });
});

describe("RTL / bidi paragraphs", () => {
  const bidiP = (inner: string, jc = "right") =>
    `<w:p><w:pPr><w:bidi/><w:jc w:val="${jc}"/></w:pPr>${inner}</w:p>`;
  const rtlRun = (t: string) =>
    `<w:r><w:rPr><w:rFonts w:cs="Arial"/><w:rtl/></w:rPr><w:t xml:space="preserve">${t}</w:t></w:r>`;
  const ltrRun = (t: string) => `<w:r><w:t xml:space="preserve">${t}</w:t></w:r>`;
  const narrowSection =
    `<w:sectPr><w:pgSz w:w="5880" w:h="15840"/>` +
    `<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>`;

  it("marks RTL runs so the renderer sets direction:rtl", () => {
    const { result } = layout({ "word/document.xml": wrapDocument(bidiP(rtlRun("שלום"))) });
    const t = result.pages[0].items.find((i) => i.kind === "text");
    if (t?.kind !== "text") throw new Error();
    expect(t.rtl).toBe(true);
  });

  it("aligns a bidi jc=right paragraph to the physical left (Word swaps end->left)", () => {
    const { result } = layout({ "word/document.xml": wrapDocument(bidiP(rtlRun("שלום"))) });
    const t = result.pages[0].items.find((i) => i.kind === "text");
    if (t?.kind !== "text") throw new Error();
    // Flush left: the single word sits at the left margin, not the right edge.
    expect(t.x).toBeLessThan(200); // near the left margin, not the right edge
  });

  it("reorders a mixed run visually: the LTR run sits left of the RTL run", () => {
    const { result } = layout({
      "word/document.xml": wrapDocument(bidiP(rtlRun("אב") + ltrRun("12"))),
    });
    const texts = result.pages[0].items.filter((i) => i.kind === "text");
    const rtl = texts.find((i) => i.kind === "text" && i.rtl);
    const ltr = texts.find((i) => i.kind === "text" && !i.rtl && i.text.trim());
    if (rtl?.kind !== "text" || ltr?.kind !== "text") throw new Error();
    // Logical order is RTL-run then LTR-run; visually the LTR run moves left of
    // the RTL run (base RTL: first logical run is rightmost).
    expect(ltr.x).toBeLessThan(rtl.x);
  });

  it("applies first-line indentation from the paragraph's logical start", () => {
    const indentedP = (bidi: boolean, firstLine: number, inner: string) =>
      `<w:p><w:pPr>${bidi ? "<w:bidi/>" : ""}` +
      `<w:ind w:left="240" w:right="360" w:firstLine="${firstLine}"/></w:pPr>${inner}</w:p>`;
    const itemOf = (bidi: boolean, firstLine: number) => {
      const text = bidi ? "שלום" : "text";
      const inner = bidi ? rtlRun(text) : ltrRun(text);
      const { result } = layout({
        "word/document.xml": wrapDocument(indentedP(bidi, firstLine, inner)),
      });
      const item = result.pages[0].items.find((i) => i.kind === "text" && i.text === text);
      if (item?.kind !== "text") throw new Error();
      return item;
    };

    const ltrBase = itemOf(false, 0);
    const ltrIndented = itemOf(false, 720);
    const bidiBase = itemOf(true, 0);
    const bidiIndented = itemOf(true, 720);

    expect(ltrBase.x).toBeCloseTo(112, 3); // 96px margin + 16px physical left indent
    // w:ind left/right are LOGICAL start/end: in a bidi paragraph the start
    // (w:left, 16px) insets the physical RIGHT edge. Word-verified: with the
    // logical model wild2-lit-yiddish-rtl p126 (quote blocks, ind left up to
    // 4956tw) and staging-bidi both render 0.00 vs their PDFs; the physical
    // model scored 4.56 / 1.34.
    expect(bidiBase.x + bidiBase.width).toBeCloseTo(704, 3); // right edge - 16px logical start indent
    expect(ltrIndented.x - ltrBase.x).toBeCloseTo(48, 3);
    expect(
      bidiBase.x + bidiBase.width - (bidiIndented.x + bidiIndented.width),
    ).toBeCloseTo(48, 3);
  });

  it("resolves a first-line tab from the logical start without changing LTR tabs", () => {
    const paragraph = (bidi: boolean) =>
      `<w:p><w:pPr>${bidi ? "<w:bidi/>" : ""}` +
      `<w:tabs><w:tab w:val="left" w:pos="720"/></w:tabs>` +
      `<w:ind w:firstLine="240"/></w:pPr>` +
      ltrRun("2") +
      `<w:r>${bidi ? "<w:rPr><w:rtl/></w:rPr>" : ""}<w:tab/></w:r>` +
      (bidi ? rtlRun("נ".repeat(20)) : ltrRun("n".repeat(20))) +
      `</w:p>`;
    const resultFor = (bidi: boolean) =>
      layout({ "word/document.xml": wrapDocument(paragraph(bidi) + narrowSection) }).result;
    const lineCount = (result: ReturnType<typeof layoutDocument>) =>
      new Set(
        result.pages[0].items
          .filter((item) => item.kind === "text" && item.text.trim())
          .map((item) => (item.kind === "text" ? item.lineTop : 0)),
      ).size;

    const ltr = resultFor(false);
    const bidi = resultFor(true);
    expect(lineCount(ltr)).toBe(1);
    expect(lineCount(bidi)).toBe(1);

    const ltrText = ltr.pages[0].items.find(
      (item) => item.kind === "text" && item.text === "n".repeat(20),
    );
    const bidiMark = bidi.pages[0].items.find(
      (item) => item.kind === "text" && item.text === "2",
    );
    if (ltrText?.kind !== "text" || bidiMark?.kind !== "text") throw new Error();
    expect(ltrText.x).toBeCloseTo(144, 3);
    expect(bidiMark.x + bidiMark.width).toBeCloseTo(280, 3);
  });

  it("keeps a bidi numbering tab between the label and first-line text", () => {
    const numbering = `<?xml version="1.0"?>
      <w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:abstractNum w:abstractNumId="0">
          <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/>
            <w:lvlText w:val="—"/><w:lvlJc w:val="left"/></w:lvl>
        </w:abstractNum>
        <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
      </w:numbering>`;
    const paragraph = (bidi: boolean) =>
      `<w:p><w:pPr>${bidi ? "<w:bidi/>" : ""}` +
      `<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>` +
      `<w:tabs><w:tab w:val="left" w:pos="720"/></w:tabs>` +
      `<w:ind w:firstLine="240"/></w:pPr>` +
      (bidi ? rtlRun("נ".repeat(20)) : ltrRun("n".repeat(20))) +
      `</w:p>`;
    const resultFor = (bidi: boolean) =>
      layout({
        "word/document.xml": wrapDocument(paragraph(bidi) + narrowSection),
        "word/numbering.xml": numbering,
      }).result;
    const lineCount = (result: ReturnType<typeof layoutDocument>) =>
      new Set(
        result.pages[0].items
          .filter((item) => item.kind === "text" && item.text.trim())
          .map((item) => (item.kind === "text" ? item.lineTop : 0)),
      ).size;

    const ltr = resultFor(false);
    const bidi = resultFor(true);
    expect(lineCount(ltr)).toBe(1);
    expect(lineCount(bidi)).toBe(1);

    const ltrText = ltr.pages[0].items.find(
      (item) => item.kind === "text" && item.text === "n".repeat(20),
    );
    const bidiText = bidi.pages[0].items.find(
      (item) => item.kind === "text" && item.text === "נ".repeat(20),
    );
    const bidiLabel = bidi.pages[0].items.find(
      (item) => item.kind === "text" && item.text === "—",
    );
    if (
      ltrText?.kind !== "text" ||
      bidiText?.kind !== "text" ||
      bidiLabel?.kind !== "text"
    ) throw new Error();
    expect(ltrText.x).toBeCloseTo(144, 3);
    expect(bidiLabel.x - bidiText.x - bidiText.width).toBeCloseTo(24.667, 3);
    expect(bidiLabel.x + bidiLabel.width).toBeCloseTo(280, 3);
  });

  it("aligns a final justified line to the paragraph's logical start", () => {
    const resultFor = (bidi: boolean) =>
      layout({
        "word/document.xml": wrapDocument(
          `<w:p><w:pPr>${bidi ? "<w:bidi/>" : ""}<w:jc w:val="both"/></w:pPr>` +
          (bidi ? rtlRun("נ".repeat(10)) : ltrRun("n".repeat(10))) +
          `</w:p>` +
          narrowSection,
        ),
      }).result;
    const ltr = resultFor(false).pages[0].items.find(
      (item) => item.kind === "text" && item.text === "n".repeat(10),
    );
    const bidi = resultFor(true).pages[0].items.find(
      (item) => item.kind === "text" && item.text === "נ".repeat(10),
    );
    if (ltr?.kind !== "text" || bidi?.kind !== "text") throw new Error();

    expect(ltr.x).toBeCloseTo(96, 3);
    expect(bidi.x + bidi.width).toBeCloseTo(296, 3);
  });

  it("mirrors columns of a bidiVisual table (source col 0 lands on the right)", () => {
    const tbl =
      `<w:tbl><w:tblPr><w:bidiVisual/><w:tblW w:w="0" w:type="auto"/></w:tblPr>` +
      `<w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>` +
      `<w:tr>` +
      `<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>${p("AAA")}</w:tc>` +
      `<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>${p("BBB")}</w:tc>` +
      `</w:tr></w:tbl>`;
    const { result } = layout({ "word/document.xml": wrapDocument(tbl) });
    const texts = result.pages[0].items.filter((i) => i.kind === "text");
    const a = texts.find((i) => i.kind === "text" && i.text === "AAA");
    const b = texts.find((i) => i.kind === "text" && i.text === "BBB");
    if (a?.kind !== "text" || b?.kind !== "text") throw new Error();
    // Source order A,B -> visual order B,A (A on the right).
    expect(a.x).toBeGreaterThan(b.x);
  });
});

describe("anchored drawing position variants", () => {
  const WP = 'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"';
  const A = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';
  const WPS = 'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"';
  const WP14 = 'xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"';

  /** Anchored wps textbox run (mirrors scripts/make-staging-fixtures.py). */
  const anchorBox = (opts: {
    x: number; // EMU
    y: number; // EMU
    w: number; // EMU
    h: number; // EMU
    relH?: string;
    relV?: string;
    wrap?: "none" | "square";
    fill?: string;
    behind?: boolean;
    allowOverlap?: boolean;
    sizeRel?: { rel: string; pctW: number; pctH: number };
  }): string => {
    const wrap = opts.wrap === "square" ? '<wp:wrapSquare wrapText="bothSides"/>' : "<wp:wrapNone/>";
    const sizeRel = opts.sizeRel
      ? `<wp14:sizeRelH ${WP14} relativeFrom="${opts.sizeRel.rel}"><wp14:pctWidth>${opts.sizeRel.pctW}</wp14:pctWidth></wp14:sizeRelH>` +
        `<wp14:sizeRelV ${WP14} relativeFrom="${opts.sizeRel.rel}"><wp14:pctHeight>${opts.sizeRel.pctH}</wp14:pctHeight></wp14:sizeRelV>`
      : "";
    return (
      `<w:r><w:drawing><wp:anchor ${WP} distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="2" ` +
      `behindDoc="${opts.behind ? 1 : 0}" locked="0" layoutInCell="1" allowOverlap="${opts.allowOverlap === false ? 0 : 1}">` +
      `<wp:simplePos x="0" y="0"/>` +
      `<wp:positionH relativeFrom="${opts.relH ?? "column"}"><wp:posOffset>${opts.x}</wp:posOffset></wp:positionH>` +
      `<wp:positionV relativeFrom="${opts.relV ?? "paragraph"}"><wp:posOffset>${opts.y}</wp:posOffset></wp:positionV>` +
      `<wp:extent cx="${opts.w}" cy="${opts.h}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>` +
      wrap +
      `<wp:docPr id="9" name="Box"/><wp:cNvGraphicFramePr/>` +
      `<a:graphic ${A}><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">` +
      `<wps:wsp ${WPS}><wps:cNvSpPr/><wps:spPr>` +
      `<a:xfrm><a:off x="0" y="0"/><a:ext cx="${opts.w}" cy="${opts.h}"/></a:xfrm>` +
      `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>` +
      `<a:solidFill><a:srgbClr val="${opts.fill ?? "DDEEFF"}"/></a:solidFill>` +
      `</wps:spPr><wps:txbx><w:txbxContent><w:p><w:r><w:t>B</w:t></w:r></w:p></w:txbxContent></wps:txbx>` +
      `<wps:bodyPr rot="0" anchor="t"><a:noAutofit/></wps:bodyPr>` +
      `</wps:wsp></a:graphicData></a:graphic>` +
      sizeRel +
      `</wp:anchor></w:drawing></w:r>`
    );
  };
  const boxRect = (result: ReturnType<typeof layoutDocument>, fill: string) => {
    const r = result.pages[0].items.find((i) => i.kind === "rect" && i.fill === fill);
    if (r?.kind !== "rect") throw new Error(`missing box rect ${fill}`);
    return r;
  };

  it("resolves relH=character/relV=line from the anchor run's pen position and line top", () => {
    // wrapNone box: the paragraph's layout is identical with and without the
    // box, so its position must equal the following run's start exactly.
    const para =
      `<w:p><w:r><w:t xml:space="preserve">AB CD </w:t></w:r>` +
      anchorBox({ x: 0, y: 0, w: 914400, h: 457200, relH: "character", relV: "line", wrap: "none", fill: "AA0001" }) +
      `<w:r><w:t xml:space="preserve">ZZZ tail</w:t></w:r></w:p>`;
    const { result } = layout({ "word/document.xml": wrapDocument(para) });
    const rect = boxRect(result, "#aa0001");
    const zzz = result.pages[0].items.find((i) => i.kind === "text" && i.text.startsWith("ZZZ"));
    if (zzz?.kind !== "text") throw new Error("missing tail run");
    expect(rect.x).toBeCloseTo(zzz.x, 3);
    expect(rect.y).toBeCloseTo(zzz.lineTop, 3);
    // Sanity: the pen position is past the margin (not the column origin).
    expect(rect.x).toBeGreaterThan(100);
  });

  it("resolves relV=line against the anchor's own (wrapped) line, not the paragraph top", () => {
    // Force the anchor onto line 2 with a leading line break.
    const para =
      `<w:p><w:r><w:t>first line</w:t><w:br/><w:t xml:space="preserve">go </w:t></w:r>` +
      anchorBox({ x: 0, y: 0, w: 914400, h: 457200, relH: "character", relV: "line", wrap: "none", fill: "AA0002" }) +
      `<w:r><w:t>ZZZ</w:t></w:r></w:p>`;
    const { result } = layout({ "word/document.xml": wrapDocument(para) });
    const rect = boxRect(result, "#aa0002");
    const zzz = result.pages[0].items.find((i) => i.kind === "text" && i.text === "ZZZ");
    if (zzz?.kind !== "text") throw new Error("missing ZZZ");
    const first = result.pages[0].items.find((i) => i.kind === "text" && i.text.includes("first"));
    if (first?.kind !== "text") throw new Error("missing first line");
    expect(rect.y).toBeCloseTo(zzz.lineTop, 3);
    expect(rect.y).toBeGreaterThan(first.lineTop + 5);
  });

  it("sizes a wp14 pct box from the page and keeps its declared offsets", () => {
    const para =
      `<w:p>` +
      anchorBox({
        x: 914400, y: 0, w: 1645920, h: 731520, relH: "column", relV: "paragraph",
        wrap: "none", fill: "AA0003", sizeRel: { rel: "page", pctW: 20000, pctH: 8000 },
      }) +
      `<w:r><w:t>host</w:t></w:r></w:p>`;
    const { result } = layout({ "word/document.xml": wrapDocument(para) });
    const rect = boxRect(result, "#aa0003");
    const page = result.pages[0];
    // 20% of page width x 8% of page height override the extent.
    expect(rect.width).toBeCloseTo(0.2 * page.width, 1);
    expect(rect.height).toBeCloseTo(0.08 * page.height, 1);
    expect(rect.x).toBeCloseTo(96 + 96, 1); // margin + 1in posOffset
  });

  it("slides an allowOverlap=0 box right past earlier overlapping boxes", () => {
    const boxes =
      anchorBox({ x: 457200, y: 457200, w: 914400, h: 914400, relH: "page", relV: "page", wrap: "none", fill: "AA0004" }) +
      anchorBox({ x: 685800, y: 685800, w: 914400, h: 914400, relH: "page", relV: "page", wrap: "none", fill: "AA0005" }) +
      anchorBox({ x: 914400, y: 914400, w: 914400, h: 914400, relH: "page", relV: "page", wrap: "none", fill: "AA0006", allowOverlap: false });
    const { result } = layout({
      "word/document.xml": wrapDocument(`<w:p>${boxes}<w:r><w:t>text</w:t></w:r></w:p>`),
    });
    const b2 = boxRect(result, "#aa0005");
    const b3 = boxRect(result, "#aa0006");
    // Declared at 1in; slides right to the second box's right edge (Word:
    // staging-anchors2's z=30 locked no-overlap box lands at x=3.8in).
    expect(b3.x).toBeCloseTo(b2.x + b2.width, 1);
    expect(b3.y).toBeCloseTo(96, 1); // vertical position unchanged
  });

  it("wraps an earlier paragraph around an absolutely positioned float anchored later", () => {
    // Margin-anchored square box, anchored two paragraphs later: the FIRST
    // paragraph's lines must already flow beside it (Word reflows earlier
    // page content around page/margin-anchored floats).
    const host =
      `<w:p>` +
      anchorBox({ x: 0, y: 0, w: 914400, h: 914400, relH: "margin", relV: "margin", wrap: "square", fill: "AA0007" }) +
      `<w:r><w:t>host paragraph</w:t></w:r></w:p>`;
    const long = `<w:p><w:r><w:t xml:space="preserve">${"wrap me around the box ".repeat(20)}</w:t></w:r></w:p>`;
    const { result } = layout({ "word/document.xml": wrapDocument(long + p("middle") + host) });
    const rect = boxRect(result, "#aa0007");
    expect(rect.x).toBeCloseTo(96, 1);
    expect(rect.y).toBeCloseTo(96, 1);
    const items = result.pages[0].items;
    const beside = items.filter(
      (i) => i.kind === "text" && i.text.includes("wrap") && i.lineTop < rect.y + rect.height,
    );
    expect(beside.length).toBeGreaterThan(0);
    for (const t of beside) {
      if (t.kind !== "text") continue;
      expect(t.x).toBeGreaterThanOrEqual(rect.x + rect.width - 0.01);
    }
  });

  it("paints non-behind anchored boxes above body text and behindDoc boxes below", () => {
    const para =
      `<w:p>` +
      anchorBox({ x: 0, y: 0, w: 914400, h: 457200, relH: "page", relV: "page", wrap: "none", fill: "AA0008" }) +
      anchorBox({ x: 1828800, y: 0, w: 914400, h: 457200, relH: "page", relV: "page", wrap: "none", fill: "AA0009", behind: true }) +
      `<w:r><w:t>text</w:t></w:r></w:p>`;
    const { result } = layout({ "word/document.xml": wrapDocument(para) });
    const front = boxRect(result, "#aa0008");
    const behind = boxRect(result, "#aa0009");
    expect(front.front).toBe(true);
    expect(front.behind).toBeUndefined();
    expect(behind.behind).toBe(true);
    expect(behind.front).toBeUndefined();
  });
});

describe("wild2 legal agreement rules", () => {
  const SECT =
    `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>` +
    `<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>`;
  const textItem = (result: ReturnType<typeof layoutDocument>, text: string) => {
    const found = result.pages[0].items.find((i) => i.kind === "text" && i.text === text);
    if (found?.kind !== "text") throw new Error(`missing item ${JSON.stringify(text)}`);
    return found;
  };

  it("places the numbering label at a positive firstLine indent (direct ind overrides the level's hanging)", () => {
    // wild2-legal: numbered lvl ind left=2520 hanging=360 overridden by direct
    // ind left=0 firstLine=1530 puts "A." at 1530tw and the suffix tab runs to
    // the next default stop.
    const numbering =
      `<?xml version="1.0"?><w:numbering ${W_NS}>` +
      `<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/>` +
      `<w:numFmt w:val="upperLetter"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/>` +
      `<w:pPr><w:ind w:left="2520" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum>` +
      `<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`;
    const para =
      `<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>` +
      `<w:ind w:left="0" w:firstLine="1530"/></w:pPr>` +
      `<w:r><w:t>Body</w:t></w:r></w:p>`;
    const { result } = layout({
      "word/document.xml": wrapDocument(para + SECT),
      "word/numbering.xml": numbering,
    });
    const label = textItem(result, "A.");
    // margin 96px + firstLine 102px
    expect(label.x).toBeCloseTo(198, 1);
    // suffix tab from the label end to the next 0.5in default stop (144px rel)
    expect(textItem(result, "Body").x).toBeCloseTo(240, 1);
  });

  it("ignores whitespace-only runs when sizing a line with solid content", () => {
    // A lone oversized space run between normal words must not grow the line
    // (Word measured: 12pt space inside a 10pt body keeps the 10pt pitch).
    const bigSpace =
      `<w:p><w:r><w:t>aa</w:t></w:r>` +
      `<w:r><w:rPr><w:sz w:val="96"/></w:rPr><w:t xml:space="preserve"> </w:t></w:r>` +
      `<w:r><w:t>bb</w:t></w:r></w:p>`;
    const plain = `<w:p><w:r><w:t xml:space="preserve">aa bb</w:t></w:r></w:p>`;
    const next = `<w:p><w:r><w:t>next</w:t></w:r></w:p>`;
    const grown = layout({ "word/document.xml": wrapDocument(bigSpace + next + SECT) }).result;
    const control = layout({ "word/document.xml": wrapDocument(plain + next + SECT) }).result;
    expect(textItem(grown, "next").lineTop).toBeCloseTo(textItem(control, "next").lineTop, 2);
  });

  it("keeps the Symbol font's metrics for a substituted bullet label", () => {
    // The bullet paints via Unicode substitution in the body font but the
    // line is sized by the symbol font (Word: Symbol 10pt bullet -> 12.25pt
    // line where the body is 11.5pt).
    const symbolMeasurer: TextMeasurer = {
      width: (text, font, ls) => measurer.width(text, font, ls),
      metrics: (font) =>
        /symbol/i.test(font.family)
          ? { ascent: font.size * 1.1, descent: font.size * 0.3, lineHeight: font.size * 1.4 }
          : measurer.metrics(font),
    };
    const numbering =
      `<?xml version="1.0"?><w:numbering ${W_NS}>` +
      `<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/>` +
      `<w:numFmt w:val="bullet"/><w:lvlText w:val="&#xF0B7;"/><w:lvlJc w:val="left"/>` +
      `<w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>` +
      `<w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol"/></w:rPr></w:lvl></w:abstractNum>` +
      `<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`;
    const bullet =
      `<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>` +
      `<w:r><w:t xml:space="preserve"> </w:t></w:r></w:p>`;
    const next = `<w:p><w:r><w:t>next</w:t></w:r></w:p>`;
    const doc = DocxDocument.load(
      makeDocx({
        "word/document.xml": wrapDocument(bullet + next + SECT),
        "word/numbering.xml": numbering,
      }),
    );
    const result = layoutDocument(doc, { measurer: symbolMeasurer });
    const bulletItem = textItem(result, "•");
    const nextItem = textItem(result, "next");
    // the label inherits the default size; the SYMBOL metrics (1.4x) size the
    // line, not the painted body font's 1.15x
    expect(nextItem.lineTop - bulletItem.lineTop).toBeCloseTo(bulletItem.font.size * 1.4, 1);
  });

  it("reserves the full painted width of double table borders in row heights", () => {
    // A double rule paints two lines plus the gap = 3x the declared width.
    const table = (style: string) =>
      `<w:tbl><w:tblPr><w:tblW w:w="4000" w:type="dxa"/>` +
      `<w:tblBorders>` +
      ["top", "left", "bottom", "right", "insideH", "insideV"]
        .map((s) => `<w:${s} w:val="${style}" w:sz="6" w:space="0" w:color="auto"/>`)
        .join("") +
      `</w:tblBorders><w:tblLayout w:type="fixed"/></w:tblPr>` +
      `<w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>` +
      `<w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>R1</w:t></w:r></w:p></w:tc></w:tr>` +
      `<w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>R2</w:t></w:r></w:p></w:tc></w:tr>` +
      `</w:tbl><w:p/>`;
    const dbl = layout({ "word/document.xml": wrapDocument(table("double") + SECT) }).result;
    const sgl = layout({ "word/document.xml": wrapDocument(table("single") + SECT) }).result;
    const advance = (r: ReturnType<typeof layoutDocument>) =>
      textItem(r, "R2").lineTop - textItem(r, "R1").lineTop;
    // sz=6 -> 0.75pt = 1px painted single; double paints 3px. Row share grows
    // by (3-1)/2 at each of the row's two boundaries = 2px.
    expect(advance(dbl) - advance(sgl)).toBeCloseTo(2, 1);
  });

  it("tabs to the implicit stop at the left indent of a hanging-indent paragraph", () => {
    // Literal "4.<tab>" head with ind left=-450 hanging=270: the tab stops at
    // the left indent (-22.5pt), not at the margin's default grid.
    const para =
      `<w:p><w:pPr><w:ind w:left="-450" w:hanging="270"/></w:pPr>` +
      `<w:r><w:t>4.</w:t></w:r><w:r><w:tab/><w:t>Body</w:t></w:r></w:p>`;
    const { result } = layout({ "word/document.xml": wrapDocument(para + SECT) });
    // margin 96px + left indent (-30px) = 66px
    expect(textItem(result, "Body").x).toBeCloseTo(66, 1);
  });

  it("gives a document-opening empty paragraph before a table two mark lines", () => {
    // PDF-measured on wild2-legal p1: the top table's grid sits at margin +
    // two mark line heights; the same construct mid-flow takes one line.
    const emptyPara = `<w:p><w:pPr><w:rPr><w:sz w:val="24"/></w:rPr></w:pPr></w:p>`;
    const table =
      `<w:tbl><w:tblPr><w:tblW w:w="4000" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr>` +
      `<w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>` +
      `<w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>CELL</w:t></w:r></w:p></w:tc></w:tr>` +
      `</w:tbl><w:p/>`;
    const { result } = layout({ "word/document.xml": wrapDocument(emptyPara + table + SECT) });
    // margin 96px + 2 x (12pt = 16px x 1.15 = 18.4px) = 132.8px
    expect(textItem(result, "CELL").lineTop).toBeCloseTo(96 + 2 * 18.4, 1);
  });
});

describe("URL/long-token line breaking (wild2-legal-nih-contract corpus)", () => {
  // Word's in-token break rule, measured against every mid-token line break
  // on pp116-260 of the NIH contract's Word PDF: the ONLY soft break inside
  // an unspaced token is after a hyphen with alphanumerics on both sides
  // (digits included). '/', '_', '.', ':', '?', '=', '&' are never break
  // opportunities; with no opportunity Word breaks at the exact character
  // where the token crosses the line edge, even mid-line.
  const lineTexts = (result: ReturnType<typeof layoutDocument>, pageIdx = 0): string[] => {
    const rows = new Map<number, { x: number; text: string }[]>();
    for (const it of result.pages[pageIdx].items) {
      if (it.kind !== "text") continue;
      const key = Math.round(it.lineTop * 10);
      if (!rows.has(key)) rows.set(key, []);
      rows.get(key)!.push({ x: it.x, text: it.text });
    }
    return [...rows.entries()]
      .sort((a, b) => a[0] - b[0])
      .map(([, spans]) => spans.sort((a, b) => a.x - b.x).map((s) => s.text).join(""));
  };
  // 1000tw content column = 66.67px = 5.0em at sz 20 (10pt = 13.333px).
  const sect =
    `<w:sectPr><w:pgSz w:w="3880" w:h="15840"/>` +
    `<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>`;

  it("breaks after hyphens flanked by digits (identifier hyphens)", () => {
    // "GUF-JE-" = 4.3em fits; adding "04-" (1.5em) overflows -> break at the
    // hyphen op before the digit segment (NIH PDF p124 ".../GUF-JE-" |
    // "04-332.qigu", p153 ".../h44-" | "40.aki").
    const para =
      `<w:p><w:r><w:rPr><w:sz w:val="20"/></w:rPr><w:t>GUF-JE-04-33</w:t></w:r></w:p>`;
    const { result } = layout({ "word/document.xml": wrapDocument(para + sect) });
    const lines = lineTexts(result).filter((l) => l.trim().length > 0);
    expect(lines).toEqual(["GUF-JE-", "04-33"]);
  });

  it("emergency-breaks an NBSP-glued long token at the line edge in place", () => {
    // NIH PDF p154: "at:" + NBSP + a hyphen-free URL wider than a full line.
    // Word keeps "at:" on its line and fills URL characters to the exact
    // edge ("at:  wamuv://...BOB_HUG_Kudifup" | "a_Sucumo.idi"); it does NOT
    // flush "at:" alone and restart the token on a fresh line.
    const token = " " + "o".repeat(30); // 15.5em, line is 5em
    const para =
      `<w:p><w:r><w:rPr><w:sz w:val="20"/></w:rPr><w:t xml:space="preserve">at: ${token}</w:t></w:r></w:p>`;
    const { result } = layout({ "word/document.xml": wrapDocument(para + sect) });
    const lines = lineTexts(result).filter((l) => l.trim().length > 0);
    // First line carries "at:" AND the head of the glued token.
    expect(lines[0].startsWith("at:")).toBe(true);
    expect(lines[0]).toMatch(/o+$/);
    // No characters lost across the char-wrapped continuation lines.
    expect(lines.join("").replace(/\s| /g, "")).toBe("at:" + "o".repeat(30));
  });
});

describe("numbering fidelity (wild2-legal-nih-contract p177)", () => {
  it("keeps a numId's one-shot startOverride restart when a keepNext walk measures it first", () => {
    // The keepNext chain walk measures follower paragraphs via
    // numberingLabel(); its counter snapshot must also roll back seenNumIds
    // or the once-only startOverride restart fires during measurement and is
    // lost for the real placement (NIH p177: numId 340 rendered hh/ii/jj/kk
    // where Word restarts at a/b/c/d).
    const numberingXml = `<?xml version="1.0"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="0">
    <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/></w:lvl>
  </w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
  <w:num w:numId="2"><w:abstractNumId w:val="0"/>
    <w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride>
  </w:num>
</w:numbering>`;
    const numPara = (text: string, numId: number) =>
      `<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="${numId}"/></w:numPr></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`;
    const keeper = `<w:p><w:pPr><w:keepNext/></w:pPr><w:r><w:t>keeper</w:t></w:r></w:p>`;
    const { result } = layout({
      "word/document.xml": wrapDocument(
        numPara("one", 1) + numPara("two", 1) + keeper + numPara("restart", 2),
      ),
      "word/numbering.xml": numberingXml,
    });
    const labels = result.pages[0].items
      .filter((i) => i.kind === "text" && /^\d+\.$/.test(i.text))
      .map((i) => (i.kind === "text" ? i.text : ""));
    expect(labels).toEqual(["1.", "2.", "1."]);
  });

  it("right-aligns a lvlJc=right label at the number position, keeping text at ind.left", () => {
    // NIH p177 lowerRoman levels (lvlJc=right, ind left=2160 hanging=180):
    // every label's RIGHT edge sits at ind.left - hanging and the suffix-tab
    // text stays at ind.left even for wide labels ("viii.").
    const numberingXml = `<?xml version="1.0"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="0">
    <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="lowerRoman"/><w:lvlText w:val="%1."/><w:lvlJc w:val="right"/>
      <w:pPr><w:ind w:left="2160" w:hanging="180"/></w:pPr>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
</w:numbering>`;
    const numPara = (text: string) =>
      `<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`;
    const sect =
      `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>` +
      `<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>`;
    const { result } = layout({
      "word/document.xml": wrapDocument(
        numPara("alpha") + numPara("beta") + numPara("gamma") + sect,
      ),
      "word/numbering.xml": numberingXml,
    });
    const items = result.pages[0].items.filter((i) => i.kind === "text");
    const labels = items.filter((i) => i.kind === "text" && /^[ivx]+\.$/.test(i.text));
    expect(labels.length).toBe(3);
    // Number position: margin 96px + (2160 - 180)/15 = 228px. Right edges of
    // "i.", "ii.", "iii." all land there; text starts at ind.left = 240px.
    for (const l of labels) {
      if (l.kind !== "text") continue;
      expect(l.x + l.width).toBeCloseTo(228, 1);
    }
    for (const t of ["alpha", "beta", "gamma"]) {
      const item = items.find((i) => i.kind === "text" && i.text === t);
      expect(item && item.kind === "text" ? item.x : NaN).toBeCloseTo(240, 1);
    }  });
});

describe("OMML math line extents (wild2-math-omml-dense Word PDF rules)", () => {
  // Measured in parity/wild2-math-omml-dense-word.pdf at 12pt Cambria Math:
  //   - display fractions: numerator baseline +9.08pt / denominator -8.03pt
  //     (MATH constants 1550/2048, 1370/2048); text-style +7.03 / -6.04.
  //   - a display denominator holding bracket glyphs sits 0.78pt lower
  //     (row 5 den "2(1+h)" baseline +8.8 vs row 2 "2h" +8.03).
  //   - stretched delimiters swap in DISCRETE Cambria Math variants
  //     (11.12/14.50/19.80/23.71/30.60/35.47pt of ink at 12pt), centered on
  //     the math axis (585/2048 em), covering >= ~80% of the core extent;
  //     the variant ink defines the LINE box (rows pitch 29.5..31.3pt).
  //   - script protrusions never stretch delimiters ((0.8)'s parens stay
  //     regular around A_l p^(l+0)).
  const measurer = new ApproxMeasurer();
  const SIZE = 16; // 12pt in px
  const frac = (num: string, den: string): MathNode =>
    ({ t: "frac", num: [{ t: "run", text: num }], den: [{ t: "run", text: den }] });

  it("display fraction shifts follow the MATH constants", () => {
    const box = layoutMath([frac("1-2h", "2h")], SIZE, measurer, true);
    const num = box.pieces.find((p) => p.text.includes("1"))!;
    const den = box.pieces.find((p) => p.dy < 0)!;
    expect(num.dy).toBeCloseTo((SIZE * 1550) / 2048, 3); // +9.08pt @12
    expect(den.dy).toBeCloseTo((-SIZE * 1370) / 2048, 3); // -8.03pt @12
  });

  it("a bracketed display denominator drops one rule step lower", () => {
    const plain = layoutMath([frac("1-2h", "2h")], SIZE, measurer, true);
    const brack = layoutMath([frac("1-2h", "2(1+h)")], SIZE, measurer, true);
    const plainDen = plain.pieces.find((p) => p.dy < 0)!;
    const brackDen = brack.pieces.find((p) => p.dy < 0)!;
    expect(plainDen.dy - brackDen.dy).toBeCloseTo((SIZE * 133) / 2048, 3); // 0.78pt @12
  });

  it("a delimiter around a display fraction takes a discrete variant sized to ~80% coverage", () => {
    const box = layoutMath([{ t: "dlm", beg: "(", end: ")", e: [[frac("1-2h", "2h")]] }], SIZE, measurer, true);
    const paren = box.pieces.find((p) => p.text === "(")!;
    // ApproxMeasurer core: num 12.11+14.4 / den 10.70+4.0 -> H 41.21px,
    // 0.8H = 32.97 -> smallest variant >= that is 5223/2048 em = 40.80px.
    const variantH = (SIZE * 5223) / 2048;
    const axis = SIZE * (3.125 / 11);
    expect(paren.ownAscent).toBeCloseTo(axis + variantH / 2, 2);
    expect(paren.ownDescent).toBeCloseTo(variantH / 2 - axis, 2);
    // The variant ink drives the line DESCENT past the denominator box.
    expect(box.descent).toBeCloseTo(variantH / 2 - axis, 2);
  });

  it("script protrusions do not stretch delimiters", () => {
    const box = layoutMath(
      [{
        t: "dlm", beg: "(", end: ")",
        e: [[{ t: "sup", base: [{ t: "run", text: "p" }], script: [{ t: "run", text: "l+0" }] }]],
      }],
      SIZE, measurer, false,
    );
    const paren = box.pieces.find((p) => p.text === "(")!;
    expect(paren.ownAscent).toBeUndefined();
    expect(paren.scaleY).toBeUndefined();
  });

  it("a delimiter directly wrapping another regular delimiter takes the second size", () => {
    const box = layoutMath(
      [{ t: "dlm", beg: "(", end: ")", e: [[
        { t: "run", text: "n" },
        { t: "dlm", beg: "(", end: ")", e: [[{ t: "run", text: "n+1" }]] },
      ]] }],
      SIZE, measurer, false,
    );
    const outer = box.pieces.filter((p) => p.text === "(")[0];
    const variantH = (SIZE * 2475) / 2048; // 14.50pt @12: dense (0.1a)'s outer pair
    const axis = SIZE * (3.125 / 11);
    expect(outer.ownAscent).toBeCloseTo(axis + variantH / 2, 2);
  });

  it("display math rows carry a thin leading strip above the cluster", () => {
    // dense p13: every (6-2) row gap = prev desc + next asc + ~0.042em @12pt
    // (0.5pt), and the block's first baseline sits 21.7pt under the body
    // text line (2.65 text desc + 18.35 cluster asc + the strip).
    const M = `xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"`;
    const doc =
      p("before") +
      `<w:p><m:oMathPara ${M}><m:oMath><m:r><m:t>x=1</m:t></m:r></m:oMath></m:oMathPara></w:p>`;
    const { result } = layout({ "word/document.xml": wrapDocument(doc) });
    const items = result.pages[0].items.filter((i) => i.kind === "text");
    const before = items.find((i) => i.text === "before")!;
    const piece = items.find((i) => (i as { mathSrc?: unknown }).mathSrc)!;
    // ApproxMeasurer: prev text desc 0.25em + math cluster asc 0.9em +
    // the display lead 0.042em, quarter-pt snapped (all at the default size).
    const sz = before.font.size;
    expect(piece.baseline - before.baseline).toBeCloseTo(sz * (0.25 + 0.9 + 0.042), 0);
  });
});

describe("OMML matrices, arrays, accents, group chars, radicals, limits (probe2-math-matrices Word PDF)", () => {
  // Positions measured in parity/probe2-math-matrices-word.pdf at 11pt Cambria
  // Math. The unit metrics come from the deterministic ApproxMeasurer (the STIX
  // gap path), so these pins guard the layout STRUCTURE; the Word-real px
  // constants are calibrated by the parity harness.
  const measurer = new ApproxMeasurer();
  const S = 11;
  const run = (t: string): MathNode => ({ t: "run", text: t });
  const frac = (n: string, d: string): MathNode => ({ t: "frac", num: [run(n)], den: [run(d)] });

  it("matrix rows pitch by MAT_ROW_PITCH and columns hug the bracket edge", () => {
    const box = layoutMath(
      [{ t: "dlm", beg: "[", end: "]", e: [[{ t: "mat", rows: [[[run("a")], [run("b")]], [[run("c")], [run("d")]]] }]] }],
      S, measurer, true,
    );
    const open = box.pieces.find((p) => p.text === "[")!;
    const cells = box.pieces.filter((p) => !"[]".includes(p.text));
    const topDy = Math.max(...cells.map((p) => p.dy));
    const botDy = Math.min(...cells.map((p) => p.dy));
    const firstCol = cells.reduce((a, b) => (b.x < a.x ? b : a));
    expect(topDy - botDy).toBeCloseTo((S * 12.75) / 11, 2); // 12.75pt baseline pitch
    // First column starts AT the opening bracket's advance edge (hug, no pad).
    expect(firstCol.x).toBeCloseTo(open.x + measurer.width("[", { family: open.font.family, size: S, bold: false, italic: false }), 2);
    expect(open.scaleY).toBeGreaterThan(1); // grown bracket variant
    expect(open.ownAscent).toBeDefined();
  });

  it("a fraction-entry matrix pitches rows wider than the plain minimum", () => {
    const plain = layoutMath([{ t: "mat", rows: [[[run("a")]], [[run("b")]]] }], S, measurer, true);
    const tall = layoutMath([{ t: "mat", rows: [[[frac("1", "2")]], [[frac("1", "3")]]] }], S, measurer, true);
    const span = (b: typeof plain) => Math.abs(b.pieces[0].dy - b.pieces[b.pieces.length - 1].dy);
    expect(span(tall)).toBeGreaterThan((S * 12.75) / 11 + 1); // tall rows exceed the minimum pitch
  });

  it("a matrix delimiter too tall for the ladder is assembled to ~88% coverage", () => {
    const box = layoutMath(
      [{ t: "dlm", beg: "(", end: ")", e: [[{ t: "mat", rows: [
        [[frac("1", "2")]], [[frac("1", "5")]], [[frac("1", "8")]],
      ] }]] }],
      S, measurer, true,
    );
    const open = box.pieces.find((p) => p.text === "(")!;
    expect(open.scaleY).toBeGreaterThan(3); // assembled, well beyond a discrete step
  });

  it("eqArr stacks rows and the one-sided brace draws no closer", () => {
    const box = layoutMath(
      [run("f(x)="), { t: "dlm", beg: "{", end: "", e: [[{ t: "eqarr", rows: [[run("x,x")], [run("-x")]] }]] }],
      S, measurer, true,
    );
    const braces = box.pieces.filter((p) => p.text === "{" || p.text === "}");
    expect(braces.length).toBe(1); // only the opening brace
    expect(braces[0].text).toBe("{");
    expect(braces[0].scaleY).toBeGreaterThan(1);
  });

  it("suppresses the comma kern when the source already spells post-comma spaces", () => {
    // probe2-math-matrices' piecewise arms carry literal spaces after the comma
    // ('x², x ≥ 0'); Word draws just those spaces, so the synthetic COMMA_SPACE
    // kern must not stack on top of a following space (it still applies when the
    // comma is tight to the next atom, e.g. 'B(h,r,θ)').
    const mfont = { family: "Cambria Math", size: S, bold: false, italic: false };
    const w = (t: string) => measurer.width(t, mfont);
    const kerned = layoutMath([run("1,1")], S, measurer, true);
    const spaced = layoutMath([run("1, 1")], S, measurer, true);
    // kerned adds 0.17em after the comma; spaced relies on its literal space.
    expect(kerned.width - spaced.width).toBeCloseTo(S * 0.17 - w(" "), 1);
  });

  it("an accent composes onto its base glyph", () => {
    const box = layoutMath([{ t: "acc", chr: "̂", e: [run("x")] }], S, measurer, false);
    expect(box.pieces.length).toBe(1);
    expect(box.pieces[0].text.endsWith("̂")).toBe(true);
  });

  it("overbrace keeps a full-size base with a stretched brace above", () => {
    const box = layoutMath([{ t: "grp", chr: "⏞", pos: "top", vertJc: "bot", e: [run("a+b+c")] }], S, measurer, true);
    const brace = box.pieces.find((p) => p.text === "⏞")!;
    const base = box.pieces.find((p) => p.text !== "⏞")!;
    expect(base.dy).toBeCloseTo(0, 3); // base on the baseline
    expect(base.font.size).toBeCloseTo(S, 3); // full size
    expect(brace.scaleX).toBeGreaterThan(1); // stretched horizontally
    expect(brace.dy).toBeCloseTo((S * 3.4) / 11, 2); // brace raised above baseline
  });

  it("underbrace drops the base to script size raised over a brace on the baseline", () => {
    const box = layoutMath([{ t: "grp", chr: "⏟", pos: "bot", vertJc: "bot", e: [run("x+y")] }], S, measurer, true);
    const brace = box.pieces.find((p) => p.text === "⏟")!;
    const base = box.pieces.find((p) => p.text !== "⏟")!;
    expect(base.font.size).toBeLessThan(S); // script size
    expect(base.dy).toBeCloseTo((S * 8.5) / 11, 2); // raised above the brace
    expect(brace.dy).toBeCloseTo((S * 0.66) / 11, 2); // brace near the baseline
  });

  it("a degree index sits raised before the radical sign with a vinculum over the radicand", () => {
    const box = layoutMath([{ t: "rad", deg: [run("3")], e: [run("x+y")] }], S, measurer, false);
    const deg = box.pieces.find((p) => p.text === "3")!;
    const sign = box.pieces.find((p) => p.text === "√")!;
    expect(deg.dy).toBeGreaterThan(0); // degree raised
    expect(deg.x).toBeLessThan(sign.x); // before the sign
    expect(box.rules.length).toBe(1); // one vinculum
    expect(sign.scaleY).toBeGreaterThan(1);
  });

  it("a nested radical raises the outer vinculum above the inner one", () => {
    const box = layoutMath([{ t: "rad", e: [run("1+"), { t: "rad", e: [run("1+x")] }] }], S, measurer, false);
    const rules = box.rules.map((r) => r.dy).sort((a, b) => a - b);
    expect(rules.length).toBe(2);
    expect(rules[1]).toBeGreaterThan(rules[0]); // outer vinculum higher than inner
  });

  it("limLow drops the limit below and limUpp raises it above the operator", () => {
    const low = layoutMath([{ t: "lim", pos: "low", e: [run("lim")], lim: [run("n")] }], S, measurer, true);
    const upp = layoutMath([{ t: "lim", pos: "upp", e: [run("max")], lim: [run("y")] }], S, measurer, true);
    expect(low.pieces.find((p) => p.dy < 0)!.dy).toBeCloseTo((-S * 6.75) / 11, 2);
    expect(upp.pieces.find((p) => p.dy > 0)!.dy).toBeCloseTo((S * 7.75) / 11, 2);
  });
});

describe("m:oMathPara group justification (dense p7/p13 Word PDF)", () => {
  // Word lays a display equation broken into rows (w:br inside the math, or
  // auto-wrap) as one GROUP: rows left-align to each other; under the default
  // jc=centerGroup the group is centered on its widest row measured WITH its
  // trailing space runs (dense p13's group left is exactly colLeft +
  // (colW - widestRowWithSpaces)/2), and an auto-wrapped continuation row
  // indents a further wrapIndent (1440tw default) from the group's left edge.
  const M = `xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"`;
  const SECT =
    `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>` +
    `<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>`;
  const COL_LEFT = 96; // 1440tw margin
  const mathRowLefts = (result: ReturnType<typeof layoutDocument>): number[] => {
    const rows = new Map<number, number>();
    for (const i of result.pages[0].items) {
      if (i.kind !== "text" || !(i as { mathSrc?: unknown }).mathSrc) continue;
      rows.set(i.lineTop, Math.min(rows.get(i.lineTop) ?? Infinity, i.x));
    }
    return [...rows.entries()].sort((a, b) => a[0] - b[0]).map(([, x]) => x);
  };
  const eqPara = (paraPr: string, segments: string[]): string =>
    `<w:p><m:oMathPara ${M}>${paraPr}<m:oMath>` +
    segments
      .map((t, i) => `<m:r>${i > 0 ? "<w:br/>" : ""}<m:t xml:space="preserve">${t}</m:t></m:r>`)
      .join("") +
    `</m:oMath></m:oMathPara></w:p>`;

  it("centerGroup left-aligns explicit-break rows and counts trailing spaces in the group width", () => {
    const run = (tail: string) =>
      layout({ "word/document.xml": wrapDocument(eqPara("", ["x=aaaaaaaa", `+bb${tail}`]) + SECT) }).result;
    const tight = mathRowLefts(run(""));
    expect(tight).toHaveLength(2);
    // Rows align at one left edge, centered inside the column (not per-line).
    expect(tight[1]).toBeCloseTo(tight[0], 1);
    expect(tight[0]).toBeGreaterThan(COL_LEFT + 10);
    // 40 trailing spaces on the short row widen the GROUP: everything moves
    // left by half the added width, and the rows still share one edge.
    const spaced = mathRowLefts(run(" ".repeat(40)));
    expect(spaced[1]).toBeCloseTo(spaced[0], 1);
    expect(spaced[0]).toBeLessThan(tight[0] - 10);
  });

  it("m:oMathParaPr jc=left pins the group flush left", () => {
    const paraPr = `<m:oMathParaPr><m:jc m:val="left"/></m:oMathParaPr>`;
    const { result } = layout({
      "word/document.xml": wrapDocument(eqPara(paraPr, ["x=aaaaaaaa", "+bb"]) + SECT),
    });
    const lefts = mathRowLefts(result);
    expect(lefts).toHaveLength(2);
    expect(lefts[0]).toBeCloseTo(COL_LEFT, 1);
    expect(lefts[1]).toBeCloseTo(COL_LEFT, 1);
  });

  it("an auto-wrapped continuation row indents by wrapIndent from the group left", () => {
    // One long segment with top-level +'s: wider than the 624px column, so
    // wrapDisplayMath splits it; the continuation indents 96px (1440tw).
    const long = `${"a".repeat(40)}+${"b".repeat(40)}+${"c".repeat(40)}`;
    const { result } = layout({ "word/document.xml": wrapDocument(eqPara("", [long]) + SECT) });
    const lefts = mathRowLefts(result);
    expect(lefts.length).toBeGreaterThan(1);
    expect(lefts[1] - lefts[0]).toBeCloseTo(96, 1);
  });
});

describe("line metric rules (doerfp p8 / NIH p342)", () => {
  const SECT =
    `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>` +
    `<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>`;
  const topOf = (result: ReturnType<typeof layoutDocument>, text: string): number => {
    const found = result.pages[0].items.find((i) => i.kind === "text" && i.text === text);
    if (found?.kind !== "text") throw new Error(`missing item ${JSON.stringify(text)}`);
    return found.lineTop;
  };

  it("trailing tabs size the line by their own run props", () => {
    // wild-doerfp p8: a 10.5pt paragraph ending in four default-12pt tabs
    // keeps Word's 12pt line pitch; the trailing tabs' run props count.
    const body = `<w:r><w:rPr><w:sz w:val="21"/></w:rPr><w:t>body text here</w:t></w:r>`;
    const bigTab = `<w:r><w:rPr><w:sz w:val="40"/></w:rPr><w:tab/></w:r>`;
    const next = `<w:p><w:r><w:t>next</w:t></w:r></w:p>`;
    const withTab = layout({
      "word/document.xml": wrapDocument(`<w:p>${body}${bigTab}</w:p>` + next + SECT),
    }).result;
    const control = layout({
      "word/document.xml": wrapDocument(`<w:p>${body}</w:p>` + next + SECT),
    }).result;
    const grew =
      topOf(withTab, "next") - topOf(control, "next");
    const delta = measurer.metrics({ family: "Calibri", size: (40 / 2) * (96 / 72), bold: false, italic: false }).lineHeight -
      measurer.metrics({ family: "Calibri", size: (21 / 2) * (96 / 72), bold: false, italic: false }).lineHeight;
    expect(grew).toBeCloseTo(delta, 1);
  });

  it("an interior invisible tab still does not enlarge the line", () => {
    const next = `<w:p><w:r><w:t>next</w:t></w:r></w:p>`;
    const mid =
      `<w:p><w:r><w:rPr><w:sz w:val="21"/></w:rPr><w:t>alpha</w:t></w:r>` +
      `<w:r><w:rPr><w:sz w:val="40"/></w:rPr><w:tab/></w:r>` +
      `<w:r><w:rPr><w:sz w:val="21"/></w:rPr><w:t>omega</w:t></w:r></w:p>`;
    const plain =
      `<w:p><w:r><w:rPr><w:sz w:val="21"/></w:rPr><w:t>alpha omega</w:t></w:r></w:p>`;
    const a = layout({ "word/document.xml": wrapDocument(mid + next + SECT) }).result;
    const b = layout({ "word/document.xml": wrapDocument(plain + next + SECT) }).result;
    expect(topOf(a, "next")).toBeCloseTo(topOf(b, "next"), 2);
  });

  it("a numbering label shorter than the text leaves the line at the text pitch", () => {
    // NIH contract p342: Courier New "o" bullets among Calibri 12pt do not
    // register even Courier's larger descent - the line stays at the Calibri
    // pitch. Only a label TALLER than the text sizes the line (see the
    // Symbol-bullet test above).
    const courierMeasurer: TextMeasurer = {
      width: (text, font, ls) => measurer.width(text, font, ls),
      metrics: (font) =>
        /courier/i.test(font.family)
          ? { ascent: font.size * 0.9, descent: font.size * 0.45, lineHeight: font.size * 1.14 }
          : { ascent: font.size * 0.95, descent: font.size * 0.2, lineHeight: font.size * 1.15 },
    };
    const numbering =
      `<?xml version="1.0"?><w:numbering ${W_NS}>` +
      `<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/>` +
      `<w:numFmt w:val="bullet"/><w:lvlText w:val="o"/><w:lvlJc w:val="left"/>` +
      `<w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>` +
      `<w:rPr><w:rFonts w:ascii="Courier New" w:hAnsi="Courier New"/></w:rPr></w:lvl></w:abstractNum>` +
      `<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`;
    const bullet =
      `<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>` +
      `<w:r><w:t>itemtext</w:t></w:r></w:p>`;
    const next = `<w:p><w:r><w:t>next</w:t></w:r></w:p>`;
    const doc = DocxDocument.load(
      makeDocx({
        "word/document.xml": wrapDocument(bullet + next + SECT),
        "word/numbering.xml": numbering,
      }),
    );
    const result = layoutDocument(doc, { measurer: courierMeasurer });
    const items = result.pages[0].items.filter((i) => i.kind === "text");
    const itemLine = items.find((i) => i.kind === "text" && i.text === "itemtext");
    const nextLine = items.find((i) => i.kind === "text" && i.text === "next");
    if (itemLine?.kind !== "text" || nextLine?.kind !== "text") throw new Error("missing items");
    // Courier's 0.45em descent must NOT stretch the bullet line: the gap to
    // the next paragraph stays at the body font's 1.15em pitch.
    expect(nextLine.lineTop - itemLine.lineTop).toBeCloseTo(itemLine.font.size * 1.15, 1);
  });
});

describe("table row border share", () => {
  it("advances rows by the DECLARED sz-4 rule width (0.5pt), not the paint floor", () => {
    // phase23 p66: 45 single-line rows with sz-4 rules drift 2.5px down when
    // the 0.75px hairline paint floor leaks into the row advance.
    const rows = Array.from(
      { length: 2 },
      (_, i) => `<w:tr><w:tc><w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:t>row${i}</w:t></w:r></w:p></w:tc></w:tr>`,
    ).join("");
    const { result } = layout({
      "word/document.xml": wrapDocument(
        `<w:tbl><w:tblPr><w:tblBorders>
           <w:top w:val="single" w:sz="4"/><w:bottom w:val="single" w:sz="4"/>
           <w:left w:val="single" w:sz="4"/><w:right w:val="single" w:sz="4"/>
           <w:insideH w:val="single" w:sz="4"/><w:insideV w:val="single" w:sz="4"/>
         </w:tblBorders></w:tblPr>
         <w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid>${rows}</w:tbl>` + p("after"),
      ),
    });
    const texts = result.pages[0].items.filter((it) => it.kind === "text" && /^row/.test(it.text));
    if (texts.length !== 2 || texts[0].kind !== "text" || texts[1].kind !== "text") throw new Error("rows not found");
    const lineH = texts[0].font.size * 1.15; // ApproxMeasurer single line
    // Row pitch = content line + the sz-4 rule's TRUE width (0.5pt = 2/3px).
    expect(texts[1].lineTop - texts[0].lineTop).toBeCloseTo(lineH + 2 / 3, 2);  });
});

describe("cross-references and over-wide tables", () => {
  const FOOTER_RELS = `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdF" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>
</Relationships>`;

  it("re-renders REF fields from the bookmark range with document-order SEQ values", () => {
    // gatech's table of figures: a REF to a caption's _Ref bookmark is laid
    // out PAGES before the caption. Word recomputes both on open — the
    // cached "Bavoqe 0" (sanitizer-remapped digits) renders as the caption's
    // real "… 1". The REF must not consume the SEQ counter out of order.
    const refPara =
      `<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r>` +
      `<w:r><w:instrText xml:space="preserve"> REF _Ref123 \\h  \\* MERGEFORMAT </w:instrText></w:r>` +
      `<w:r><w:fldChar w:fldCharType="separate"/></w:r>` +
      `<w:r><w:t>Figure 9</w:t></w:r>` +
      `<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>`;
    const caption =
      `<w:p><w:bookmarkStart w:id="7" w:name="_Ref123"/>` +
      `<w:r><w:t xml:space="preserve">Figure </w:t></w:r>` +
      `<w:fldSimple w:instr=" SEQ Figure \\* ARABIC "><w:r><w:t>0</w:t></w:r></w:fldSimple>` +
      `<w:bookmarkEnd w:id="7"/>` +
      `<w:r><w:t xml:space="preserve"> caption tail</w:t></w:r></w:p>`;
    const { result } = layout({ "word/document.xml": wrapDocument(refPara + caption) });
    const text = pageText(result, 0);
    expect(text).toContain("Figure 1 caption tail"); // caption SEQ recomputed
    expect(text).not.toContain("Figure 9"); // stale REF cache replaced
    expect(text).not.toContain("Figure 0"); // fldSimple cache not doubled in
    // The REF occurrence itself renders the bookmark text.
    expect(text.indexOf("Figure 1")).toBeLessThan(text.indexOf("Figure 1 caption tail") + 1);
  });

  it("keeps a frame+trailing-empty footer bottom-anchored at footerDistance (no phantom line)", () => {
    // gatech footer2: widthless centered PAGE frame + final empty paragraph.
    // The frame overlays the empty follower and reserves NO extra height —
    // Word bottom-aligns the single line at footerDistance on the two-digit
    // pages too (the NIH phantom reserve only exists when painted content
    // FOLLOWS the overlaid follower).
    const body = Array.from({ length: 10 }, (_, i) =>
      `<w:p><w:r><w:t>Body ${i + 1}</w:t>${i < 9 ? '<w:br w:type="page"/>' : ""}</w:r></w:p>`,
    ).join("");
    const { result } = layout({
      "word/document.xml": wrapDocument(
        body +
          `<w:sectPr>
            <w:footerReference xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" w:type="default" r:id="rIdF"/>
            <w:pgSz w:w="12240" w:h="15840"/>
            <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:footer="720"/>
          </w:sectPr>`,
      ),
      "word/_rels/document.xml.rels": FOOTER_RELS,
      "word/footer1.xml": `<?xml version="1.0"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p><w:pPr><w:framePr w:wrap="around" w:vAnchor="text" w:hAnchor="margin" w:xAlign="center" w:y="1"/></w:pPr>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>PAGE</w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
  </w:p>
  <w:p></w:p>
</w:ftr>`,
    });
    expect(result.totalPages).toBe(10);
    const pageNumber = (page: number) => {
      const laidPage = result.pages[page - 1];
      const item = laidPage.items.slice(laidPage.hfStart).find(
        (candidate) => candidate.kind === "text" && candidate.text === String(page),
      );
      if (item?.kind !== "text") throw new Error(`missing page ${page} footer`);
      return item;
    };
    // Two-digit page 10 sits exactly where page 1 does: no phantom reserve.
    expect(pageNumber(10).lineTop).toBeCloseTo(pageNumber(1).lineTop, 3);
    // Bottom-anchored: the single footer line ends at pageHeight - footerDistance.
    const n1 = pageNumber(1);
    expect(n1.lineTop + n1.lineHeight).toBeCloseTo(15840 / 15 - 720 / 15, 0);
  });

  it("lets a body-level fixed-layout table keep its over-wide grid (Word overflows the margin)", () => {
    // ca-agreement p1: tblLayout=fixed with tblW/grid wider than the text
    // column renders at full grid width into the right margin, not shrunk.
    const tbl =
      `<w:tbl><w:tblPr><w:tblW w:w="10170" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr>` +
      `<w:tblGrid><w:gridCol w:w="4770"/><w:gridCol w:w="5400"/></w:tblGrid>` +
      `<w:tr><w:tc><w:tcPr><w:tcW w:w="4770" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>left</w:t></w:r></w:p></w:tc>` +
      `<w:tc><w:tcPr><w:tcW w:w="5400" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>right</w:t></w:r></w:p></w:tc></w:tr></w:tbl>` +
      `<w:p/>`;
    const sect =
      `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>` +
      `<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>`;
    const { result } = layout({ "word/document.xml": wrapDocument(tbl + sect) });
    const right = result.pages[0].items.find((i) => i.kind === "text" && i.text === "right");
    if (right?.kind !== "text") throw new Error("missing cell text");
    // Second column starts at marginLeft + 4770tw, NOT scaled into the
    // 9360tw content box (which would land it at 96 + 293.8px).
    expect(right.x).toBeGreaterThan(96 + 4770 / 15 - 1);
  });

  it("moves a row whole when a cell's first line would miss the split cut", () => {
    // parity2-nestedtables p2: three one-line cells fit above the cut but the
    // taller cell's first line does not — Word pushes the ENTIRE row to the
    // next page instead of stranding an empty cell beside painted neighbours.
    const filler =
      `<w:tbl><w:tblPr><w:tblW w:w="9360" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr>` +
      `<w:tblGrid><w:gridCol w:w="9360"/></w:tblGrid>` +
      `<w:tr><w:trPr><w:trHeight w:hRule="exact" w:val="13900"/></w:trPr>` +
      `<w:tc><w:tcPr><w:tcW w:w="9360" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>spacer</w:t></w:r></w:p></w:tc></w:tr>` +
      `<w:tr><w:tc><w:tcPr><w:tcW w:w="4680" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>Alpha</w:t></w:r></w:p></w:tc>` +
      `<w:tc><w:tcPr><w:tcW w:w="4680" w:type="dxa"/></w:tcPr>` +
      `<w:p><w:pPr><w:rPr><w:sz w:val="48"/></w:rPr></w:pPr><w:r><w:rPr><w:sz w:val="48"/></w:rPr><w:t>Bravo big</w:t></w:r></w:p>` +
      `<w:p><w:r><w:t>Bravo two</w:t></w:r></w:p><w:p><w:r><w:t>Bravo three</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p/>`;
    const sect =
      `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>` +
      `<w:pgMar w:top="720" w:right="1440" w:bottom="720" w:left="1440"/></w:sectPr>`;
    const { result } = layout({ "word/document.xml": wrapDocument(filler + sect) });
    expect(result.totalPages).toBeGreaterThan(1);
    expect(pageText(result, 0)).not.toContain("Alpha"); // not stranded beside an empty cell
    expect(pageText(result, 1)).toContain("Alpha");
    expect(pageText(result, 1)).toContain("Bravo big");
  });
});

describe("margin line numbers (w:lnNumType)", () => {
  // Word formats margin line numbers with the DEFAULT PARAGRAPH STYLE's
  // resolved run properties (docDefaults + Normal chain) overlaid with the
  // "line number" character style — NOT raw docDefaults. In the elsevier
  // template docDefaults say Calibri 11pt but the PDF prints the numbers in
  // Normal's Times New Roman 12pt, baseline-aligned with the body line
  // ('117' and its 12pt body line share the glyph top exactly; on a 14pt
  // heading line the 12pt number's top sits ~1.6pt lower — pure baseline
  // alignment, not line-box bottoming).
  const styles = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles ${W_NS}>
<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="24"/></w:rPr></w:style>
<w:style w:type="character" w:styleId="LineNumber"><w:name w:val="line number"/><w:basedOn w:val="DefaultParagraphFont"/></w:style>
</w:styles>`;
  const section =
    `<w:sectPr><w:lnNumType w:countBy="1" w:distance="240"/>` +
    `<w:pgSz w:w="12240" w:h="15840"/>` +
    `<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>`;

  it("uses the default paragraph style font and baseline-aligns the number", () => {
    const { result } = layout({
      "word/document.xml": wrapDocument(p("Hello world") + section),
      "word/styles.xml": styles,
    });
    const items = result.pages[0].items.filter((i) => i.kind === "text");
    const num = items.find((i) => i.kind === "text" && i.text === "1");
    const body = items.find((i) => i.kind === "text" && i.text.startsWith("Hello"));
    expect(num && num.kind === "text").toBeTruthy();
    expect(body && body.kind === "text").toBeTruthy();
    if (!num || num.kind !== "text" || !body || body.kind !== "text") return;
    // Normal's rPr, not docDefaults: Times New Roman 12pt (24 half-points = 16px).
    expect(num.font.family).toBe("Times New Roman");
    expect(num.font.size).toBeCloseTo(16, 5);
    // Baseline-aligned with the body line via the exact glyph box.
    expect(num.baseline).toBeCloseTo(body.baseline, 3);
    const m = measurer.metrics(num.font);
    expect(num.glyphTop).toBeCloseTo(num.baseline - m.ascent, 3);
    expect(num.glyphBoxH).toBeCloseTo(m.ascent + m.descent, 3);
    // Number sits in the left margin, right-aligned against the distance gap.
    expect(num.x + num.width).toBeLessThan(body.x);  });
});

describe("tail parity rules (textboxes/nccih/hf2/yiddish)", () => {
  const RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
  const OD_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

  it("re-applies a displaced paragraph's space-before below a topAndBottom band (parity2-textboxes p1)", () => {
    // Heading with spacing-before, then the anchor paragraph whose
    // wrapTopAndBottom box is predicted at the heading's bottom: the heading
    // reflows BELOW the band and Word re-applies its space-before there
    // (measured: band bottom + 12pt, not the +2px mid-paragraph fudge).
    const WP = 'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"';
    const A = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';
    const WPS = 'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"';
    const box =
      `<w:r><w:drawing><wp:anchor ${WP} distT="91440" distB="91440" distL="114300" distR="114300" simplePos="0" relativeHeight="2" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">` +
      `<wp:simplePos x="0" y="0"/>` +
      `<wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>` +
      `<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>` +
      `<wp:extent cx="3657600" cy="822960"/><wp:effectExtent l="0" t="0" r="0" b="0"/>` +
      `<wp:wrapTopAndBottom/>` +
      `<wp:docPr id="9" name="Box"/><wp:cNvGraphicFramePr/>` +
      `<a:graphic ${A}><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">` +
      `<wps:wsp ${WPS}><wps:cNvSpPr/><wps:spPr>` +
      `<a:xfrm><a:off x="0" y="0"/><a:ext cx="3657600" cy="822960"/></a:xfrm>` +
      `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>` +
      `<a:solidFill><a:srgbClr val="DDEEFF"/></a:solidFill>` +
      `</wps:spPr><wps:txbx><w:txbxContent><w:p><w:r><w:t>B</w:t></w:r></w:p></w:txbxContent></wps:txbx>` +
      `<wps:bodyPr rot="0" anchor="t"><a:noAutofit/></wps:bodyPr>` +
      `</wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>`;
    const render = (before: number) => {
      const heading =
        `<w:p><w:pPr><w:spacing w:before="${before}" w:after="0"/></w:pPr>` +
        `<w:r><w:t>Displaced heading</w:t></w:r></w:p>`;
      const anchorPara = `<w:p>${box}</w:p>`;
      const { result } = layout({
        "word/document.xml": wrapDocument(p("Intro text") + heading + anchorPara + p("Tail")),
      });
      const h = result.pages[0].items.find((i) => i.kind === "text" && i.text.includes("Displaced"));
      const rect = result.pages[0].items.find((i) => i.kind === "rect" && i.fill === "#ddeeff");
      if (h?.kind !== "text" || rect?.kind !== "rect") throw new Error("missing heading/box");
      // Gap between the band's box bottom and the displaced heading's line top
      // (the box itself anchors at the heading's undisplaced bottom, so it
      // moves with spacing-before; the GAP isolates the re-applied spacing).
      return h.lineTop - (rect.y + rect.height);
    };
    // before=240tw (16px) vs before=0: gap grows by max(2, 16) - 2 = 14px.
    expect(render(240) - render(0)).toBeCloseTo(14, 1);
  });

  it("spans a body-level fixed pct table over content width + edge cell margins (nccih p14)", () => {
    // tblW 5000 pct with tblLayout fixed: Word measures 100% against the
    // text column PLUS the table's left+right cell margins, rendering the
    // authored grid (content + 216tw) raw — rules at margin -/+ 7.2px.
    const content = 9360; // 12240 - 2*1440
    const grid = `<w:gridCol w:w="${content / 2 + 108}"/><w:gridCol w:w="${content / 2 + 108}"/>`;
    const cell = (w: number) =>
      `<w:tc><w:tcPr><w:tcW w:w="${w}" w:type="dxa"/>` +
      `<w:tcBorders><w:left w:val="single" w:sz="4" w:color="000000"/><w:right w:val="single" w:sz="4" w:color="000000"/></w:tcBorders>` +
      `</w:tcPr><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc>`;
    const tbl =
      `<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/><w:tblLayout w:type="fixed"/>` +
      `<w:tblCellMar><w:left w:w="108" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr>` +
      `<w:tblGrid>${grid}</w:tblGrid>` +
      `<w:tr>${cell(content / 2 + 108)}${cell(content / 2 + 108)}</w:tr></w:tbl><w:p/>`;
    const { result } = layout({ "word/document.xml": wrapDocument(tbl) });
    const edges = result.pages[0].items.filter((i) => i.kind === "edge");
    const xs = edges.flatMap((e) => (e.kind === "edge" ? [e.x1, e.x2] : []));
    // The grid renders RAW: rules span content (624px) + 2 x 7.2px cell
    // margins = 638.4px, not scaled down to the bare 624px column.
    expect(Math.max(...xs) - Math.min(...xs)).toBeCloseTo(624 + 14.4, 0);
  });

  it("lets an explicit tab stop before the left indent capture the numbering suffix tab (nccih p16)", () => {
    // ind left=1800 hanging=720 with a num-tab override at 1440: the bullet
    // sits at 1080tw and single-line TEXT at the 1440tw stop, not at the
    // 1800tw indent.
    const numbering = `<?xml version="1.0"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0">
    <w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="-"/><w:lvlJc w:val="left"/>
    <w:pPr><w:tabs><w:tab w:val="num" w:pos="1800"/></w:tabs><w:ind w:left="1800" w:hanging="360"/></w:pPr>
  </w:lvl></w:abstractNum>
  <w:num w:numId="8"><w:abstractNumId w:val="1"/></w:num>
</w:numbering>`;
    const rels = `<?xml version="1.0"?>
<Relationships xmlns="${RELS_NS}">
  <Relationship Id="rIdN" Type="${OD_REL}/numbering" Target="numbering.xml"/>
</Relationships>`;
    const item =
      `<w:p><w:pPr>` +
      `<w:numPr><w:ilvl w:val="0"/><w:numId w:val="8"/></w:numPr>` +
      `<w:tabs><w:tab w:val="clear" w:pos="1800"/><w:tab w:val="num" w:pos="1440"/></w:tabs>` +
      `<w:ind w:left="1800" w:hanging="720"/>` +
      `</w:pPr><w:r><w:t>Item text</w:t></w:r></w:p>`;
    const { result } = layout({
      "word/document.xml": wrapDocument(item),
      "word/_rels/document.xml.rels": rels,
      "word/numbering.xml": numbering,
    });
    const t = result.pages[0].items.find((i) => i.kind === "text" && i.text.includes("Item"));
    if (t?.kind !== "text") throw new Error("missing item text");
    // margin 96px + stop 1440tw (96px) — not the 1800tw (120px) indent.
    expect(t.x).toBeCloseTo(96 + 96, 1);
  });

  it("leaves even pages blank when evenAndOddHeaders is on and no even footer exists (staging-hf2 p2)", () => {
    const settings = `<?xml version="1.0"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:evenAndOddHeaders/>
</w:settings>`;
    const { result } = layout({
      "word/document.xml": wrapDocument(
        `<w:p><w:r><w:t>One</w:t><w:br w:type="page"/><w:t>Two</w:t></w:r></w:p>` +
          `<w:sectPr>
            <w:footerReference xmlns:r="${OD_REL.replace("/relationships", "/relationships")}" w:type="default" r:id="rIdF"/>
            <w:pgSz w:w="12240" w:h="15840"/>
            <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:footer="720"/>
          </w:sectPr>`,
      ),
      "word/_rels/document.xml.rels": `<?xml version="1.0"?>
<Relationships xmlns="${RELS_NS}">
  <Relationship Id="rIdF" Type="${OD_REL}/footer" Target="footer1.xml"/>
</Relationships>`,
      "word/footer1.xml": `<?xml version="1.0"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p><w:r><w:t>Default footer</w:t></w:r></w:p>
</w:ftr>`,
      "word/settings.xml": settings,
    });
    expect(result.totalPages).toBe(2);
    expect(pageText(result, 0)).toContain("Default footer"); // odd page keeps it
    expect(pageText(result, 1)).not.toContain("Default footer"); // even page is BLANK
  });

  it("keeps split digit runs in logical order inside an RTL line (yiddish p214 TOC)", () => {
    // "101" cached as runs "1" + "01": European Numbers take an EVEN bidi
    // level, so the spans keep their order (pre-fix the line-level reversal
    // painted "011").
    const rtl = (t: string) => `<w:r><w:rPr><w:rtl/></w:rPr><w:t xml:space="preserve">${t}</w:t></w:r>`;
    const para = `<w:p><w:pPr><w:bidi/></w:pPr>${rtl("שלום ")}${rtl("1")}${rtl("01")}</w:p>`;
    const { result } = layout({ "word/document.xml": wrapDocument(para) });
    const one = result.pages[0].items.find((i) => i.kind === "text" && i.text === "1");
    const oh = result.pages[0].items.find((i) => i.kind === "text" && i.text === "01");
    if (one?.kind !== "text" || oh?.kind !== "text") throw new Error("missing digit spans");
    // Visual order must read "101": the "1" span sits LEFT of the "01" span.
    expect(one.x).toBeLessThan(oh.x);
    expect(oh.x).toBeCloseTo(one.x + one.width, 1);
  });

  it("advances default tab stops on the settings.xml w:defaultTabStop grid (yiddish p214)", () => {
    const settings = `<?xml version="1.0"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:defaultTabStop w:val="708"/>
</w:settings>`;
    const para = `<w:p><w:r><w:t>a</w:t><w:tab/><w:t>b</w:t></w:r></w:p>`;
    const { result } = layout({
      "word/document.xml": wrapDocument(para),
      "word/settings.xml": settings,
    });
    const b = result.pages[0].items.find((i) => i.kind === "text" && i.text === "b");
    if (b?.kind !== "text") throw new Error("missing b");
    // 708tw = 47.2px grid, not the built-in 48px.
    expect(b.x).toBeCloseTo(96 + 47.2, 1);
  });
});

describe("mirror margins", () => {
  // settings.xml w:mirrorMargins: odd (recto) pages keep the gutter on the
  // left (content origin = left margin + gutter); even (verso) pages swap the
  // left/right margins and move the gutter to the inside (right) edge, so the
  // content origin drops to the (swapped) left margin with no gutter. Measured
  // from probe3-mirror-book's Word PDF: page 1 content x0 = 120px
  // (1080tw margin + 720tw gutter), page 2 x0 = 72px (1080tw margin only).
  const minTextX = (result: ReturnType<typeof layoutDocument>, pageIdx: number): number =>
    Math.min(
      ...result.pages[pageIdx].items
        .filter((i): i is Extract<typeof i, { kind: "text" }> => i.kind === "text")
        .map((i) => i.x),
    );
  const sectPr =
    '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>' +
    '<w:pgMar w:top="1440" w:right="1080" w:bottom="1440" w:left="1080"' +
    ' w:header="720" w:footer="720" w:gutter="720"/></w:sectPr>';
  const body = Array.from({ length: 90 }, () => p("Line")).join("") + sectPr;

  it("swaps the gutter to the inside edge on even pages", () => {
    const result = layout({
      "word/document.xml": wrapDocument(body),
      "word/settings.xml": `<w:settings ${W_NS}><w:mirrorMargins/></w:settings>`,
    }).result;
    expect(result.pages.length).toBeGreaterThanOrEqual(2);
    expect(minTextX(result, 0)).toBeCloseTo(120, 1); // recto: margin + gutter
    expect(minTextX(result, 1)).toBeCloseTo(72, 1); // verso: margin only
  });

  it("without w:mirrorMargins the gutter stays left on every page", () => {
    const result = layout({ "word/document.xml": wrapDocument(body) }).result;
    expect(result.pages.length).toBeGreaterThanOrEqual(2);
    expect(minTextX(result, 0)).toBeCloseTo(120, 1);
    expect(minTextX(result, 1)).toBeCloseTo(120, 1);
  });
});

describe("prstTxWarp WordArt", () => {
  const WP = 'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"';
  const A = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';
  const WPS = 'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"';
  const warpBox = (prst: string) =>
    `<w:p><w:r><w:drawing><wp:anchor ${WP} distT="0" distB="0" distL="0" distR="0" simplePos="0" ` +
    `relativeHeight="2" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">` +
    `<wp:simplePos x="0" y="0"/>` +
    `<wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>` +
    `<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>` +
    `<wp:extent cx="2743200" cy="1463040"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>` +
    `<wp:docPr id="9" name="WordArt"/><wp:cNvGraphicFramePr/>` +
    `<a:graphic ${A}><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">` +
    `<wps:wsp ${WPS}><wps:cNvSpPr/><wps:spPr>` +
    `<a:xfrm><a:off x="0" y="0"/><a:ext cx="2743200" cy="1463040"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="2E74B5"/></a:solidFill>` +
    `</wps:spPr><wps:txbx><w:txbxContent><w:p><w:pPr><w:jc w:val="center"/></w:pPr>` +
    `<w:r><w:rPr><w:b/><w:color w:val="FFFFFF"/><w:sz w:val="40"/></w:rPr><w:t>ARCH UP</w:t></w:r>` +
    `</w:p></w:txbxContent></wps:txbx>` +
    `<wps:bodyPr rot="0" anchor="ctr"><a:prstTxWarp prst="${prst}"><a:avLst/></a:prstTxWarp></wps:bodyPr>` +
    `</wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r></w:p>`;

  it("bends the shape text onto a warptext item instead of flowing it", () => {
    const { result } = layout({ "word/document.xml": wrapDocument(warpBox("textArchUp")) });
    const items = result.pages[0].items;
    const warp = items.find((i) => i.kind === "warptext");
    if (warp?.kind !== "warptext") throw new Error("expected a warptext item");
    expect(warp.warp).toBe("textArchUp");
    expect(warp.text).toBe("ARCH UP");
    expect(warp.fill.toLowerCase()).toBe("#ffffff");
    expect(warp.bold).toBe(true);
    expect(warp.fontSize).toBeGreaterThan(0);
    // The box fill still paints; the flowed text lines do NOT (warp replaces them).
    expect(items.some((i) => i.kind === "rect" && i.fill === "#2e74b5")).toBe(true);
    expect(items.some((i) => i.kind === "text" && i.text.includes("ARCH"))).toBe(false);
  });

  it("treats textNoShape as no warp (text flows normally)", () => {
    const { result } = layout({ "word/document.xml": wrapDocument(warpBox("textNoShape")) });
    const items = result.pages[0].items;
    expect(items.some((i) => i.kind === "warptext")).toBe(false);
    expect(items.some((i) => i.kind === "text" && i.text.includes("ARCH"))).toBe(true);
  });
});
