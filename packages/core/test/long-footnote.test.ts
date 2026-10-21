// @vitest-environment jsdom
import { describe, expect, it } from "vitest";
import { DocxDocument } from "../src/docx.js";
import { layoutDocument } from "../src/layout/engine.js";
import { ApproxMeasurer } from "../src/layout/measure.js";
import type { LayoutResult, TextItem } from "../src/layout/types.js";
import { renderToDom } from "../src/render/dom.js";
import { makeDocx, wrapDocument } from "./helpers.js";

// A footnote longer than the space under its reference splits at a line and
// continues at the foot of the next page, under the continuation separator,
// before that page's own notes. The body never overlaps the notes.

const RELS = `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdFn" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>
  <Relationship Id="rIdEn" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" Target="endnotes.xml"/>
</Relationships>`;

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const SPACING = `<w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>`;
const NOTE_RUN = `<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="20"/></w:rPr>`;
const LETTER = `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440"/></w:sectPr>`;

/** `count` one-line note paragraphs, "<prefix> 001: ..." upward. */
function noteParas(prefix: string, count: number): string {
  return Array.from(
    { length: count },
    (_, i) =>
      `<w:p>${SPACING}<w:r>${NOTE_RUN}<w:t>${prefix} ${String(i + 1).padStart(3, "0")}: This cited decision supports the legal analysis.</w:t></w:r></w:p>`,
  ).join("");
}

function bodyPara(text: string, refs: number[] = []): string {
  const marks = refs
    .map((id) => `<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="${id}"/></w:r>`)
    .join("");
  return `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r>${marks}</w:p>`;
}

function filler(from: number, to: number): string {
  let xml = "";
  for (let k = from; k < to; k++) xml += bodyPara(`Body paragraph ${String(k).padStart(3, "0")} fills one line.`);
  return xml;
}

function footnotes(notes: Record<number, string>): string {
  return (
    `<w:footnotes ${W}>` +
    `<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>` +
    `<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>` +
    Object.entries(notes)
      .map(([id, xml]) => `<w:footnote w:id="${id}">${xml}</w:footnote>`)
      .join("") +
    `</w:footnotes>`
  );
}

function load(body: string, notes: Record<number, string>): DocxDocument {
  return DocxDocument.load(
    makeDocx({
      "word/document.xml": wrapDocument(body + LETTER),
      "word/_rels/document.xml.rels": RELS,
      "word/footnotes.xml": footnotes(notes),
    }),
  );
}

function lay(body: string, notes: Record<number, string>): LayoutResult {
  return layoutDocument(load(body, notes), { measurer: new ApproxMeasurer() });
}

const isNoteText = (text: string) => /^(Authority|Alpha|Beta|Short|Tall) \d{3}/.test(text);
const isBodyText = (text: string) =>
  /^(Body paragraph|Opening|Reference|First reference|Second reference|A later|A legal|The source|This body|Row line)/.test(text);

/** A painted line: the text items that share a line top in one column,
 * joined. `column` is 0 left of the page's middle, 1 right of it. */
interface Line {
  text: string;
  top: number;
  bottom: number;
  column: number;
}

function lines(result: LayoutResult, page: number): Line[] {
  const middle = result.pages[page].width / 2;
  const byTop = new Map<number, TextItem[]>();
  for (const it of result.pages[page].items) {
    if (it.kind !== "text") continue;
    const key = Math.round(it.lineTop * 100) / 100;
    byTop.set(key, [...(byTop.get(key) ?? []), it]);
  }
  // Items on one line top split into one line per column at a wide gap.
  const runs: TextItem[][] = [];
  for (const items of byTop.values()) {
    items.sort((a, b) => a.x - b.x);
    let run: TextItem[] = [];
    for (const it of items) {
      const last = run[run.length - 1];
      if (last && it.x - (last.x + last.width) > 20) {
        runs.push(run);
        run = [];
      }
      run.push(it);
    }
    runs.push(run);
  }
  return runs
    .map((items) => ({
      text: items.map((it) => it.text).join("").trim(),
      top: items[0].lineTop,
      bottom: Math.max(...items.map((it) => it.lineTop + it.lineHeight)),
      column: items[0].x < middle ? 0 : 1,
    }))
    .sort((a, b) => a.column - b.column || a.top - b.top);
}

/** The note lines on each page, in paint order, as "Alpha 001" labels. */
function noteLines(result: LayoutResult): string[][] {
  return result.pages.map((_, page) =>
    lines(result, page)
      .filter((line) => isNoteText(line.text))
      .map((line) => line.text.slice(0, line.text.indexOf(":"))),
  );
}

function pageOf(result: LayoutResult, needle: string): number {
  return result.pages.findIndex((_, page) => lines(result, page).some((line) => line.text.includes(needle)));
}

/** Every note line sits inside the page's body box, below every body line
 * of its column, and the note lines of a column do not overlap one another. */
function expectNotesClearOfBody(result: LayoutResult): void {
  result.pages.forEach((page, index) => {
    for (const column of [0, 1]) {
      const all = lines(result, index).filter((line) => line.column === column);
      const body = all.filter((line) => isBodyText(line.text));
      const notes = all.filter((line) => !body.includes(line));
      const bodyBottom = Math.max(page.bodyTop, ...body.map((line) => line.bottom));
      for (const note of notes) {
        expect(note.top).toBeGreaterThanOrEqual(bodyBottom - 0.5);
        expect(note.bottom).toBeLessThanOrEqual(page.bodyBottom + 0.5);
      }
      for (let i = 1; i < notes.length; i++) {
        expect(notes[i].top).toBeGreaterThanOrEqual(notes[i - 1].bottom - 0.5);
      }
    }
  });
}

function labels(prefix: string, count: number): string[] {
  return Array.from({ length: count }, (_, i) => `${prefix} ${String(i + 1).padStart(3, "0")}`);
}

describe("long footnotes", () => {
  // The reviewer's document: one footnote of 80 one-line paragraphs on Letter
  // paper. It used to paint upward from the page foot as one block, so its
  // first 17 paragraphs fell above the page and the rest overlapped the body.
  const reviewerBody =
    `<w:p>${SPACING}<w:r>${NOTE_RUN}<w:t>A legal memorandum with one extended footnote.</w:t></w:r></w:p>` +
    `<w:p><w:r><w:t>The source authorities are set out in the following footnote.</w:t></w:r><w:r><w:footnoteReference w:id="1"/></w:r></w:p>` +
    `<w:p>${SPACING}<w:r>${NOTE_RUN}<w:t>This body paragraph should remain readable above the note.</w:t></w:r></w:p>`;

  it("continues an 80-paragraph footnote on the next page with nothing lost", () => {
    const result = lay(reviewerBody, { 1: noteParas("Authority", 80) });
    expect(result.totalPages).toBe(2);
    const notes = noteLines(result);
    expect(notes.flat()).toEqual(labels("Authority", 80));
    // The note starts on its reference's page and continues on the next.
    expect(notes[0].length).toBeGreaterThan(30);
    expect(notes[1].length).toBeGreaterThan(0);
    expect(pageOf(result, "following footnote")).toBe(0);
    expectNotesClearOfBody(result);
  });

  it("opens the continued part with a rule across the column", () => {
    const result = lay(reviewerBody, { 1: noteParas("Authority", 80) });
    const rules = (page: number) =>
      result.pages[page].items.filter((it) => it.kind === "edge" && Math.abs(it.y1 - it.y2) < 0.01);
    const firstNote = lines(result, 1).find((line) => isNoteText(line.text));
    if (!firstNote) throw new Error("no continued note");
    const width = result.pages[1].width - 192;
    expect(
      rules(1).some((it) => it.kind === "edge" && Math.abs(it.x2 - it.x1 - width) < 1 && it.y1 < firstNote.top),
    ).toBe(true);
    // The reference page keeps Word's short separator.
    expect(rules(0).some((it) => it.kind === "edge" && Math.abs(it.x2 - it.x1 - 192) < 1)).toBe(true);
  });

  it("continues a note longer than two pages across several pages", () => {
    const body = bodyPara("Opening paragraph with the long note.", [1]) + filler(1, 61);
    const result = lay(body, { 1: noteParas("Alpha", 160) });
    const notes = noteLines(result);
    expect(notes.flat()).toEqual(labels("Alpha", 160));
    expect(notes.filter((page) => page.length > 0).length).toBeGreaterThanOrEqual(4);
    // Every body paragraph is still there, once, in order.
    const bodyLines = result.pages.flatMap((_, i) => lines(result, i).map((line) => line.text));
    const expected = Array.from({ length: 60 }, (_, k) => `Body paragraph ${String(k + 1).padStart(3, "0")} fills one line.`);
    expect(bodyLines.filter((text) => text.startsWith("Body paragraph"))).toEqual(expected);
    expectNotesClearOfBody(result);
  });

  it("keeps two long notes on one page in order, the second after the first", () => {
    const body = bodyPara("First reference.", [1]) + bodyPara("Second reference.", [2]) + filler(1, 41);
    const result = lay(body, { 1: noteParas("Alpha", 60), 2: noteParas("Beta", 60) });
    const notes = noteLines(result);
    expect(notes.flat()).toEqual([...labels("Alpha", 60), ...labels("Beta", 60)]);
    // The second note starts on the page where the first one ends.
    const alphaEnd = notes.findIndex((page) => page.includes("Alpha 060"));
    expect(notes[alphaEnd]).toContain("Beta 001");
    // Each reference shares a page with at least the first line of its note.
    expect(notes[pageOf(result, "First reference")]).toContain("Alpha 001");
    expect(notes[pageOf(result, "Second reference")]).toContain("Beta 001");
    expectNotesClearOfBody(result);
  });

  it("keeps a note's first lines under a reference on the last line of the page", () => {
    const body = filler(1, 43) + bodyPara("Reference at the foot of the page.", [1]) + filler(43, 60);
    const result = lay(body, { 1: noteParas("Alpha", 60) });
    const notes = noteLines(result);
    expect(notes.flat()).toEqual(labels("Alpha", 60));
    const refPage = pageOf(result, "Reference at the foot");
    expect(notes[refPage][0]).toBe("Alpha 001");
    // The reference is the last body line of its page.
    const body0 = lines(result, refPage).filter((line) => !isNoteText(line.text));
    expect(body0[body0.length - 1].text).toContain("Reference at the foot");
    expectNotesClearOfBody(result);
  });

  it("moves the reference line when not even the note's first line fits", () => {
    // The reference falls on the page's last body line, where no note line
    // fits under it: line and note move to the next page together.
    const body = filler(1, 50) + bodyPara("Reference past the foot.", [1]) + filler(50, 55);
    const result = lay(body, { 1: noteParas("Alpha", 60) });
    const refPage = pageOf(result, "Reference past the foot");
    expect(refPage).toBe(1);
    expect(pageOf(result, "Body paragraph 049")).toBe(0);
    expect(noteLines(result)[0]).toEqual([]);
    expect(noteLines(result)[refPage][0]).toBe("Alpha 001");
    expect(noteLines(result).flat()).toEqual(labels("Alpha", 60));
    expectNotesClearOfBody(result);
  });

  it("puts a short note on the continuation page after the continued part", () => {
    const body =
      bodyPara("Opening paragraph with the long note.", [1]) +
      filler(1, 8) +
      bodyPara("A later paragraph with a short note.", [2]) +
      filler(8, 50);
    const result = lay(body, { 1: noteParas("Alpha", 60), 2: noteParas("Short", 2) });
    const notes = noteLines(result);
    expect(notes.flat()).toEqual([...labels("Alpha", 60), ...labels("Short", 2)]);
    const shortPage = pageOf(result, "A later paragraph");
    expect(shortPage).toBeGreaterThan(0);
    // The short note shares its page with the end of the continued note and follows it.
    expect(notes[shortPage].slice(-3)).toEqual(["Alpha 060", "Short 001", "Short 002"]);
    expectNotesClearOfBody(result);
  });

  it("keeps a note with no line to split at whole instead of adding pages forever", () => {
    // One exact-spaced line taller than the page: nothing to split at.
    const tall = `<w:p><w:pPr><w:spacing w:line="20000" w:lineRule="exact"/></w:pPr><w:r><w:t>Tall 001: one line</w:t></w:r></w:p>`;
    const result = lay(bodyPara("Reference.", [1]) + filler(1, 5), { 1: tall });
    expect(result.totalPages).toBeLessThanOrEqual(3);
  });

  it("starts the note area of a table row's reference below the whole row", () => {
    const cell = bodyPara("Row line 01 with a note.", [1]) + Array.from({ length: 9 }, (_, i) => bodyPara(`Row line ${String(i + 2).padStart(2, "0")}`)).join("");
    const table =
      `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="9000"/></w:tblGrid>` +
      `<w:tr><w:trPr><w:cantSplit/></w:trPr><w:tc><w:tcPr><w:tcW w:w="9000" w:type="dxa"/></w:tcPr>${cell}</w:tc></w:tr></w:tbl>`;
    const result = lay(filler(1, 11) + table + filler(11, 15), { 1: noteParas("Alpha", 35) });
    expect(noteLines(result).flat()).toEqual(labels("Alpha", 35));
    result.pages.forEach((_, index) => {
      const all = lines(result, index);
      const rowBottom = Math.max(0, ...all.filter((line) => line.text.startsWith("Row line")).map((line) => line.bottom));
      for (const note of all.filter((line) => isNoteText(line.text))) expect(note.top).toBeGreaterThanOrEqual(rowBottom - 0.5);
    });
  });

  const TWO_COLUMNS = `<w:sectPr><w:type w:val="continuous"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440"/><w:cols w:num="2" w:space="720"/></w:sectPr>`;

  function layRaw(body: string, notes: Record<number, string>): LayoutResult {
    const doc = DocxDocument.load(
      makeDocx({
        "word/document.xml": wrapDocument(body),
        "word/_rels/document.xml.rels": RELS,
        "word/footnotes.xml": footnotes(notes),
      }),
    );
    return layoutDocument(doc, { measurer: new ApproxMeasurer() });
  }

  it("fills the next columns with a note that continues past the end of the body", () => {
    const result = layRaw(bodyPara("Reference in two columns.", [1]) + TWO_COLUMNS.replace("<w:type w:val=\"continuous\"/>", ""), {
      1: noteParas("Alpha", 100),
    });
    expect(noteLines(result).flat()).toEqual(labels("Alpha", 100));
    // The note continues in the second column, not only on later pages.
    const secondColumn = lines(result, 0).filter((line) => line.column === 1 && isNoteText(line.text));
    expect(secondColumn.length).toBeGreaterThan(0);
    expectNotesClearOfBody(result);
  });

  it("continues a note at the new width below a two-column band", () => {
    // A one-column section whose note continues into a continuous two-column section.
    const body =
      bodyPara("Opening paragraph with the long note.", [1]) +
      `<w:p><w:pPr><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440"/></w:sectPr></w:pPr></w:p>` +
      filler(1, 120) +
      TWO_COLUMNS;
    const result = layRaw(body, { 1: noteParas("Alpha", 100) });
    const notes = noteLines(result).flat();
    // Lines re-laid at the narrower width wrap differently, so labels may repeat
    // a line's start; none is lost and the order holds.
    expect([...new Set(notes)]).toEqual(labels("Alpha", 100));
    for (const page of result.pages) {
      for (const it of page.items) if (it.kind === "text") expect(it.x + it.width).toBeLessThanOrEqual(page.width - 96 + 1);
    }
    expectNotesClearOfBody(result);
  });

  it("leaves a long endnote flowing after the body as before", () => {
    const doc = DocxDocument.load(
      makeDocx({
        "word/document.xml": wrapDocument(
          `<w:p><w:r><w:t>Opening.</w:t></w:r><w:r><w:endnoteReference w:id="1"/></w:r></w:p>` + filler(1, 4) + LETTER,
        ),
        "word/_rels/document.xml.rels": RELS,
        "word/endnotes.xml":
          `<w:endnotes ${W}><w:endnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:endnote>` +
          `<w:endnote w:id="1">${noteParas("Alpha", 80)}</w:endnote></w:endnotes>`,
      }),
    );
    const result = layoutDocument(doc, { measurer: new ApproxMeasurer() });
    const notes = noteLines(result);
    expect(notes.flat()).toEqual(labels("Alpha", 80));
    // The same pages as the layout from before footnotes could split.
    expect(result.totalPages).toBe(2);
    expect(notes.map((page) => page.length)).toEqual([50, 30]);
    expect(lines(result, 0).find((line) => isNoteText(line.text))?.top).toBeCloseTo(181.8636, 3);
    // Endnotes flow in the body: no continuation rule opens page 2.
    expect(result.pages[1].items.some((it) => it.kind === "edge")).toBe(false);
  });

  it("paints every line of a continued note inside its page in the DOM", () => {
    const doc = load(reviewerBody, { 1: noteParas("Authority", 80) });
    const result = layoutDocument(doc, { measurer: new ApproxMeasurer() });
    const container = document.createElement("div");
    document.body.appendChild(container);
    renderToDom(doc, result, container);
    const pages = [...container.querySelectorAll<HTMLElement>(".dxw-page")];
    expect(pages.length).toBe(2);
    const seen: string[] = [];
    pages.forEach((pageEl, index) => {
      const spans = [...pageEl.querySelectorAll<HTMLElement>("span")].filter((span) => span.style.top !== "");
      const byTop = new Map<string, string>();
      for (const span of spans) byTop.set(span.style.top, (byTop.get(span.style.top) ?? "") + (span.textContent ?? ""));
      for (const [top, text] of [...byTop].sort((a, b) => parseFloat(a[0]) - parseFloat(b[0]))) {
        const label = text.trim().match(/^Authority \d{3}/);
        if (!label) continue;
        seen.push(label[0]);
        expect(parseFloat(top)).toBeGreaterThanOrEqual(result.pages[index].bodyTop);
        expect(parseFloat(top)).toBeLessThan(result.pages[index].bodyBottom);
      }
    });
    expect(seen).toEqual(labels("Authority", 80));
  });
});
