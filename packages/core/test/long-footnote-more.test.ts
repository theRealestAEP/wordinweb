import { describe, expect, it } from "vitest";
import { DocxDocument } from "../src/docx.js";
import { layoutDocument, layoutDocumentStream } from "../src/layout/engine.js";
import { ApproxMeasurer } from "../src/layout/measure.js";
import type { LaidOutPage, TextItem } from "../src/layout/types.js";
import { openStreamingDocument } from "../src/stream/body.js";
import { bytesSource } from "../src/stream/source.js";
import { makeDocx, wrapDocument } from "./helpers.js";

const PAGE = `<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440"/>`;
const LETTER = `<w:sectPr>${PAGE}</w:sectPr>`;
const SPACING = `<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>`;
const FONT = `<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="20"/></w:rPr>`;

function para(text: string, reference = false): string {
  return `<w:p><w:pPr>${SPACING}</w:pPr><w:r>${FONT}<w:t>${text}</w:t></w:r>` +
    (reference ? `<w:r>${FONT}<w:footnoteReference w:id="1"/></w:r>` : "") + `</w:p>`;
}

function labels(prefix: string, count: number): string[] {
  return Array.from({ length: count }, (_, i) => `${prefix} ${String(i + 1).padStart(3, "0")}`);
}

function paragraphs(prefix: string, count: number): string {
  return labels(prefix, count).map((label) => para(label)).join("");
}

function documentBytes(body: string, note: string): Uint8Array {
  return makeDocx({
    "word/document.xml": wrapDocument(body),
    "word/_rels/document.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdFn" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>`,
    "word/footnotes.xml": `<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>
      <w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>
      <w:footnote w:id="1">${note}</w:footnote></w:footnotes>`,
  });
}

function lay(bytes: Uint8Array) {
  return layoutDocument(DocxDocument.load(bytes), { measurer: new ApproxMeasurer() });
}

// Join text items by line top, splitting columns at the gap between them.
function lines(page: LaidOutPage) {
  const byTop = new Map<number, TextItem[]>();
  for (const item of page.items) {
    if (item.kind !== "text") continue;
    const key = Math.round(item.lineTop * 100) / 100;
    byTop.set(key, [...(byTop.get(key) ?? []), item]);
  }
  const runs: TextItem[][] = [];
  for (const items of byTop.values()) {
    items.sort((a, b) => a.x - b.x);
    let run: TextItem[] = [];
    for (const item of items) {
      const last = run[run.length - 1];
      if (last && item.x - last.x - last.width > 20) {
        runs.push(run);
        run = [];
      }
      run.push(item);
    }
    runs.push(run);
  }
  return runs.map((items) => ({
    text: items.map((item) => item.text).join("").trim(),
    top: items[0].lineTop,
    bottom: Math.max(...items.map((item) => item.lineTop + item.lineHeight)),
    column: items[0].x < page.width / 2 ? 0 : 1,
  })).sort((a, b) => a.column - b.column || a.top - b.top);
}

function noteLabels(pages: LaidOutPage[]): string[][] {
  return pages.map((page) => lines(page).filter((line) => /^Note \d{3}$/.test(line.text)).map((line) => line.text));
}

describe("more long footnotes", () => {
  it("streams the same pages and note labels as whole-document layout", async () => {
    const bytes = documentBytes(para("Body reference", true) + paragraphs("Body", 60) + LETTER, paragraphs("Note", 160));
    const whole = lay(bytes);
    const opened = await openStreamingDocument(bytesSource(bytes));
    const measurer = new ApproxMeasurer();
    const final = await layoutDocumentStream(opened.doc, { measurer, prescan: opened.prescan, onPage: () => {} });
    const pages: LaidOutPage[] = [];
    const rendered = await layoutDocumentStream(opened.doc, {
      measurer, prescan: opened.prescan, final, onPage: (page) => { pages.push(page); },
    });
    expect(final.totalPages).toBe(whole.totalPages);
    expect(rendered.totalPages).toBe(whole.totalPages);
    expect(pages).toHaveLength(whole.pages.length);
    expect(noteLabels(pages)).toEqual(noteLabels(whole.pages));
    expect(noteLabels(pages).flat()).toEqual(labels("Note", 160));
    expect(noteLabels(pages).filter((page) => page.length > 0).length).toBeGreaterThanOrEqual(3);
  });

  it("keeps all 60 table rows with their bounding rules inside the body box", () => {
    const borders = ["top", "bottom", "left", "right", "insideH"].map((side) =>
      `<w:${side} w:val="single" w:sz="8" w:color="43729D"/>`).join("");
    const rows = labels("Row", 60).map((label) =>
      `<w:tr><w:trPr><w:trHeight w:val="360" w:hRule="exact"/></w:trPr><w:tc><w:tcPr><w:tcW w:w="9360" w:type="dxa"/></w:tcPr>${para(label)}</w:tc></w:tr>`).join("");
    const table = `<w:tbl><w:tblPr><w:tblW w:w="9360" w:type="dxa"/><w:tblBorders>${borders}</w:tblBorders><w:tblCellMar><w:top w:w="60" w:type="dxa"/><w:bottom w:w="60" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid><w:gridCol w:w="9360"/></w:tblGrid>${rows}</w:tbl>`;
    const result = lay(documentBytes(para("Body reference", true) + LETTER, para("Note introduction") + table));
    expect(result.pages.flatMap((page) => lines(page).filter((line) => /^Row \d{3}$/.test(line.text)).map((line) => line.text)))
      .toEqual(labels("Row", 60));
    expect(result.pages.length).toBeGreaterThan(1);
    for (const page of result.pages) {
      const rules = page.items.filter((item) => item.kind === "edge" && item.role === "table-rule" && Math.abs(item.y1 - item.y2) < 0.01);
      for (const line of lines(page).filter((line) => /^(Row \d{3}|Note introduction)$/.test(line.text))) {
        expect(line.top).toBeGreaterThanOrEqual(page.bodyTop - 0.5);
        expect(line.bottom).toBeLessThanOrEqual(page.bodyBottom + 0.5);
        if (!line.text.startsWith("Row")) continue;
        // The cell padding puts each bounding rule within 24px of the line top.
        expect(rules.some((rule) => rule.kind === "edge" && rule.y1 <= line.top + 0.5 && line.top - rule.y1 < 24),
          `${line.text}: upper rule on page ${page.index}`).toBe(true);
        expect(rules.some((rule) => rule.kind === "edge" && rule.y1 >= line.bottom - 0.5 && rule.y1 - line.top <= 24.5),
          `${line.text}: lower rule on page ${page.index}`).toBe(true);
      }
    }
  });

  it("balances two columns with an 80-line note before a continuous one-column section", () => {
    const twoColumns = `<w:sectPr><w:type w:val="continuous"/>${PAGE}<w:cols w:num="2" w:space="720"/></w:sectPr>`;
    const oneColumn = `<w:sectPr><w:type w:val="continuous"/>${PAGE}<w:cols w:num="1"/></w:sectPr>`;
    const body = para("Body reference", true) + paragraphs("Body", 20) +
      `<w:p><w:pPr>${twoColumns}</w:pPr></w:p>` + para("Body after columns") + oneColumn;
    const result = lay(documentBytes(body, paragraphs("Note", 80)));
    expect(noteLabels(result.pages).flat()).toEqual(labels("Note", 80));
    expect(result.pages.some((page) => page.columnBands.some((band) => band.colXs.length === 2) &&
      page.columnBands.some((band) => band.colXs.length === 1))).toBe(true);
    for (const page of result.pages) {
      const all = lines(page);
      for (const note of all.filter((line) => line.text.startsWith("Note"))) {
        expect(note.top).toBeGreaterThanOrEqual(page.bodyTop - 0.5);
        expect(note.bottom).toBeLessThanOrEqual(page.bodyBottom + 0.5);
        for (const bodyLine of all.filter((line) => line.text.startsWith("Body") && line.column === note.column)) {
          expect(note.bottom <= bodyLine.top + 0.5 || note.top >= bodyLine.bottom - 0.5).toBe(true);
        }
      }
    }
  });

  it("keeps at least two wrapped paragraph lines on each side of a note split", () => {
    const words = Array.from({ length: 440 }, (_, i) => `word${String(i + 1).padStart(3, "0")}`).join(" ");
    const note = para(words).replace(SPACING, `${SPACING}<w:widowControl/>`);
    // 45 body lines leave about ten note lines below the reference on Letter.
    const result = lay(documentBytes(paragraphs("Body", 44) + para("Body reference", true) + LETTER, note));
    const parts = result.pages.map((page) => lines(page).filter((line) => line.text.startsWith("word")));
    expect(parts).toHaveLength(2);
    expect(parts[0].length).toBeGreaterThanOrEqual(8);
    expect(parts[0].length).toBeLessThanOrEqual(12);
    expect(parts.flat().length).toBeGreaterThanOrEqual(38);
    expect(parts.flat().length).toBeLessThanOrEqual(42);
    for (const part of parts) expect(part.length).toBeGreaterThanOrEqual(2);
    expect(parts.flat().map((line) => line.text).join(" ").split(/\s+/)).toEqual(words.split(" "));
  });

  it("keeps a continued note's picture and final paragraph when one column becomes two", async () => {
    const { makeDocxWithMedia } = await import("./helpers.js");
    const { unzipSync, strFromU8 } = await import("fflate");
    const twoColumns = `<w:sectPr><w:type w:val="continuous"/>${PAGE}<w:cols w:num="2" w:space="720"/></w:sectPr>`;
    const body = para("Body reference", true) + `<w:p><w:pPr>${LETTER}</w:pPr></w:p>` +
      para("Body after columns") + twoColumns;
    const picture = `<w:p><w:pPr>${SPACING}</w:pPr><w:r><w:drawing>
      <wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">
        <wp:extent cx="914400" cy="914400"/><wp:docPr id="1" name="Note picture"/>
        <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
            <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
              <pic:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rIdImg"/></pic:blipFill>
              <pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm></pic:spPr>
            </pic:pic>
          </a:graphicData>
        </a:graphic>
      </wp:inline></w:drawing></w:r></w:p>`;
    const note = paragraphs("Note", 47) + picture + para("Note final paragraph");
    const parts = Object.fromEntries(Object.entries(unzipSync(documentBytes(body, note)))
      .map(([name, bytes]) => [name, strFromU8(bytes)]));
    parts["[Content_Types].xml"] = parts["[Content_Types].xml"].replace("</Types>",
      `<Default Extension="png" ContentType="image/png"/></Types>`);
    parts["word/_rels/footnotes.xml.rels"] = `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdImg" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/note.png"/></Relationships>`;
    const png = Uint8Array.from(atob("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+aV1sAAAAASUVORK5CYII="), (char) => char.charCodeAt(0));
    const result = lay(makeDocxWithMedia(parts, { "word/media/note.png": png }));

    expect(noteLabels(result.pages).flat()).toEqual(labels("Note", 47));
    expect(result.pages.flatMap((page) => page.items.filter((item) => item.kind === "image"))).toHaveLength(1);
    expect(result.pages.findIndex((page) => page.items.some((item) => item.kind === "image")))
      .toBeGreaterThan(result.pages.findIndex((page) => noteLabels([page])[0].length > 0));
    expect(result.pages[0].columnBands[0].colXs).toHaveLength(1);
    expect(result.pages.at(-1)?.columnBands[0].colXs).toHaveLength(2);
    expect(result.pages.flatMap((page) => lines(page).filter((line) => line.text === "Note final paragraph"))).toHaveLength(1);
  });

  it("paints every repeated digit once when a continued note changes from two columns to one", () => {
    const twoColumns = `<w:sectPr>${PAGE}<w:cols w:num="2" w:space="720"/></w:sectPr>`;
    const oneColumn = `<w:sectPr><w:type w:val="continuous"/>${PAGE}<w:cols w:num="1"/></w:sectPr>`;
    const body = para("Body reference", true) + `<w:p><w:pPr>${twoColumns}</w:pPr></w:p>` +
      para("Body after columns") + oneColumn;
    const digits = "0123456789".repeat(300);
    const result = lay(documentBytes(body, para(digits)));
    const painted = result.pages.map((page) => page.items.flatMap((item) =>
      item.kind === "text" && item.src?.t?.text === digits ? [item.text] : []).join(""));

    expect(painted.filter((text) => text.length > 0).length).toBeGreaterThan(1);
    expect(result.pages[0].columnBands[0].colXs).toHaveLength(2);
    expect(result.pages.at(-1)?.columnBands[0].colXs).toHaveLength(1);
    expect(painted.join("")).toBe(digits);
  });

  it("paints a note picture once in streamed and whole layout across one, two, and one columns", async () => {
    const { makeDocxWithMedia } = await import("./helpers.js");
    const { unzipSync, strFromU8 } = await import("fflate");
    const twoColumns = `<w:sectPr><w:type w:val="continuous"/>${PAGE}<w:cols w:num="2" w:space="720"/></w:sectPr>`;
    const oneColumn = `<w:sectPr><w:type w:val="continuous"/>${PAGE}<w:cols w:num="1"/></w:sectPr>`;
    const body = paragraphs("Body", 40) + para("Body reference", true) +
      `<w:p><w:pPr>${LETTER}</w:pPr></w:p>` + para("Body in two columns") +
      `<w:p><w:pPr>${twoColumns}</w:pPr></w:p>` + para("Body after columns") + oneColumn;
    const picture = `<w:p><w:pPr>${SPACING}</w:pPr><w:r><w:drawing>
      <wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">
        <wp:extent cx="914400" cy="914400"/><wp:docPr id="1" name="Note picture"/>
        <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
            <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
              <pic:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rIdImg"/></pic:blipFill>
              <pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm></pic:spPr>
            </pic:pic>
          </a:graphicData>
        </a:graphic>
      </wp:inline></w:drawing></w:r></w:p>`;
    const parts = Object.fromEntries(Object.entries(unzipSync(documentBytes(body, picture + paragraphs("Note", 100))))
      .map(([name, bytes]) => [name, strFromU8(bytes)]));
    parts["[Content_Types].xml"] = parts["[Content_Types].xml"].replace("</Types>",
      `<Default Extension="png" ContentType="image/png"/></Types>`);
    parts["word/_rels/footnotes.xml.rels"] = `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdImg" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/note.png"/></Relationships>`;
    const png = Uint8Array.from(atob("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+aV1sAAAAASUVORK5CYII="), (char) => char.charCodeAt(0));
    const bytes = makeDocxWithMedia(parts, { "word/media/note.png": png });
    const whole = lay(bytes);
    const opened = await openStreamingDocument(bytesSource(bytes));
    const measurer = new ApproxMeasurer();
    const final = await layoutDocumentStream(opened.doc, { measurer, prescan: opened.prescan, onPage: () => {} });
    const pages: LaidOutPage[] = [];
    await layoutDocumentStream(opened.doc, {
      measurer, prescan: opened.prescan, final, onPage: (page) => { pages.push(page); },
    });

    expect(whole.pages[0].columnBands.map((band) => band.colXs.length)).toEqual([1, 2, 1]);
    expect(pages[0].columnBands.map((band) => band.colXs.length)).toEqual([1, 2, 1]);
    expect(noteLabels(whole.pages).flat()).toEqual(labels("Note", 100));
    expect(noteLabels(pages).flat()).toEqual(labels("Note", 100));
    expect(whole.pages.flatMap((page) => page.items.filter((item) => item.kind === "image"))).toHaveLength(1);
    expect(pages.flatMap((page) => page.items.filter((item) => item.kind === "image"))).toHaveLength(1);
  });

  it("keeps a one-line footnote on the reference page when a table row splits", () => {
    const cell = para("Row reference", true) + paragraphs("Row", 49);
    const table = `<w:tbl><w:tblPr><w:tblW w:w="9360" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="9360"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="9360" w:type="dxa"/></w:tcPr>${cell}</w:tc></w:tr></w:tbl>`;
    const result = lay(documentBytes(paragraphs("Body", 15) + table + LETTER, para("One-line note")));
    const referencePages = result.pages.filter((page) => page.items.some((item) => item.kind === "text" && item.noteId === 1));
    const notePages = result.pages.filter((page) => lines(page).some((line) => line.text === "One-line note"));

    expect(result.pages.filter((page) => lines(page).some((line) => /^Row \d{3}$/.test(line.text))).length).toBeGreaterThan(1);
    expect(referencePages).toHaveLength(1);
    expect(notePages).toHaveLength(1);
    expect(notePages[0].index).toBe(referencePages[0].index);
  });

  it("splits a row below a repeated header when its footnote leaves no room for a cut", () => {
    // The 200px body has only 60px left below the 140px repeated header.
    // Twelve 15px lines fit a fresh body, so splitting relies on treating
    // the position below repeated headers as the column top.
    const section = `<w:sectPr><w:pgSz w:w="12240" w:h="3600"/><w:pgMar w:top="300" w:bottom="300" w:left="1440" w:right="1440"/></w:sectPr>`;
    const cellPr = `<w:tcPr><w:tcW w:w="9360" w:type="dxa"/></w:tcPr>`;
    const header = `<w:tr><w:trPr><w:tblHeader/><w:trHeight w:val="2100" w:hRule="exact"/></w:trPr><w:tc>${cellPr}${para("Repeated header")}</w:tc></w:tr>`;
    const rowText = (para("Row reference", true) + paragraphs("Row", 11))
      .replaceAll(SPACING, `<w:spacing w:before="0" w:after="0" w:line="225" w:lineRule="exact"/>`);
    const row = `<w:tr><w:tc>${cellPr}${rowText}</w:tc></w:tr>`;
    const table = `<w:tbl><w:tblPr><w:tblW w:w="9360" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="9360"/></w:tblGrid>${header}${row}</w:tbl>`;
    const result = lay(documentBytes(table + section, paragraphs("Note", 4)));
    const rowPages = result.pages.map((page) => lines(page).filter((line) => line.text.startsWith("Row")));

    expect(rowPages.flat().filter((line) => /^Row \d{3}$/.test(line.text)).map((line) => line.text)).toEqual(labels("Row", 11));
    expect(rowPages.flat().filter((line) => line.text.startsWith("Row reference"))).toHaveLength(1);
    expect(noteLabels(result.pages).flat()).toEqual(labels("Note", 4));
    for (const page of result.pages) {
      for (const line of lines(page)) {
        expect(line.bottom, `${line.text}: bottom on page ${page.index}`).toBeLessThanOrEqual(page.bodyBottom + 0.5);
      }
    }
    expect(rowPages.filter((rows) => rows.length > 0).length).toBeGreaterThan(1);
  });

  it("keeps the text of three bordered 400px note rows inside page body boxes", () => {
    const borders = ["top", "bottom", "left", "right", "insideH"].map((side) =>
      `<w:${side} w:val="single" w:sz="8" w:color="43729D"/>`).join("");
    const rows = labels("Tall row", 3).map((label) =>
      `<w:tr><w:trPr><w:trHeight w:val="6000" w:hRule="exact"/></w:trPr><w:tc><w:tcPr><w:tcW w:w="9360" w:type="dxa"/></w:tcPr>${para(label)}</w:tc></w:tr>`).join("");
    const table = `<w:tbl><w:tblPr><w:tblW w:w="9360" w:type="dxa"/><w:tblBorders>${borders}</w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="9360"/></w:tblGrid>${rows}</w:tbl>`;
    const result = lay(documentBytes(para("Body reference", true) + LETTER, table));
    const rowPages = result.pages.map((page) => lines(page).filter((line) => /^Tall row \d{3}$/.test(line.text)));

    expect(rowPages.flat().map((line) => line.text)).toEqual(labels("Tall row", 3));
    expect(rowPages.filter((page) => page.length > 0).length).toBeGreaterThan(1);
    for (const [index, rows] of rowPages.entries()) {
      for (const row of rows) {
        expect(row.top, `${row.text}: top on page ${result.pages[index].index}`).toBeGreaterThanOrEqual(result.pages[index].bodyTop - 0.5);
        expect(row.bottom, `${row.text}: bottom on page ${result.pages[index].index}`).toBeLessThanOrEqual(result.pages[index].bodyBottom + 0.5);
      }
    }
  });
});
