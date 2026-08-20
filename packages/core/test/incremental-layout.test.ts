import { describe, expect, it } from "vitest";
import { DocxDocument } from "../src/docx.js";
import { __incrStats, layoutDocument } from "../src/layout/engine.js";
import { invalidateParagraphSignature } from "../src/layout/inline.js";
import { ApproxMeasurer } from "../src/layout/measure.js";
import { setListType } from "../src/edit/lists.js";
import type { Paragraph, Run, TextContent } from "../src/model.js";
import type { XmlElement } from "../src/xml.js";
import { localName } from "../src/xml.js";
import { makeDocx, p, wrapDocument } from "./helpers.js";

const measurer = new ApproxMeasurer();
const section =
  `<w:sectPr><w:pgSz w:w="7200" w:h="10000"/>` +
  `<w:pgMar w:top="360" w:right="360" w:bottom="360" w:left="360"/></w:sectPr>`;

function denseDoc(): DocxDocument {
  const body = Array.from({ length: 96 }, (_, i) => p(`block-${i} alpha bravo charlie delta`)).join("");
  return DocxDocument.load(makeDocx({ "word/document.xml": wrapDocument(body + section) }));
}

function editParagraph(doc: DocxDocument, index: number, suffix: string): { block: Paragraph; source: NonNullable<TextContent["srcT"]> } {
  const block = doc.sections[0].blocks[index] as Paragraph;
  const run = block.children[0] as Run;
  const text = run.content[0] as TextContent;
  if (!text.srcT || !block.src) throw new Error("expected retained paragraph text");
  text.text += suffix;
  text.srcT.text = text.text;
  invalidateParagraphSignature(block.src);
  return { block, source: text.srcT };
}

function samePageSparseTarget(result: ReturnType<typeof layoutDocument>): number {
  const points = (result._incr as { points: Array<{ blockIdx: number; pageCount: number }> }).points;
  for (let start = 16; start < 64; start += 16) {
    const before = points.find((point) => point.blockIdx === start);
    const after = points.find((point) => point.blockIdx === start + 16);
    if (before && after && before.pageCount > 0 && before.pageCount === after.pageCount) return start + 1;
  }
  throw new Error("fixture did not produce a same-page sparse checkpoint interval");
}

function paintProjection(result: ReturnType<typeof layoutDocument>): string {
  return JSON.stringify(result.pages, (key, value) => (key === "src" || key === "tbl" ? undefined : value));
}

function pageProjection(result: ReturnType<typeof layoutDocument>, pageIndex: number): string {
  return JSON.stringify(result.pages[pageIndex], (key, value) =>
    key === "src" || key === "tbl" ? undefined : value,
  );
}

describe("page-model resume points", () => {
  it("keeps the complete model below the virtualization threshold", () => {
    const body = Array.from({ length: 20 }, (_, index) =>
      p(`page-${index} alpha bravo charlie delta`) +
      (index < 19 ? `<w:p><w:r><w:br w:type="page"/></w:r></w:p>` : ""),
    ).join("");
    const doc = DocxDocument.load(
      makeDocx({ "word/document.xml": wrapDocument(body + section) }),
    );
    const expected = layoutDocument(doc, { measurer });
    const result = layoutDocument(doc, { measurer, windowModel: true });

    expect(result.totalPages).toBe(20);
    expect(result._window).toBeUndefined();
    expect(paintProjection(result)).toBe(paintProjection(expected));
  });

  it("reproduces every page-top point and detects altered resume state", () => {
    const body = Array.from({ length: 24 }, (_, index) =>
      p(`page-${index} alpha bravo charlie delta`) +
      (index < 23 ? `<w:p><w:r><w:br w:type="page"/></w:r></w:p>` : ""),
    ).join("");
    const doc = DocxDocument.load(makeDocx({ "word/document.xml": wrapDocument(body + section) }));
    const result = layoutDocument(doc, { measurer, windowModel: true });
    expect(result.totalPages).toBeGreaterThan(20);
    expect(result._window).toBeDefined();

    result._window!.materialize(result.pages.keys());
    const expected = result.pages.map((_, index) => pageProjection(result, index));
    const points = (result._incr as {
      points: Array<{
        pageCount: number;
        pageItemCount: number;
        state: { page: { colXs: number[] } };
      }>;
    }).points.filter((point) => point.pageItemCount === 0);
    const pageIndexes = [...new Set(points.map((point) => point.pageCount))];
    expect(pageIndexes.length).toBeGreaterThan(2);

    result._window!.releaseExcept([0]);
    for (const pageIndex of pageIndexes.slice(1)) {
      result._window!.materialize([pageIndex]);
      expect(pageProjection(result, pageIndex)).toBe(expected[pageIndex]);
      result._window!.releaseExcept([0]);
    }

    const changedPage = pageIndexes.find((pageIndex) => pageIndex > 1)!;
    const changed = points.filter((point) => point.pageCount === changedPage).at(-1)!;
    const originalX = changed.state.page.colXs[0];
    changed.state.page.colXs[0] += 8;
    const preceding = points
      .filter((point) => point.pageCount < changed.pageCount)
      .at(-1)!.pageCount;
    result._window!.releaseExcept([preceding]);
    result._window!.materialize([changed.pageCount]);
    expect(pageProjection(result, changed.pageCount)).not.toBe(expected[changed.pageCount]);
    changed.state.page.colXs[0] = originalX;
  });

  it("reproduces the next page from every intra-page point", () => {
    const shortSection =
      `<w:sectPr><w:pgSz w:w="7200" w:h="3000"/>` +
      `<w:pgMar w:top="360" w:right="360" w:bottom="360" w:left="360"/></w:sectPr>`;
    const body = Array.from(
      { length: 320 },
      (_, index) => p(`block-${index} alpha bravo charlie delta`),
    ).join("");
    const doc = DocxDocument.load(
      makeDocx({ "word/document.xml": wrapDocument(body + shortSection) }),
    );
    const expected = layoutDocument(doc, { measurer });
    const result = layoutDocument(doc, { measurer, windowModel: true });
    expect(result.totalPages).toBeGreaterThan(20);
    expect(result._window).toBeDefined();

    const data = result._incr as {
      points: Array<{
        pageCount: number;
        pageItemCount: number;
        state: { y: number };
      }>;
    };
    const allPoints = data.points;
    const intraPagePoints = allPoints.filter(
      (point) => point.pageItemCount > 0 && point.pageCount + 1 < result.totalPages,
    );
    expect(intraPagePoints.length).toBeGreaterThan(5);

    for (const point of intraPagePoints) {
      data.points = allPoints.slice(0, allPoints.indexOf(point) + 1);
      result._window!.releaseExcept([]);
      result._window!.materialize([point.pageCount + 1]);
      expect(pageProjection(result, point.pageCount + 1)).toBe(
        pageProjection(expected, point.pageCount + 1),
      );
    }

    const changed = intraPagePoints[Math.floor(intraPagePoints.length / 2)];
    const originalY = changed.state.y;
    changed.state.y += 200;
    data.points = allPoints.slice(0, allPoints.indexOf(changed) + 1);
    result._window!.releaseExcept([]);
    result._window!.materialize([changed.pageCount + 1]);
    expect(pageProjection(result, changed.pageCount + 1)).not.toBe(
      pageProjection(expected, changed.pageCount + 1),
    );
    changed.state.y = originalY;
    data.points = allPoints;
  });

  it("keeps offscreen pages rematerializable after an incremental edit", () => {
    const body = Array.from({ length: 24 }, (_, index) =>
      p(`page-${index} alpha bravo charlie delta`) +
      (index < 23 ? `<w:p><w:r><w:br w:type="page"/></w:r></w:p>` : ""),
    ).join("");
    const doc = DocxDocument.load(makeDocx({ "word/document.xml": wrapDocument(body + section) }));
    const first = layoutDocument(doc, { measurer, windowModel: true });
    const targetPage = 15;
    first._window!.materialize([targetPage]);

    const changed = editParagraph(doc, targetPage * 2, "x");
    const incremental = layoutDocument(doc, {
      measurer,
      windowModel: true,
      prev: first,
      dirtyHint: changed.block.src,
      dirtySource: changed.source,
    });
    expect(incremental._incremental).toBe(true);
    expect(incremental._window).toBeDefined();

    incremental._window!.materialize(incremental.pages.keys());
    const full = layoutDocument(doc, { measurer });
    expect(paintProjection(incremental)).toBe(paintProjection(full));

    incremental._window!.releaseExcept([targetPage]);
    incremental._window!.materialize([20]);
    expect(pageProjection(incremental, 20)).toBe(pageProjection(full, 20));
  });
});

describe("incremental same-page block checkpoints", () => {
  it("matches a full layout, reuses every unchanged page, and remains local on repeated edits", () => {
    const doc = denseDoc();
    const first = layoutDocument(doc, { measurer });
    expect(first.totalPages).toBeGreaterThan(2);
    const target = samePageSparseTarget(first);

    const changed = editParagraph(doc, target, "x");
    const incremental = layoutDocument(doc, {
      measurer,
      prev: first,
      dirtyHint: changed.block.src,
      dirtySource: changed.source,
    });
    const full = layoutDocument(doc, { measurer });
    expect(incremental._incremental).toBe(true);
    expect(paintProjection(incremental)).toBe(paintProjection(full));
    // A checkpoint can sit at the bottom of the preceding page when its next
    // paragraph moves to a fresh page. The engine relays both pages; the DOM
    // renderer structurally adopts the unchanged leading page and rebuilds the
    // dirty page, while every later layout page retains identity here.
    expect(incremental.pages.filter((page, i) => page === first.pages[i]).length).toBeGreaterThanOrEqual(
      first.totalPages - 2,
    );
    expect(__incrStats.blocksLaid).toBeLessThanOrEqual(16);
    expect(__incrStats.resumeBlock).toBeLessThanOrEqual(__incrStats.firstDirty);
    expect(__incrStats.convergedBlock).toBeGreaterThan(__incrStats.firstDirty);

    const changedAgain = editParagraph(doc, target, "y");
    const repeated = layoutDocument(doc, {
      measurer,
      prev: incremental,
      dirtyHint: changedAgain.block.src,
      dirtySource: changedAgain.source,
    });
    const repeatedFull = layoutDocument(doc, { measurer });
    expect(repeated._incremental).toBe(true);
    expect(paintProjection(repeated)).toBe(paintProjection(repeatedFull));
    expect(__incrStats.blocksLaid).toBeLessThanOrEqual(16);
  });

  it("falls through safely when wrapping changes pagination", () => {
    const doc = denseDoc();
    const first = layoutDocument(doc, { measurer });
    const changed = editParagraph(doc, 37, ` ${"wrapping content ".repeat(90)}`);
    const attempted = layoutDocument(doc, {
      measurer,
      prev: first,
      dirtyHint: changed.block.src,
      dirtySource: changed.source,
    });
    const full = layoutDocument(doc, { measurer });
    expect(attempted.totalPages).toBeGreaterThan(first.totalPages);
    expect(paintProjection(attempted)).toBe(paintProjection(full));
  });

  it("converges after the final numbered paragraph exits its list", () => {
    const numbering = `<?xml version="1.0"?><w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`;
    const numbered = (text: string) =>
      `<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`;
    const trailing = Array.from({ length: 94 }, (_, index) => p(`tail-${index} alpha bravo charlie delta`)).join("");
    const doc = DocxDocument.load(makeDocx({
      "word/document.xml": wrapDocument(numbered("first") + numbered("second") + trailing + section),
      "word/numbering.xml": numbering,
    }));
    const first = layoutDocument(doc, { measurer });
    const block = doc.sections[0].blocks[1] as Paragraph;
    const source = block.src!;
    const pPr = source.children.find((child) => localName(child.name) === "pPr")!;
    pPr.children = pPr.children.filter((child) => localName(child.name) !== "numPr");
    const reparsed = doc.reparseBodyParagraph(source);
    expect(reparsed).not.toBeNull();
    invalidateParagraphSignature(source);

    const incremental = layoutDocument(doc, {
      measurer,
      prev: first,
      dirtyHint: source,
      dirtySource: (reparsed!.children[0] as Run).content[0].srcT,
    });
    const full = layoutDocument(doc, { measurer });
    expect(incremental._incremental).toBe(true);
    expect(paintProjection(incremental)).toBe(paintProjection(full));
    expect(__incrStats.convergedBlock).toBeGreaterThan(1);
    expect(__incrStats.blocksLaid).toBeLessThanOrEqual(16);
  });

  it("converges after toggling a bullet when later bullets share the definition", () => {
    const numbering = `<?xml version="1.0"?><w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`;
    const bullet = (text: string) =>
      `<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`;
    const body = Array.from({ length: 96 }, (_, index) =>
      index > 16 && index % 10 === 0 ? bullet(`bullet-${index}`) : p(`block-${index} alpha bravo charlie delta`),
    ).join("");
    const doc = DocxDocument.load(makeDocx({
      "word/document.xml": wrapDocument(body + section),
      "word/numbering.xml": numbering,
    }));
    const first = layoutDocument(doc, { measurer });
    const block = doc.sections[0].blocks[1] as Paragraph;
    const text = (block.children[0] as Run).content[0] as TextContent;
    expect(setListType(doc, [text.srcT!], "bullet")).toBe(true);

    const incremental = layoutDocument(doc, {
      measurer,
      prev: first,
      dirtyHint: block.src,
      dirtySource: text.srcT,
    });
    const full = layoutDocument(doc, { measurer });
    expect(incremental._incremental).toBe(true);
    expect(paintProjection(incremental)).toBe(paintProjection(full));
    expect(__incrStats.blocksLaid).toBeLessThanOrEqual(16);
  });

  it("updates equal-width PAGEREFs without discarding a converged incremental layout", () => {
    const pageRef =
      `<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r>` +
      `<w:r><w:instrText xml:space="preserve"> PAGEREF Target \\h </w:instrText></w:r>` +
      `<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>9</w:t></w:r>` +
      `<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>`;
    const body = Array.from({ length: 96 }, (_, i) =>
      i === 36
        ? `<w:p><w:bookmarkStart w:id="1" w:name="Target"/><w:r><w:t>block-${i} alpha bravo charlie delta</w:t></w:r><w:bookmarkEnd w:id="1"/></w:p>`
        : p(`block-${i} alpha bravo charlie delta`),
    ).join("");
    const doc = DocxDocument.load(makeDocx({ "word/document.xml": wrapDocument(pageRef + body + section) }));
    const first = layoutDocument(doc, { measurer });
    const incrementalData = first._incr as {
      bookmarks: Map<string, string>;
      pages: Array<{ items: Array<{ kind: string; text?: string; pageRef?: string }> }>;
    };
    const actualPage = incrementalData.bookmarks.get("Target");
    expect(actualPage).toMatch(/^\d$/);
    const stalePage = actualPage === "8" ? "7" : "8";
    incrementalData.bookmarks.set("Target", stalePage);
    let staleRefs = 0;
    for (const page of incrementalData.pages) {
      for (const item of page.items) {
        if (item.kind !== "text" || item.pageRef !== "Target") continue;
        item.text = stalePage;
        staleRefs++;
      }
    }
    expect(staleRefs).toBeGreaterThan(0);

    const changed = editParagraph(doc, 37, "x");
    const incremental = layoutDocument(doc, {
      measurer,
      prev: first,
      dirtyHint: changed.block.src,
      dirtySource: changed.source,
    });
    const full = layoutDocument(doc, { measurer });
    expect(incremental._incremental).toBe(true);
    expect(__incrStats.fallbackReason).toBe("");
    expect(paintProjection(incremental)).toBe(paintProjection(full));
  });

  it("converges after a one-to-two paragraph split and keeps shifted checkpoints reusable", () => {
    const bodyXml = Array.from({ length: 96 }, (_, i) =>
      (i === 48 ? `<w:p><w:r><w:br w:type="page"/></w:r></w:p>` : "") +
      p(`block-${i} alpha bravo charlie delta`),
    ).join("");
    const doc = DocxDocument.load(makeDocx({ "word/document.xml": wrapDocument(bodyXml + section) }));
    const first = layoutDocument(doc, { measurer });
    const target = 37;
    const block = doc.sections[0].blocks[target] as Paragraph;
    const run = block.children[0] as Run;
    const text = run.content[0] as TextContent;
    const source = block.src!;
    const sourceText = text.srcT!;
    const splitAt = Math.floor(sourceText.text.length / 2);
    const afterText: XmlElement = {
      name: sourceText.name,
      attrs: { ...sourceText.attrs, "xml:space": "preserve" },
      children: [],
      text: sourceText.text.slice(splitAt),
    };
    const afterSource: XmlElement = {
      name: source.name,
      attrs: {},
      text: "",
      children: [{ name: run.src!.name, attrs: {}, text: "", children: [afterText] }],
    };
    sourceText.text = sourceText.text.slice(0, splitAt);
    const body = doc.docRoot.children.find((element) => localName(element.name) === "body")!;
    body.children.splice(body.children.indexOf(source) + 1, 0, afterSource);
    const version = doc.modelVersion;
    expect(doc.reparseDirectBodyParagraphSplit(source, afterSource)).not.toBeNull();
    expect(doc.modelVersion).toBe(version);
    invalidateParagraphSignature(source);
    invalidateParagraphSignature(afterSource);

    const incremental = layoutDocument(doc, {
      measurer,
      prev: first,
      dirtyHint: afterSource,
      dirtySource: afterText,
    });
    const full = layoutDocument(doc, { measurer });
    expect(incremental._incremental).toBe(true);
    expect(paintProjection(incremental)).toBe(paintProjection(full));
    expect(__incrStats.hintFastPath).toBe(true);
    expect(__incrStats.blocksHashed).toBeLessThanOrEqual(4);
    expect(__incrStats.firstDirty).toBe(target);
    expect(__incrStats.convergedBlock).toBeGreaterThan(target + 1);
    expect(__incrStats.blocksLaid).toBeLessThanOrEqual(20);

    const later = editParagraph(doc, target + 5, "z");
    const repeated = layoutDocument(doc, {
      measurer,
      prev: incremental,
      dirtyHint: later.block.src,
      dirtySource: later.source,
    });
    expect(repeated._incremental).toBe(true);
    expect(paintProjection(repeated)).toBe(paintProjection(layoutDocument(doc, { measurer })));
    expect(__incrStats.blocksLaid).toBeLessThanOrEqual(20);
  });

  it("reparses a bookmarked paragraph in place and recaptures its refBookmarks range", () => {
    // The dead-editor regression: TOC targets bookmark every heading, and a
    // reparse rejection here sent the first keystroke in any heading through
    // doc.refresh() + a full inert relayout on long documents.
    const body =
      `<w:p><w:bookmarkStart w:id="1" w:name="Target"/><w:r><w:t>heading text</w:t></w:r><w:bookmarkEnd w:id="1"/></w:p>` +
      Array.from({ length: 8 }, (_, i) => p(`block-${i} alpha bravo`)).join("");
    const doc = DocxDocument.load(makeDocx({ "word/document.xml": wrapDocument(body + section) }));
    const block = doc.sections[0].blocks[0] as Paragraph;
    const text = (block.children[0] as Run).content[0] as TextContent;
    const version = doc.modelVersion;
    text.srcT!.text = "heading textZ";

    const reparsed = doc.reparseBodyParagraph(block.src!);
    expect(reparsed).not.toBeNull();
    expect(doc.modelVersion).toBe(version);
    // REF fields read the bookmark range through doc.refBookmarks; the
    // recapture must hold the NEW runs (stale captures would render old text).
    const runs = doc.refBookmarks.get("Target")!;
    const captured = runs
      .flatMap((run) => run.content)
      .filter((content) => content.kind === "text")
      .map((content) => (content as TextContent).text)
      .join("");
    expect(captured).toBe("heading textZ");
  });

  it("rejects an in-place reparse when a bookmark range crosses the paragraph", () => {
    const body =
      `<w:p><w:bookmarkStart w:id="1" w:name="Spanning"/><w:r><w:t>range opens here</w:t></w:r></w:p>` +
      p("covered middle paragraph") +
      `<w:p><w:r><w:t>range closes here</w:t></w:r><w:bookmarkEnd w:id="1"/></w:p>`;
    const doc = DocxDocument.load(makeDocx({ "word/document.xml": wrapDocument(body + section) }));
    const blocks = doc.sections[0].blocks as Paragraph[];
    // Unbalanced markers inside the paragraph.
    expect(doc.reparseBodyParagraph(blocks[0].src!)).toBeNull();
    expect(doc.reparseBodyParagraph(blocks[2].src!)).toBeNull();
    // No markers inside, but the spanning range captured this paragraph's runs.
    expect(doc.reparseBodyParagraph(blocks[1].src!)).toBeNull();
  });

  it("reparses a paragraph split inside a table cell without refreshing the document", () => {
    const tableXml =
      `<w:tbl><w:tblPr><w:tblW w:w="4000" w:type="dxa"/></w:tblPr>` +
      `<w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc>` +
      `<w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>` +
      p("alpha bravo charlie delta") +
      `</w:tc></w:tr></w:tbl>`;
    const trailing = Array.from({ length: 80 }, (_, i) => p(`tail-${i} echo foxtrot golf hotel`)).join("");
    const doc = DocxDocument.load(makeDocx({ "word/document.xml": wrapDocument(tableXml + trailing + section) }));
    const first = layoutDocument(doc, { measurer });
    const table = doc.sections[0].blocks[0];
    expect(table.type).toBe("table");
    if (table.type !== "table") throw new Error("expected table");
    const cellBlocks = table.rows[0].cells[0].blocks;
    const before = cellBlocks[0] as Paragraph;
    const beforeSource = before.src!;
    const beforeRun = before.children[0] as Run;
    const beforeText = beforeRun.content[0] as TextContent;
    const beforeT = beforeText.srcT!;
    const splitAt = 12;
    const afterT: XmlElement = {
      name: beforeT.name,
      attrs: { ...beforeT.attrs, "xml:space": "preserve" },
      children: [],
      text: beforeT.text.slice(splitAt),
    };
    const afterSource: XmlElement = {
      name: beforeSource.name,
      attrs: {},
      text: "",
      children: [{ name: beforeRun.src!.name, attrs: {}, text: "", children: [afterT] }],
    };
    beforeT.text = beforeT.text.slice(0, splitAt);
    const cellSource = doc.findParentOf(beforeSource)!;
    cellSource.children.splice(cellSource.children.indexOf(beforeSource) + 1, 0, afterSource);
    const version = doc.modelVersion;

    const reparsed = doc.reparseDirectBodyParagraphSplit(beforeSource, afterSource);
    expect(reparsed).not.toBeNull();
    expect(doc.modelVersion).toBe(version);
    expect(cellBlocks).toHaveLength(2);
    expect(cellBlocks[0].src).toBe(beforeSource);
    expect(cellBlocks[1].src).toBe(afterSource);
    invalidateParagraphSignature(beforeSource);
    invalidateParagraphSignature(afterSource);

    const incremental = layoutDocument(doc, {
      measurer,
      prev: first,
      dirtyHint: table.src,
      dirtySource: afterT,
    });
    expect(incremental._incremental).toBe(true);
    expect(paintProjection(incremental)).toBe(paintProjection(layoutDocument(doc, { measurer })));
  });
});
