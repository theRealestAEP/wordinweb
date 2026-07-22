import { describe, expect, it } from "vitest";
import { DocxDocument } from "../src/docx.js";
import { __incrStats, layoutDocument } from "../src/layout/engine.js";
import { invalidateParagraphSignature } from "../src/layout/inline.js";
import { ApproxMeasurer } from "../src/layout/measure.js";
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
