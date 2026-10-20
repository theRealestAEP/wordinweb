// Word for MS-DOS documents (oldword/dos.ts). The fixtures are public domain
// US government files from the GovDocs1 corpus (digitalcorpora.org,
// govdocs1/zipfiles/NNN.zip#NNN/NNNNNN.doc), saved by Word 4 and Word 5.
import { readFileSync } from "node:fs";
import { zipSync } from "fflate";
import { describe, expect, it } from "vitest";
import { DocxDocument } from "../src/docx.js";
import { compareDocuments } from "../src/edit/compare/index.js";
import { collectRevisions } from "../src/edit/suggest.js";
import { isDosWordFile, readDosWord } from "../src/oldword/dos.js";
import { Block, Inline, OldDocument } from "../src/oldword/model.js";
import { oldDocumentFiles } from "../src/oldword/package.js";

const fixture = (name: string): Uint8Array => new Uint8Array(readFileSync(new URL(`./fixtures/oldword-dos/${name}.doc`, import.meta.url)));

/** The text of blocks, a paragraph a line, with its tabs, breaks and notes as characters. */
function text(blocks: Block[]): string {
  const inline = (item: Inline): string => {
    switch (item.kind) {
      case "text": return item.text;
      case "tab": return "\t";
      case "noBreakHyphen": return "‑";
      case "footnote": return `[${item.id}]`;
      case "field": return `{${item.code}}`;
      default: return "";
    }
  };
  return blocks.map((block) => (block.type === "paragraph" ? block.runs.flatMap((run) => run.content.map(inline)).join("") : "")).join("\n");
}

function bodyText(doc: OldDocument): string {
  return doc.sections.map((section) => text(section.blocks)).join("\n");
}

describe("Word for MS-DOS", () => {
  it("tells Word for MS-DOS from Windows Write", () => {
    expect(isDosWordFile(fixture("govdocs-804182"))).toBe(true);
    // Write: the same signature, its page count at 0x60 and no style sheet name.
    const write = new Uint8Array(128);
    write.set([0x31, 0xbe, 0, 0, 0, 0xab]);
    write[0x60] = 4;
    expect(isDosWordFile(write)).toBe(false);
  });

  it("reads the text and the section's page", () => {
    const doc = readDosWord(fixture("govdocs-804182"));
    expect(doc.sections).toHaveLength(1);
    const [section] = doc.sections;
    expect(section.props).toMatchObject({ xaPage: 12240, yaPage: 15840, dyaTop: 1440, dyaBottom: 1440, dxaLeft: 1440, dxaRight: 1440 });
    const body = bodyText(doc);
    expect(body).toContain("ISO/IEC JTC1/SC18/WG8 1031\nTitle:\nLiaison to WG1, WG3 on WG8 Model Response to WG8 N963 = WG3 N1416");
    expect(body).toMatch(/forward a copy to you\.$/);
  });

  it("reads running heads for odd and even pages, and how sections start", () => {
    const doc = readDosWord(fixture("govdocs-262493"));
    // A cover page, a section of running heads only (on the same page), the report on a new page.
    expect(doc.sections.map((s) => s.props.bkc)).toEqual([undefined, 0, undefined]);
    const heads = doc.sections[1].headers;
    expect(text(heads.find((h) => h.kind === "header" && h.type === "default")!.blocks)).toBe("ISO/IEC JTC1/SC18/WG8 N849\nPage {PAGE}\n");
    expect(text(heads.find((h) => h.kind === "header" && h.type === "even")!.blocks)).toBe("ISO/IEC JTC1/SC18/WG8 N849\nPage {PAGE}\n");
    expect(doc.evenAndOddHeaders).toBe(true);
    expect(bodyText(doc)).not.toContain("Page {PAGE}");
  });

  it("decodes the characters a PostScript printer printed", () => {
    // Word 4's driver: Adobe's standard encoding (0xAA, 0xBA quotes, 0xD0 em dash).
    expect(bodyText(readDosWord(fixture("govdocs-199583")))).toMatch(/“User Requirements for SGML”.*\n?.*systems—that is DIS 9541/s);
    // Word 5's driver: 0xD3 an en dash.
    expect(bodyText(readDosWord(fixture("govdocs-974019")))).toContain("N1106–1155");
  });

  it("reads footnotes and Word's nonbreaking hyphen", () => {
    const doc = readDosWord(fixture("govdocs-191777"));
    expect(doc.footnotes).toHaveLength(1);
    expect(text(doc.footnotes[0])).toContain("All references to the National Electrical Code in this document pertain to the 1993 edition.");
    const body = bodyText(doc);
    expect(body).toContain("[1]");
    expect(body).toContain("Solid‑State Arresters");
    expect(body).not.toContain("─");
  });

  it("makes a table of side-by-side paragraphs", () => {
    const doc = readDosWord(fixture("govdocs-921227"));
    const tables = doc.sections.flatMap((s) => s.blocks).filter((b) => b.type === "table");
    const first = tables[0];
    if (first?.type !== "table") throw new Error("no table");
    const cells = first.rows.map((row) => row.cells.map((cell) => text(cell)));
    expect(cells[0]).toEqual(["Volcano", "SO2-start ", "number of measurements", "CO2 - start", "number of measurements"]);
    expect(cells[1]).toEqual(["Redoubt", "March, 1990", "   64", "June, 1990", "   7"]);
    // The cells' edges are the paragraphs' left indents.
    expect(first.rows[1].props.cells.map((c) => c.right)).toEqual([1296, 2880, 4464, 6048, expect.any(Number)]);
  });

  it("converts to a .docx that compares equal to itself", () => {
    for (const name of ["govdocs-804182", "govdocs-262493", "govdocs-921227", "govdocs-199583", "govdocs-974019", "govdocs-191777"]) {
      const docx = zipSync(oldDocumentFiles(readDosWord(fixture(name))));
      const result = compareDocuments(DocxDocument.load(docx), DocxDocument.load(docx), { author: "test", date: "2026-01-01T00:00:00Z" });
      expect(collectRevisions(result), name).toHaveLength(0);
    }
  });
});
