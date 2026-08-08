import { describe, expect, it } from "vitest";
import { DocxDocument } from "../src/docx.js";
import { documentTextStatistics } from "../src/word-count.js";
import { updateFields } from "../src/edit/update-fields.js";
import { XmlElement, localName, serializeXml } from "../src/xml.js";
import { makeDocx, wrapDocument, p } from "./helpers.js";

function loadDoc(body: string) {
  return DocxDocument.load(makeDocx({ "word/document.xml": wrapDocument(body) }));
}

describe("documentTextStatistics", () => {
  it("counts words, characters, and non-empty paragraphs, tables included", () => {
    const body =
      p("Hello brave new world") + // 4 words, 18 chars (no spaces)
      p("") + // empty: not a paragraph for the dialog
      `<w:tbl><w:tr><w:tc>${p("cell one")}</w:tc><w:tc>${p("two")}</w:tc></w:tr></w:tbl>` +
      p("End.");
    const stats = documentTextStatistics(loadDoc(body));
    expect(stats.words).toBe(4 + 2 + 1 + 1);
    expect(stats.paragraphs).toBe(4);
    expect(stats.characters).toBe("Hellobravenewworld".length + "cellone".length + "two".length + "End.".length);
    expect(stats.charactersWithSpaces).toBe(
      "Hello brave new world".length + "cell one".length + "two".length + "End.".length,
    );
  });

  it("tabs and soft breaks separate words without counting as no-space characters", () => {
    const body = `<w:p><w:r><w:t>one</w:t><w:tab/><w:t>two</w:t><w:br/><w:t>three</w:t></w:r></w:p>`;
    const stats = documentTextStatistics(loadDoc(body));
    expect(stats.words).toBe(3);
    expect(stats.characters).toBe("onetwothree".length);
    expect(stats.charactersWithSpaces).toBe("onetwothree".length + 2);
  });

  it("a field contributes its cached result text", () => {
    const body = `<w:p><w:r><w:t xml:space="preserve">Written by </w:t></w:r><w:fldSimple w:instr=" AUTHOR "><w:r><w:t>Ada Lovelace</w:t></w:r></w:fldSimple></w:p>`;
    const stats = documentTextStatistics(loadDoc(body));
    expect(stats.words).toBe(4);
  });
});

describe("NUMWORDS / NUMCHARS field recompute", () => {
  it("updateFields writes the shared-rule statistics into the field caches", () => {
    const body =
      p("Exactly four words here") +
      `<w:p><w:r><w:t xml:space="preserve">Words: </w:t></w:r><w:fldSimple w:instr=" NUMWORDS \\* MERGEFORMAT "><w:r><w:t>0</w:t></w:r></w:fldSimple></w:p>` +
      `<w:p><w:r><w:t xml:space="preserve">Chars: </w:t></w:r><w:fldSimple w:instr=" NUMCHARS "><w:r><w:t>0</w:t></w:r></w:fldSimple></w:p>`;
    const doc = loadDoc(body);
    // The pass computes from the PRE-pass state (a field's own cache counts as
    // text, exactly as Word counts a field's result).
    const stats = documentTextStatistics(doc);
    expect(updateFields(doc)).toBe(true);
    const xml = serializeXml(doc.docRoot, true);
    const numwordsCache = fieldCache(doc, "NUMWORDS");
    const numcharsCache = fieldCache(doc, "NUMCHARS");
    expect(numwordsCache).toBe(String(stats.words));
    expect(numcharsCache).toBe(String(stats.characters));
    expect(xml).toContain(numwordsCache);
  });
});

function fieldCache(doc: DocxDocument, keyword: string): string {
  let out = "";
  const walk = (e: XmlElement): void => {
    if (localName(e.name) === "fldSimple" && (e.attrs["w:instr"] ?? "").includes(keyword)) {
      const collect = (n: XmlElement): void => {
        if (localName(n.name) === "t") out += n.text;
        for (const c of n.children) collect(c);
      };
      collect(e);
      return;
    }
    for (const c of e.children) walk(c);
  };
  walk(doc.docRoot);
  return out;
}
