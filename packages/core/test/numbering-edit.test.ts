import { unzipSync, strFromU8 } from "fflate";
import { describe, expect, it } from "vitest";
import { DocxDocument } from "../src/docx.js";
import {
  continueNumberingAt,
  detachNumbering,
  listInstanceAt,
  restartNumberingAt,
  setNumberingLevel,
  setNumberingRestart,
} from "../src/edit/numbering.js";
import { serializeXml } from "../src/xml.js";
import { makeDocx, wrapDocument } from "./helpers.js";

/**
 * One abstract multilevel definition, shared by two instances — which is the
 * arrangement that makes the abstract/instance distinction observable: an edit
 * to the definition must reach both, and a restart must reach only one.
 */
const NUMBERING_XML =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
  `<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
  `<w:abstractNum w:abstractNumId="0"><w:nsid w:val="0FC066EB"/>` +
  `<w:multiLevelType w:val="multilevel"/><w:tmpl w:val="C0200884"/>` +
  `<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/>` +
  `<w:lvlText w:val="%1."/><w:lvlJc w:val="left"/>` +
  `<w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>` +
  `<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/>` +
  `<w:lvlText w:val="%2."/><w:lvlJc w:val="left"/>` +
  `<w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr></w:lvl>` +
  `</w:abstractNum>` +
  `<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>` +
  `<w:num w:numId="2"><w:abstractNumId w:val="0"/></w:num>` +
  `</w:numbering>`;

const RELS_XML =
  `<?xml version="1.0"?>` +
  `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
  `<Relationship Id="rId10" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>` +
  `</Relationships>`;

function item(text: string, numId: number, ilvl = 0): string {
  return (
    `<w:p><w:pPr><w:numPr><w:ilvl w:val="${ilvl}"/><w:numId w:val="${numId}"/></w:numPr></w:pPr>` +
    `<w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`
  );
}

function load(body: string): DocxDocument {
  return DocxDocument.load(
    makeDocx({
      "word/document.xml": wrapDocument(body),
      "word/numbering.xml": NUMBERING_XML,
      "word/_rels/document.xml.rels": RELS_XML,
    }),
  );
}

function numberingXml(doc: DocxDocument): string {
  return serializeXml(doc.numberingTree()!);
}

/** The nth paragraph's element, which is how a caret addresses it. */
function paragraphAt(doc: DocxDocument, index: number) {
  return doc.docRoot.children[0].children[index];
}

// ---------------------------------------------------------------------------

describe("setNumberingLevel", () => {
  it("patches a level and keeps CT_Lvl's child order", () => {
    const doc = load(item("one", 1));
    expect(
      setNumberingLevel(doc, 1, 0, {
        format: "upperRoman",
        text: "%1)",
        start: 3,
        alignment: "right",
        indentLeftPt: 54,
        hangingPt: 18,
      }),
    ).toBe(true);
    const lvl = numberingXml(doc).match(/<w:lvl w:ilvl="0">.*?<\/w:lvl>/s)![0];
    expect([...lvl.matchAll(/<w:(start|numFmt|lvlText|lvlJc|pPr)[ />]/g)].map((m) => m[1])).toEqual([
      "start", "numFmt", "lvlText", "lvlJc", "pPr",
    ]);
    expect(lvl).toContain(`<w:numFmt w:val="upperRoman"/>`);
    expect(lvl).toContain(`<w:lvlText w:val="%1)"/>`);
    expect(lvl).toContain(`<w:start w:val="3"/>`);
    expect(lvl).toContain(`<w:lvlJc w:val="right"/>`);
    expect(lvl).toContain(`<w:ind w:left="1080" w:hanging="360"/>`);
  });

  it("reaches every instance built on the same abstract definition", () => {
    const doc = load(item("a", 1) + item("b", 2));
    setNumberingLevel(doc, 1, 0, { format: "upperLetter" });
    expect(doc.numberingLevel(1, 0)?.format).toBe("upperLetter");
    expect(doc.numberingLevel(2, 0)?.format).toBe("upperLetter");
  });

  it("re-resolves the label the layout paints", () => {
    const doc = load(item("a", 1));
    expect(doc.numberingLevel(1, 0)?.text).toBe("%1.");
    const before = doc.layoutGlobalSig();
    setNumberingLevel(doc, 1, 1, { text: "%1.%2." });
    expect(doc.numberingLevel(1, 1)?.text).toBe("%1.%2.");
    expect(doc.layoutGlobalSig()).not.toBe(before);
  });

  it("rejects a level or an instance the part does not declare", () => {
    const doc = load(item("a", 1));
    expect(setNumberingLevel(doc, 1, 7, { format: "decimal" })).toBe(false);
    expect(setNumberingLevel(doc, 99, 0, { format: "decimal" })).toBe(false);
  });
});

describe("detachNumbering", () => {
  it("gives one instance a private copy so the other stops following it", () => {
    const doc = load(item("a", 1) + item("b", 2));
    const fresh = detachNumbering(doc, 2);
    expect(fresh).toBe(1);
    setNumberingLevel(doc, 2, 0, { format: "upperRoman" });
    expect(doc.numberingLevel(2, 0)?.format).toBe("upperRoman");
    expect(doc.numberingLevel(1, 0)?.format).toBe("decimal");
  });

  it("drops the gallery identity Word uses to re-link definitions", () => {
    const doc = load(item("a", 1));
    detachNumbering(doc, 1);
    const copy = numberingXml(doc).match(/<w:abstractNum w:abstractNumId="1">.*?<\/w:abstractNum>/s)![0];
    expect(copy).not.toContain("w:nsid");
    expect(copy).not.toContain("w:tmpl");
    expect(copy).toContain(`<w:multiLevelType w:val="multilevel"/>`);
  });

  it("keeps abstractNum entries ahead of num entries", () => {
    const doc = load(item("a", 1));
    detachNumbering(doc, 1);
    const xml = numberingXml(doc);
    expect(xml.lastIndexOf("<w:abstractNum ")).toBeLessThan(xml.indexOf("<w:num "));
  });
});

describe("setNumberingRestart on an instance", () => {
  it("writes w:lvlOverride / w:startOverride the way the corpus does", () => {
    const doc = load(item("a", 1));
    expect(setNumberingRestart(doc, 1, 0, 5)).toBe(true);
    expect(numberingXml(doc)).toContain(
      `<w:num w:numId="1"><w:abstractNumId w:val="0"/>` +
        `<w:lvlOverride w:ilvl="0"><w:startOverride w:val="5"/></w:lvlOverride></w:num>`,
    );
  });

  it("removes the whole override when the restart is dropped", () => {
    const doc = load(item("a", 1));
    setNumberingRestart(doc, 1, 0, 5);
    expect(setNumberingRestart(doc, 1, 0, null)).toBe(true);
    expect(numberingXml(doc)).toContain(`<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>`);
    // Dropping a restart that is not there changes nothing.
    expect(setNumberingRestart(doc, 1, 0, null)).toBe(false);
  });
});

describe("restart / continue at a paragraph", () => {
  it("splits the list into a fresh instance from the caret down", () => {
    const doc = load(item("one", 1) + item("two", 1) + item("three", 1));
    expect(restartNumberingAt(doc, paragraphAt(doc, 1), 1)).toBe(true);

    // The paragraph BEFORE the caret keeps the original instance; the caret's
    // paragraph and the one after it move to the new one. A startOverride on
    // the original would have restarted the list at "one" instead.
    expect(listInstanceAt(doc, paragraphAt(doc, 0))!.numId).toBe(1);
    const fresh = listInstanceAt(doc, paragraphAt(doc, 1))!.numId;
    expect(fresh).toBe(3);
    expect(listInstanceAt(doc, paragraphAt(doc, 2))!.numId).toBe(3);

    expect(numberingXml(doc)).toContain(
      `<w:num w:numId="3"><w:abstractNumId w:val="0"/>` +
        `<w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride></w:num>`,
    );
  });

  it("stops at the first paragraph that leaves the list", () => {
    const doc = load(item("one", 1) + `<w:p><w:r><w:t>prose</w:t></w:r></w:p>` + item("two", 1));
    restartNumberingAt(doc, paragraphAt(doc, 0), 1);
    // The list item after the prose belongs to a different run, so it keeps
    // the instance it had.
    expect(listInstanceAt(doc, paragraphAt(doc, 2))!.numId).toBe(1);
  });

  it("makes a restarted list continue the preceding one again", () => {
    const doc = load(item("one", 1) + item("two", 1) + item("three", 1));
    restartNumberingAt(doc, paragraphAt(doc, 1), 1);
    expect(continueNumberingAt(doc, paragraphAt(doc, 1))).toBe(true);
    expect(listInstanceAt(doc, paragraphAt(doc, 1))!.numId).toBe(1);
    expect(listInstanceAt(doc, paragraphAt(doc, 2))!.numId).toBe(1);
  });

  it("has nothing to continue for the document's first list", () => {
    const doc = load(item("one", 1) + item("two", 1));
    expect(continueNumberingAt(doc, paragraphAt(doc, 0))).toBe(false);
  });

  it("is a no-op on a paragraph that is not a list item", () => {
    const doc = load(`<w:p><w:r><w:t>prose</w:t></w:r></w:p>`);
    expect(restartNumberingAt(doc, paragraphAt(doc, 0), 1)).toBe(false);
    expect(continueNumberingAt(doc, paragraphAt(doc, 0))).toBe(false);
  });
});

describe("numbering.xml round-trip", () => {
  it("survives save and reload, and leaves the part alone when untouched", () => {
    const untouched = load(item("a", 1));
    expect(strFromU8(unzipSync(untouched.save())["word/numbering.xml"])).toBe(NUMBERING_XML);

    const doc = load(item("a", 1) + item("b", 1));
    setNumberingLevel(doc, 1, 0, { format: "upperRoman", text: "%1)" });
    restartNumberingAt(doc, paragraphAt(doc, 1), 4);

    const reloaded = DocxDocument.load(doc.save());
    expect(reloaded.numberingLevel(1, 0)?.format).toBe("upperRoman");
    expect(reloaded.numberingLevel(1, 0)?.text).toBe("%1)");
    const instance = reloaded.numbering.instances.get(3)!;
    expect(instance.abstractNumId).toBe(0);
    expect(instance.overrides.get(0)?.startOverride).toBe(4);
  });
});
