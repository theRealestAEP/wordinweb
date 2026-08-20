import { describe, expect, it } from "vitest";
import { zipSync, strToU8 } from "fflate";
import { DocxDocument, serializeXml } from "@wordinweb/core";
import { DocumentSession } from "../src/session.js";
import { validateIntent } from "../src/validate.js";

const STYLES_XML =
  `<?xml version="1.0"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
  `<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>` +
  `<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/>` +
  `<w:basedOn w:val="Normal"/><w:rPr><w:b/></w:rPr></w:style>` +
  `<w:style w:type="character" w:styleId="Emphasis"><w:name w:val="Emphasis"/>` +
  `<w:rPr><w:i/></w:rPr></w:style>` +
  `</w:styles>`;

function makeDoc(): DocxDocument {
  const documentXml =
    `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
    `<w:body><w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t xml:space="preserve">head</w:t></w:r></w:p>` +
    `<w:p><w:r><w:t xml:space="preserve">body</w:t></w:r></w:p>` +
    `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`;
  return DocxDocument.load(
    zipSync({
      "[Content_Types].xml": strToU8(
        `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
          `<Default Extension="xml" ContentType="application/xml"/>` +
          `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
          `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`,
      ),
      "_rels/.rels": strToU8(
        `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
          `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`,
      ),
      "word/document.xml": strToU8(documentXml),
      "word/styles.xml": strToU8(STYLES_XML),
    }),
  );
}

/** Every replicated tree, including styles.xml — which is the part these
 * operations move, and so the part convergence has to be checked on. */
function xmlOf(s: DocumentSession): string {
  return [...s.doc.editableRoots(), s.doc.stylesTree()!].map((r) => serializeXml(r)).join("|");
}

const base = { clientId: "a", clientSeq: 1, base: 0 } as const;

describe("style definition operations over the wire", () => {
  it("creates the same definition on two replicas", () => {
    const a = new DocumentSession(makeDoc());
    const b = new DocumentSession(makeDoc());
    const intent = {
      ...base,
      kind: "createStyle" as const,
      style: {
        styleId: "PullQuote",
        type: "paragraph" as const,
        name: "Pull Quote",
        quickStyle: true,
        paragraph: { alignment: "center" as const },
        run: { italic: true },
      },
    };
    expect(a.submit(intent).kind).toBe("applied");
    expect(b.submit(intent).kind).toBe("applied");
    expect(xmlOf(a)).toBe(xmlOf(b));
    expect(xmlOf(a)).toContain(`w:styleId="PullQuote"`);
  });

  it("rejects a create whose id the replica already declares, changing nothing", () => {
    // The styleId is the rejection predicate a document address has no stable
    // id for: styles.xml is sequenced state, so every replica decides alike.
    const s = new DocumentSession(makeDoc());
    const before = xmlOf(s);
    expect(
      s.submit({
        ...base,
        kind: "createStyle",
        style: { styleId: "Heading1", type: "paragraph", name: "Mine" },
      }).kind,
    ).toBe("rejected");
    expect(xmlOf(s)).toBe(before);
  });

  it("modifies and deletes, and rejects both for a style it does not declare", () => {
    const s = new DocumentSession(makeDoc());
    expect(
      s.submit({ ...base, kind: "modifyStyle", styleId: "Heading1", patch: { uiPriority: 4 } }).kind,
    ).toBe("applied");
    expect(xmlOf(s)).toContain(`<w:uiPriority w:val="4"/>`);

    expect(s.submit({ ...base, clientSeq: 2, kind: "deleteStyle", styleId: "Emphasis" }).kind).toBe(
      "applied",
    );
    expect(xmlOf(s)).not.toContain(`w:styleId="Emphasis"`);

    const after = xmlOf(s);
    expect(s.submit({ ...base, clientSeq: 3, kind: "deleteStyle", styleId: "Emphasis" }).kind).toBe(
      "rejected",
    );
    expect(
      s.submit({ ...base, clientSeq: 4, kind: "modifyStyle", styleId: "Gone", patch: { name: "x" } })
        .kind,
    ).toBe("rejected");
    expect(xmlOf(s)).toBe(after);
  });

  it("re-points a deleted style's users identically on both replicas", () => {
    const a = new DocumentSession(makeDoc());
    const b = new DocumentSession(makeDoc());
    const intent = { ...base, kind: "deleteStyle" as const, styleId: "Heading1" };
    expect(a.submit(intent).kind).toBe("applied");
    expect(b.submit(intent).kind).toBe("applied");
    expect(xmlOf(a)).toBe(xmlOf(b));
    expect(xmlOf(a)).toContain(`<w:pStyle w:val="Normal"/>`);
  });
});

describe("style payload validation runs before sequencing", () => {
  const cases: [string, Record<string, unknown>][] = [
    ["a styleId with a space", { kind: "createStyle", style: { styleId: "Pull Quote", type: "paragraph", name: "x" } }],
    ["an unknown style type", { kind: "createStyle", style: { styleId: "X", type: "list", name: "x" } }],
    ["paragraph props on a character style", { kind: "createStyle", style: { styleId: "X", type: "character", name: "x", paragraph: { alignment: "center" } } }],
    ["an unknown paragraph property", { kind: "createStyle", style: { styleId: "X", type: "paragraph", name: "x", paragraph: { leading: 2 } } }],
    ["an out-of-range outline level", { kind: "createStyle", style: { styleId: "X", type: "paragraph", name: "x", paragraph: { outlineLevel: 9 } } }],
    // The four style types the engine writes, each with the payload that only
    // makes sense for it — a mismatch is refused rather than dropped.
    ["a numbering style with no numbering definition", { kind: "createStyle", style: { styleId: "X", type: "numbering", name: "x" } }],
    ["formatting on a numbering style", { kind: "createStyle", style: { styleId: "X", type: "numbering", name: "x", numbering: { numId: 1 }, run: { bold: true } } }],
    ["a numbering reference on a paragraph style", { kind: "createStyle", style: { styleId: "X", type: "paragraph", name: "x", numbering: { numId: 1 } } }],
    ["table properties on a paragraph style", { kind: "createStyle", style: { styleId: "X", type: "paragraph", name: "x", table: { borders: {} } } }],
    ["a cell-only border edge on a table style", { kind: "createStyle", style: { styleId: "X", type: "table", name: "x", table: { borders: { tl2br: { style: "single" } } } } }],
    ["an out-of-range border weight on a table style", { kind: "createStyle", style: { styleId: "X", type: "table", name: "x", table: { borders: { top: { style: "single", sz: 200 } } } } }],
    ["a linked companion for a character style", { kind: "createStyle", style: { styleId: "X", type: "character", name: "x", linked: true } }],
    ["an empty patch", { kind: "modifyStyle", styleId: "Heading1", patch: {} }],
    ["a bad colour", { kind: "modifyStyle", styleId: "Heading1", patch: { run: { color: "red" } } }],
    ["an unknown numbering format", { kind: "setNumberingLevel", blockId: 1, ilvl: 0, patch: { format: "hex" } }],
    ["an out-of-range list level", { kind: "setNumberingLevel", blockId: 1, ilvl: 9, patch: { format: "decimal" } }],
  ];
  for (const [what, intent] of cases) {
    it(`rejects ${what}`, () => {
      expect(validateIntent({ ...base, ...intent } as never)).not.toBeNull();
    });
  }

  it("accepts a well-formed style spec and level patch", () => {
    expect(
      validateIntent({
        ...base,
        kind: "createStyle",
        style: {
          styleId: "Pull-Quote_2",
          type: "paragraph",
          name: "Pull Quote",
          basedOn: "Normal",
          quickStyle: true,
          uiPriority: 29,
          paragraph: { alignment: "both", spacingAfterPt: 12, outlineLevel: 0 },
          run: { italic: true, color: "#556677", fontSizePt: 13 },
        },
      } as never),
    ).toBeNull();
    expect(
      validateIntent({
        ...base,
        kind: "setNumberingLevel",
        blockId: 1,
        ilvl: null,
        patch: { format: "upperRoman", text: "%1)", indentLeftPt: 36 },
      } as never),
    ).toBeNull();
  });

  it("accepts a linked pair, a table style and a numbering style", () => {
    const specs = [
      { styleId: "Lead", type: "paragraph", name: "Lead", linked: true, run: { bold: true } },
      { styleId: "Ledger", type: "table", name: "Ledger", table: { borders: { top: { style: "single", sz: 8, color: "4472C4" }, insideV: { style: "dotted" } } } },
      { styleId: "Steps", type: "numbering", name: "Steps", numbering: { numId: 3 } },
    ];
    for (const style of specs) {
      expect(validateIntent({ ...base, kind: "createStyle", style } as never), style.styleId).toBeNull();
    }
  });

  it("bounds the character style a run patch carries", () => {
    const patch = (characterStyleId: unknown) => ({
      ...base,
      kind: "formatRun",
      blockId: 1,
      runId: 2,
      patch: { characterStyleId },
    });
    expect(validateIntent(patch("Emphasis") as never)).toBeNull();
    expect(validateIntent(patch(null) as never)).toBeNull();
    expect(validateIntent(patch("has space") as never)).not.toBeNull();
    expect(validateIntent(patch("x".repeat(300)) as never)).not.toBeNull();
  });

  it("bounds the text-effect a run patch carries", () => {
    const patch = (textEffect: unknown) => ({
      ...base,
      kind: "formatRun",
      blockId: 1,
      runId: 2,
      patch: { textEffect },
    });
    expect(validateIntent(patch(null) as never)).toBeNull();
    expect(validateIntent(patch({ shadow: true }) as never)).toBeNull();
    expect(validateIntent(patch({ outline: { color: "#4472C4", widthPt: 1 } }) as never)).toBeNull();
    expect(validateIntent(patch({ outline: null, shadow: false }) as never)).toBeNull();
    expect(validateIntent(patch({ shadow: "yes" }) as never)).not.toBeNull();
    expect(validateIntent(patch({ outline: { color: "not-a-color", widthPt: 1 } }) as never)).not.toBeNull();
    expect(validateIntent(patch({ outline: { color: "#4472C4", widthPt: 0 } }) as never)).not.toBeNull();
    expect(validateIntent(patch({ outline: { color: "#4472C4", widthPt: 100 } }) as never)).not.toBeNull();
    expect(validateIntent(patch("bogus") as never)).not.toBeNull();
  });
});
