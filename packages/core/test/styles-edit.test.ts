import { unzipSync, strFromU8 } from "fflate";
import { describe, expect, it } from "vitest";
import { DocxDocument } from "../src/docx.js";
import { applyRunFormat, formatPatchFrom, summarizeSelection } from "../src/edit/commands.js";
import { setParagraphStyle } from "../src/edit/blocks.js";
import {
  createStyle,
  deleteStyle,
  listStyles,
  modifyStyle,
  styleIdFromName,
  styleUsageCount,
  uniqueStyleId,
} from "../src/edit/styles.js";
import { layoutDocument } from "../src/layout/engine.js";
import { ApproxMeasurer } from "../src/layout/measure.js";
import { serializeXml } from "../src/xml.js";
import { makeDocx, wrapDocument } from "./helpers.js";

const measurer = new ApproxMeasurer();

/**
 * A minimal Word-shaped styles part: docDefaults first, then style entries,
 * which is the order every fixture in the parity corpus writes.
 */
const STYLES_XML =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
  `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
  `<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/>` +
  `<w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults>` +
  `<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>` +
  `<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/>` +
  `<w:basedOn w:val="Normal"/><w:uiPriority w:val="9"/><w:qFormat/>` +
  `<w:pPr><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style>` +
  `<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont">` +
  `<w:name w:val="Default Paragraph Font"/><w:uiPriority w:val="1"/><w:semiHidden/></w:style>` +
  `<w:style w:type="character" w:styleId="Emphasis"><w:name w:val="Emphasis"/>` +
  `<w:basedOn w:val="DefaultParagraphFont"/><w:uiPriority w:val="20"/><w:qFormat/>` +
  `<w:rPr><w:i/></w:rPr></w:style>` +
  `</w:styles>`;

function load(body: string, styles = STYLES_XML): DocxDocument {
  return DocxDocument.load(
    makeDocx({ "word/document.xml": wrapDocument(body), "word/styles.xml": styles }),
  );
}

function para(text: string, styleId?: string): string {
  const pPr = styleId ? `<w:pPr><w:pStyle w:val="${styleId}"/></w:pPr>` : "";
  return `<w:p>${pPr}<w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
}

function stylesXml(doc: DocxDocument): string {
  return serializeXml(doc.stylesTree()!);
}

/** The styles.xml a save() actually writes into the package. */
function savedStylesXml(doc: DocxDocument): string {
  return strFromU8(unzipSync(doc.save())["word/styles.xml"]);
}

/** The definition of one style, verbatim, out of the serialized part. */
function styleEntry(xml: string, styleId: string): string {
  const match = xml.match(new RegExp(`<w:style [^>]*w:styleId="${styleId}"[^>]*>.*?</w:style>`, "s"));
  return match ? match[0] : "";
}

function firstRunProps(doc: DocxDocument, blockIndex: number) {
  const para = doc.sections[0].blocks[blockIndex];
  if (para.type !== "paragraph") throw new Error("not a paragraph");
  const run = para.children.find((c) => c.type === "run");
  if (run?.type !== "run") throw new Error("no run");
  return doc.effectiveRunProps(para, run.props);
}

function effectiveParaProps(doc: DocxDocument, blockIndex: number) {
  const para = doc.sections[0].blocks[blockIndex];
  if (para.type !== "paragraph") throw new Error("not a paragraph");
  return doc.effectiveParaProps(para);
}

// ---------------------------------------------------------------------------

describe("createStyle", () => {
  it("writes the child order and attribute order Word writes", () => {
    const doc = load(para("body"));
    expect(
      createStyle(doc, {
        styleId: "CaseCaption",
        type: "paragraph",
        name: "Case Caption",
        basedOn: "Normal",
        next: "Normal",
        quickStyle: true,
        uiPriority: 12,
        paragraph: { alignment: "center", spacingAfterPt: 12, keepNext: true },
        run: { bold: true, fontSizePt: 14 },
      }),
    ).toBe(true);

    const entry = styleEntry(stylesXml(doc), "CaseCaption");
    // Attributes: type, customStyle, styleId (791 custom styles in the corpus
    // are written this way, and none other).
    expect(entry).toContain(
      `<w:style w:type="paragraph" w:customStyle="1" w:styleId="CaseCaption">`,
    );
    // Children in CT_Style sequence: name, basedOn, next, uiPriority, qFormat,
    // pPr, rPr — and pPr's own children in CT_PPrBase sequence.
    const order = [...entry.matchAll(/<w:(name|basedOn|next|uiPriority|qFormat|pPr|rPr)[ />]/g)].map(
      (m) => m[1],
    );
    expect(order).toEqual(["name", "basedOn", "next", "uiPriority", "qFormat", "pPr", "rPr"]);
    const pPrOrder = [...entry.matchAll(/<w:(keepNext|spacing|jc)[ />]/g)].map((m) => m[1]);
    expect(pPrOrder).toEqual(["keepNext", "spacing", "jc"]);
    expect(entry).toContain(`<w:spacing w:after="240"/>`);
    expect(entry).toContain(`<w:jc w:val="center"/>`);
    expect(entry).toContain(`<w:rPr><w:b/><w:bCs/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr>`);
  });

  it("defaults basedOn per style type, and omits a self-referential next", () => {
    const doc = load(para("body"));
    createStyle(doc, { styleId: "Lead", type: "paragraph", name: "Lead", next: "Lead" });
    createStyle(doc, { styleId: "Cite", type: "character", name: "Cite" });
    expect(styleEntry(stylesXml(doc), "Lead")).toContain(`<w:basedOn w:val="Normal"/>`);
    expect(styleEntry(stylesXml(doc), "Lead")).not.toContain("w:next");
    expect(styleEntry(stylesXml(doc), "Cite")).toContain(
      `<w:basedOn w:val="DefaultParagraphFont"/>`,
    );
  });

  it("rejects an id styles.xml already declares", () => {
    const doc = load(para("body"));
    expect(createStyle(doc, { styleId: "Heading1", type: "paragraph", name: "Mine" })).toBe(false);
  });

  it("is applied by setParagraphStyle and resolves through the cascade", () => {
    const doc = load(para("body"));
    createStyle(doc, {
      styleId: "BigRed",
      type: "paragraph",
      name: "Big Red",
      basedOn: "Heading1",
      run: { color: "#CC0000" },
    });
    const target = doc.docRoot.children[0].children[0].children[0];
    expect(setParagraphStyle(doc, [target], "BigRed")).toBe(true);
    const props = firstRunProps(doc, 0);
    expect(props.color).toBe("#CC0000");
    // Bold and 16pt come from Heading1, the style BigRed is based on.
    expect(props.bold).toBe(true);
    expect(props.size).toBeCloseTo((16 * 4) / 3, 5);
  });
});

describe("modifyStyle", () => {
  it("patches a definition without disturbing the rest of it", () => {
    const doc = load(para("h", "Heading1"));
    expect(modifyStyle(doc, "Heading1", { run: { color: "#123456" }, uiPriority: 3 })).toBe(true);
    const entry = styleEntry(stylesXml(doc), "Heading1");
    expect(entry).toContain(`<w:color w:val="123456"/>`);
    expect(entry).toContain(`<w:uiPriority w:val="3"/>`);
    // The outline level and the bold it already carried survive.
    expect(entry).toContain(`<w:outlineLvl w:val="0"/>`);
    expect(entry).toContain("<w:b/>");
    // uiPriority still precedes qFormat and pPr.
    const order = [...entry.matchAll(/<w:(name|basedOn|uiPriority|qFormat|pPr|rPr)[ />]/g)].map(
      (m) => m[1],
    );
    expect(order).toEqual(["name", "basedOn", "uiPriority", "qFormat", "pPr", "rPr"]);
  });

  it("patches one spacing attribute and leaves the others alone", () => {
    const doc = load(para("body"));
    createStyle(doc, {
      styleId: "Body",
      type: "paragraph",
      name: "Body",
      paragraph: { spacingBeforePt: 6, spacingAfterPt: 6, lineMultiple: 1.5 },
    });
    modifyStyle(doc, "Body", { paragraph: { spacingAfterPt: 18 } });
    expect(styleEntry(stylesXml(doc), "Body")).toContain(
      `<w:spacing w:before="120" w:after="360" w:line="360" w:lineRule="auto"/>`,
    );
  });

  it("refuses a self-referential basedOn", () => {
    const doc = load(para("body"));
    expect(modifyStyle(doc, "Heading1", { basedOn: "Heading1" })).toBe(false);
  });

  it("rejects a style styles.xml does not declare", () => {
    const doc = load(para("body"));
    expect(modifyStyle(doc, "Nope", { name: "Nope" })).toBe(false);
  });

  it("re-resolves the cascade live and repaints the affected paragraphs", () => {
    const doc = load(para("heading", "Heading1") + para("body"));
    /** The heading's painted line height, and where the body line lands. */
    const paint = () => {
      const items = layoutDocument(doc, measurer).pages[0].items.filter((i) => i.kind === "text");
      return { headingHeight: items[0].lineHeight, bodyTop: items[1].lineTop };
    };
    const before = paint();
    const beforeSig = doc.layoutGlobalSig();

    expect(modifyStyle(doc, "Heading1", { run: { fontSizePt: 48 } })).toBe(true);

    // The paragraph's own XML is untouched — only the definition moved — so the
    // line-break cache has to be invalidated by the global signature.
    expect(doc.layoutGlobalSig()).not.toBe(beforeSig);
    expect(firstRunProps(doc, 0).size).toBeCloseTo((48 * 4) / 3, 5);

    const after = paint();
    expect(after.headingHeight).toBeGreaterThan(before.headingHeight);
    // The paragraph BELOW moved down, which is what "the document repainted"
    // means: the change propagated past the styled paragraph itself.
    expect(after.bodyTop).toBeGreaterThan(before.bodyTop);
  });

  it("re-resolves paragraph properties too, not just runs", () => {
    const doc = load(para("body", "Heading1"));
    expect(effectiveParaProps(doc, 0).alignment ?? "left").toBe("left");
    modifyStyle(doc, "Heading1", { paragraph: { alignment: "center" } });
    expect(effectiveParaProps(doc, 0).alignment).toBe("center");
  });
});

describe("deleteStyle", () => {
  it("re-points users at the deleted style's basedOn", () => {
    const doc = load(para("a", "Subhead") + para("b"));
    createStyle(doc, {
      styleId: "Subhead",
      type: "paragraph",
      name: "Subhead",
      basedOn: "Heading1",
    });
    expect(deleteStyle(doc, "Subhead")).toBe(true);
    expect(stylesXml(doc)).not.toContain(`w:styleId="Subhead"`);
    expect(serializeXml(doc.docRoot)).toContain(`<w:pStyle w:val="Heading1"/>`);
    // The heir's formatting is what the paragraph paints now.
    expect(firstRunProps(doc, 0).bold).toBe(true);
  });

  it("removes the reference outright when there is no surviving parent", () => {
    const doc = load(para("a", "Loose"));
    createStyle(doc, { styleId: "Loose", type: "paragraph", name: "Loose", basedOn: null });
    expect(deleteStyle(doc, "Loose")).toBe(true);
    expect(serializeXml(doc.docRoot)).not.toContain("w:pStyle");
  });

  it("re-parents a style that was based on the deleted one", () => {
    const doc = load(para("a"));
    createStyle(doc, { styleId: "Mid", type: "paragraph", name: "Mid", basedOn: "Heading1" });
    createStyle(doc, { styleId: "Leaf", type: "paragraph", name: "Leaf", basedOn: "Mid" });
    deleteStyle(doc, "Mid");
    expect(styleEntry(stylesXml(doc), "Leaf")).toContain(`<w:basedOn w:val="Heading1"/>`);
  });

  it("re-points character-style runs, and leaves paragraph styles alone", () => {
    const doc = load(
      `<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr>` +
        `<w:r><w:rPr><w:rStyle w:val="Emphasis"/></w:rPr><w:t>x</w:t></w:r></w:p>`,
    );
    expect(deleteStyle(doc, "Emphasis")).toBe(true);
    const xml = serializeXml(doc.docRoot);
    expect(xml).toContain(`<w:rStyle w:val="DefaultParagraphFont"/>`);
    expect(xml).toContain(`<w:pStyle w:val="Heading1"/>`);
  });

  it("refuses to delete the default paragraph style", () => {
    const doc = load(para("a"));
    expect(deleteStyle(doc, "Normal")).toBe(false);
    expect(doc.styles.byId.has("Normal")).toBe(true);
  });

  it("rejects a style styles.xml does not declare", () => {
    expect(deleteStyle(load(para("a")), "Nope")).toBe(false);
  });
});

describe("character styles over a selection", () => {
  /** The one run of the one paragraph, as a selection segment. */
  function wholeRun(doc: DocxDocument) {
    const paraBlock = doc.sections[0].blocks[0];
    if (paraBlock.type !== "paragraph") throw new Error("not a paragraph");
    const run = paraBlock.children.find((c) => c.type === "run");
    if (run?.type !== "run") throw new Error("no run");
    const t = run.content.find((c) => c.kind === "text");
    if (t?.kind !== "text") throw new Error("no text");
    return {
      run,
      t: run.src?.children.find((c) => c.name.endsWith("t")) ?? null,
      start: 0,
      end: t.text.length,
      props: doc.effectiveRunProps(paraBlock, run.props),
    };
  }

  it("writes w:rStyle and resolves the style's formatting", () => {
    const doc = load(para("emphatic"));
    applyRunFormat(doc, [wholeRun(doc)], { characterStyleId: "Emphasis" });
    expect(serializeXml(doc.docRoot)).toContain(`<w:rStyle w:val="Emphasis"/>`);
    expect(firstRunProps(doc, 0).italic).toBe(true);
  });

  it("puts rStyle first in rPr, ahead of direct formatting", () => {
    const doc = load(para("emphatic"));
    applyRunFormat(doc, [wholeRun(doc)], { bold: true });
    applyRunFormat(doc, [wholeRun(doc)], { characterStyleId: "Emphasis" });
    expect(serializeXml(doc.docRoot)).toContain(
      `<w:rPr><w:rStyle w:val="Emphasis"/><w:b/><w:bCs/></w:rPr>`,
    );
  });

  it("splits the run when only part of it is selected", () => {
    const doc = load(para("abcdef"));
    const seg = wholeRun(doc);
    applyRunFormat(doc, [{ ...seg, start: 2, end: 4 }], { characterStyleId: "Emphasis" });
    const xml = serializeXml(doc.docRoot);
    expect(xml).toContain(`<w:r><w:t xml:space="preserve">ab</w:t></w:r>`);
    expect(xml).toContain(
      `<w:r><w:rPr><w:rStyle w:val="Emphasis"/></w:rPr><w:t xml:space="preserve">cd</w:t></w:r>`,
    );
    expect(xml).toContain(`<w:r><w:t xml:space="preserve">ef</w:t></w:r>`);
  });

  it("removes the reference with null", () => {
    const doc = load(para("plain"));
    applyRunFormat(doc, [wholeRun(doc)], { characterStyleId: "Emphasis" });
    applyRunFormat(doc, [wholeRun(doc)], { characterStyleId: null });
    expect(serializeXml(doc.docRoot)).not.toContain("w:rStyle");
    expect(firstRunProps(doc, 0).italic).toBeUndefined();
  });

  it("is reported by getSelectionFormat's summary", () => {
    const doc = load(para("styled"));
    expect(summarizeSelection([wholeRun(doc)])?.characterStyleId).toBe(null);
    applyRunFormat(doc, [wholeRun(doc)], { characterStyleId: "Emphasis" });
    expect(summarizeSelection([wholeRun(doc)])?.characterStyleId).toBe("Emphasis");
  });

  it("reports undefined when the selected runs disagree", () => {
    const doc = load(
      `<w:p><w:r><w:rPr><w:rStyle w:val="Emphasis"/></w:rPr><w:t>a</w:t></w:r>` +
        `<w:r><w:t>b</w:t></w:r></w:p>`,
    );
    const paraBlock = doc.sections[0].blocks[0];
    if (paraBlock.type !== "paragraph") throw new Error("not a paragraph");
    const runs = paraBlock.children.filter((c) => c.type === "run");
    const segments = runs.map((r) => {
      if (r.type !== "run") throw new Error("not a run");
      return { run: r, t: null, start: 0, end: 1, props: r.props };
    });
    expect(summarizeSelection(segments)?.characterStyleId).toBeUndefined();
  });
});

describe("listStyles gallery data", () => {
  it("reports identity, cascade, quick-style flag, usage and a preview", () => {
    const doc = load(para("h", "Heading1") + para("h2", "Heading1") + para("b"));
    const rows = listStyles(doc);
    const heading = rows.find((r) => r.id === "Heading1")!;
    expect(heading.name).toBe("heading 1");
    expect(heading.type).toBe("paragraph");
    expect(heading.basedOn).toBe("Normal");
    expect(heading.quickStyle).toBe(true);
    expect(heading.uiPriority).toBe(9);
    expect(heading.usageCount).toBe(2);
    // The preview resolves through the chain AND docDefaults, so a gallery can
    // paint the entry without resolving anything itself.
    expect(heading.preview.bold).toBe(true);
    expect(heading.preview.size).toBeCloseTo((16 * 4) / 3, 5);
    expect(heading.preview.font).toBe("Calibri");

    const emphasis = rows.find((r) => r.id === "Emphasis")!;
    expect(emphasis.type).toBe("character");
    expect(emphasis.preview.italic).toBe(true);
    expect(emphasis.usageCount).toBe(0);
  });

  it("counts character-style runs against the character style", () => {
    const doc = load(
      `<w:p><w:r><w:rPr><w:rStyle w:val="Emphasis"/></w:rPr><w:t>a</w:t></w:r>` +
        `<w:r><w:rPr><w:rStyle w:val="Emphasis"/></w:rPr><w:t>b</w:t></w:r></w:p>`,
    );
    expect(listStyles(doc).find((r) => r.id === "Emphasis")!.usageCount).toBe(2);
    expect(styleUsageCount(doc, "Emphasis")).toBe(2);
  });

  it("filters by type and by the quick-style flag, and sorts by uiPriority", () => {
    const doc = load(para("a"));
    const paragraphs = listStyles(doc, { type: "paragraph" });
    expect(paragraphs.every((r) => r.type === "paragraph")).toBe(true);
    expect(listStyles(doc, { quickStyleOnly: true }).map((r) => r.id)).toEqual([
      "Heading1",
      "Emphasis",
      "Normal",
    ]);
  });
});

describe("format painter (formatPatchFrom)", () => {
  it("carries the booleans, including off, so it can un-bold a target", () => {
    const patch = formatPatchFrom({ bold: false, italic: true, underline: false, strike: false });
    expect(patch.bold).toBe(false);
    expect(patch.italic).toBe(true);
    // A source with no superscript returns the target to the baseline.
    expect(patch.verticalAlign).toBe(null);
  });

  it("leaves out what the source selection did not agree on", () => {
    const mixed = formatPatchFrom({ bold: true, italic: false, underline: false, strike: false });
    expect("fontSizePt" in mixed).toBe(false);
    expect("fontFamily" in mixed).toBe(false);
    expect("color" in mixed).toBe(false);
    expect("characterStyleId" in mixed).toBe(false);
  });

  it("carries an agreed value, and an agreed absence", () => {
    const patch = formatPatchFrom({
      bold: true,
      italic: false,
      underline: false,
      strike: false,
      fontSizePt: 13,
      fontFamily: "Georgia",
      color: "#334455",
      characterStyleId: null,
    });
    expect(patch.fontSizePt).toBe(13);
    expect(patch.fontFamily).toBe("Georgia");
    expect(patch.color).toBe("#334455");
    // The source carried no character style, so painting removes the target's.
    expect(patch.characterStyleId).toBe(null);
  });

  it("round-trips a real selection: copy from one run, paint onto another", () => {
    const doc = load(
      `<w:p><w:r><w:rPr><w:rStyle w:val="Emphasis"/><w:b/><w:color w:val="AA0000"/>` +
        `<w:sz w:val="28"/></w:rPr><w:t>source</w:t></w:r>` +
        `<w:r><w:t>target</w:t></w:r></w:p>`,
    );
    const paraBlock = doc.sections[0].blocks[0];
    if (paraBlock.type !== "paragraph") throw new Error("not a paragraph");
    const runs = paraBlock.children.filter((c) => c.type === "run");
    const segment = (index: number) => {
      const run = runs[index];
      if (run?.type !== "run") throw new Error("no run");
      return {
        run,
        t: run.src?.children.find((c) => c.name.endsWith("t")) ?? null,
        start: 0,
        end: 6,
        props: doc.effectiveRunProps(paraBlock, run.props),
      };
    };
    const copied = summarizeSelection([segment(0)])!;
    applyRunFormat(doc, [segment(1)], formatPatchFrom(copied));

    const painted = serializeXml(doc.docRoot).match(
      /<w:r>(?:(?!<\/w:r>).)*target(?:(?!<\/w:r>).)*<\/w:r>/s,
    )![0];
    expect(painted).toContain(`<w:rStyle w:val="Emphasis"/>`);
    expect(painted).toContain("<w:b/>");
    expect(painted).toContain(`<w:color w:val="AA0000"/>`);
    expect(painted).toContain(`<w:sz w:val="28"/>`);
  });
});

describe("styleId derivation", () => {
  it("matches Word's: drop whitespace, keep alphanumerics and hyphens", () => {
    expect(styleIdFromName("Case Caption")).toBe("CaseCaption");
    expect(styleIdFromName("E-mail Signature Char")).toBe("E-mailSignatureChar");
    expect(styleIdFromName("C-Heading 1")).toBe("C-Heading1");
  });

  it("disambiguates a taken id with a counter, as Word does", () => {
    const doc = load(para("a"));
    expect(uniqueStyleId(doc, "Fresh")).toBe("Fresh");
    expect(uniqueStyleId(doc, "Heading1")).toBe("Heading10");
  });
});

// ---------------------------------------------------------------------------
// SCENARIO SPEC — the round-trip gate
// ---------------------------------------------------------------------------

/**
 * SCENARIO: "style session".
 *
 * A user creates a style, applies it, modifies it, applies a character style,
 * and deletes a different style — the sequence a styles pane produces in one
 * sitting. The gate is threefold:
 *
 *  1. Every style the session did NOT touch is byte-identical in the saved
 *     part. A style editor that reformats the whole file would pass a
 *     "contains" assertion and fail this one.
 *  2. The session's own edits survive save and reload — the retained-root and
 *     dirty-flag machinery is what carries them.
 *  3. The saved definitions are shaped the way Word writes them, so Word opens
 *     the file without a repair prompt: CT_Style child sequence, the
 *     type/customStyle/styleId attribute order, and no dangling reference to
 *     the deleted style anywhere in the document.
 */
describe("SCENARIO: a style session survives save and reload", () => {
  function session(): DocxDocument {
    const doc = load(para("Title text") + para("body one", "Doomed") + para("body two"));
    createStyle(doc, {
      styleId: "Doomed",
      type: "paragraph",
      name: "Doomed",
      basedOn: "Heading1",
    });
    createStyle(doc, {
      styleId: "PullQuote",
      type: "paragraph",
      name: "Pull Quote",
      basedOn: "Normal",
      next: "Normal",
      quickStyle: true,
      uiPriority: 29,
      paragraph: { alignment: "center", indentLeftPt: 36, spacingAfterPt: 12 },
      run: { italic: true, fontSizePt: 13 },
    });
    const firstParagraph = doc.docRoot.children[0].children[0].children[0];
    setParagraphStyle(doc, [firstParagraph], "PullQuote");
    modifyStyle(doc, "PullQuote", { run: { color: "#556677" }, paragraph: { keepNext: true } });
    deleteStyle(doc, "Doomed");
    return doc;
  }

  it("leaves untouched styles byte-identical in the saved part", () => {
    const saved = savedStylesXml(session());
    for (const styleId of ["Normal", "Heading1", "DefaultParagraphFont", "Emphasis"]) {
      expect(styleEntry(saved, styleId)).toBe(styleEntry(STYLES_XML, styleId));
    }
    // docDefaults still leads the part, ahead of every style entry.
    expect(saved.indexOf("<w:docDefaults>")).toBeLessThan(saved.indexOf("<w:style "));
  });

  it("writes the created style the way Word writes one", () => {
    const entry = styleEntry(savedStylesXml(session()), "PullQuote");
    expect(entry).toBe(
      `<w:style w:type="paragraph" w:customStyle="1" w:styleId="PullQuote">` +
        `<w:name w:val="Pull Quote"/>` +
        `<w:basedOn w:val="Normal"/>` +
        `<w:next w:val="Normal"/>` +
        `<w:uiPriority w:val="29"/>` +
        `<w:qFormat/>` +
        `<w:pPr><w:keepNext/><w:spacing w:after="240"/><w:ind w:left="720"/><w:jc w:val="center"/></w:pPr>` +
        `<w:rPr><w:i/><w:iCs/><w:color w:val="556677"/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr>` +
        `</w:style>`,
    );
  });

  it("leaves no reference to the deleted style, and re-points its users", () => {
    const doc = session();
    const files = unzipSync(doc.save());
    const savedDocument = strFromU8(files["word/document.xml"]);
    expect(strFromU8(files["word/styles.xml"])).not.toContain(`w:styleId="Doomed"`);
    expect(savedDocument).not.toContain("Doomed");
    // Its user inherited Doomed's basedOn rather than losing its formatting.
    expect(savedDocument).toContain(`<w:pStyle w:val="Heading1"/>`);
  });

  it("reloads with the session's cascade intact", () => {
    const reloaded = DocxDocument.load(session().save());
    const pullQuote = reloaded.styles.byId.get("PullQuote")!;
    expect(pullQuote.name).toBe("Pull Quote");
    expect(pullQuote.basedOn).toBe("Normal");
    expect(pullQuote.rPr?.italic).toBe(true);
    expect(pullQuote.rPr?.color).toBe("#556677");
    expect(pullQuote.pPr?.alignment).toBe("center");
    expect(reloaded.styles.byId.has("Doomed")).toBe(false);

    // The applied style still resolves, and the deleted one's user landed on
    // the heir — both read back off a fresh parse of the saved bytes.
    expect(firstRunProps(reloaded, 0).italic).toBe(true);
    expect(firstRunProps(reloaded, 0).color).toBe("#556677");
    expect(firstRunProps(reloaded, 1).bold).toBe(true);

    // And the whole thing lays out: a reloaded package is a real document.
    expect(layoutDocument(reloaded, measurer).pages.length).toBeGreaterThan(0);
  });

  it("saves the same bytes twice, so the session is deterministic", () => {
    const doc = session();
    expect(savedStylesXml(doc)).toBe(savedStylesXml(doc));
  });

  it("leaves styles.xml alone when the session changes nothing", () => {
    const doc = load(para("untouched"));
    expect(savedStylesXml(doc)).toBe(STYLES_XML);
  });
});

// ---------------------------------------------------------------------------
// Linked companions (w:link) and the table / numbering style types
// ---------------------------------------------------------------------------

/** The w:style element for `id`, as XML. */
function definition(doc: DocxDocument, id: string): string {
  const xml = stylesXml(doc);
  const at = xml.indexOf(`w:styleId="${id}"`);
  expect(at, `definition for ${id}`).toBeGreaterThan(-1);
  return xml.slice(xml.lastIndexOf("<w:style ", at), xml.indexOf("</w:style>", at) + 10);
}

describe("createStyle: the linked character companion", () => {
  const spec = {
    styleId: "PullQuote",
    type: "paragraph" as const,
    name: "Pull Quote",
    quickStyle: true,
    uiPriority: 30,
    run: { italic: true, fontSizePt: 13 },
  };

  it("writes both halves, pointing at each other", () => {
    const doc = load(para("body"));
    expect(createStyle(doc, { ...spec, linked: true })).toBe(true);
    // Word's naming: the companion takes the id and the name with "Char".
    expect(definition(doc, "PullQuote")).toContain(`<w:link w:val="PullQuoteChar"/>`);
    const companion = definition(doc, "PullQuoteChar");
    expect(companion).toContain(`w:type="character"`);
    expect(companion).toContain(`<w:name w:val="Pull Quote Char"/>`);
    expect(companion).toContain(`<w:link w:val="PullQuote"/>`);
    expect(companion).toContain(`<w:basedOn w:val="DefaultParagraphFont"/>`);
  });

  it("gives the companion the paragraph style's run properties, and only those", () => {
    const doc = load(para("body"));
    createStyle(doc, { ...spec, linked: true, paragraph: { alignment: "center" } });
    const companion = definition(doc, "PullQuoteChar");
    expect(companion).toContain(`<w:i/>`);
    expect(companion).toContain(`<w:sz w:val="26"/>`);
    // A character style has no paragraph properties, so the alignment stays
    // on the paragraph half alone.
    expect(companion).not.toContain(`<w:jc`);
    expect(definition(doc, "PullQuote")).toContain(`<w:jc w:val="center"/>`);
  });

  it("is opt-in: without it there is no companion and no w:link", () => {
    const doc = load(para("body"));
    createStyle(doc, spec);
    expect(stylesXml(doc)).not.toContain("PullQuoteChar");
    expect(definition(doc, "PullQuote")).not.toContain("<w:link");
  });

  it("refuses rather than renaming when the companion id is taken", () => {
    const taken =
      STYLES_XML.replace(
        "</w:styles>",
        `<w:style w:type="character" w:styleId="PullQuoteChar"><w:name w:val="Squatter"/></w:style></w:styles>`,
      );
    const doc = load(para("body"), taken);
    // Deterministic refusal, not a second id: two replicas applying the same
    // operation have to reach the same verdict.
    expect(createStyle(doc, { ...spec, linked: true })).toBe(false);
    expect(stylesXml(doc)).not.toContain("PullQuote<");
    expect(definition(doc, "PullQuoteChar")).toContain("Squatter");
  });

  it("refuses a companion for a style that cannot have one", () => {
    const doc = load(para("body"));
    expect(createStyle(doc, { styleId: "Marker", type: "character", name: "Marker", linked: true })).toBe(true);
    // The flag is paragraph-only, so it is ignored rather than obeyed.
    expect(stylesXml(doc)).not.toContain("MarkerChar");
  });
});

describe("modifyStyle: a linked pair stays one look", () => {
  function linkedDoc(): DocxDocument {
    const doc = load(para("body"));
    createStyle(doc, {
      styleId: "Lead", type: "paragraph", name: "Lead", linked: true,
      run: { bold: true },
    });
    return doc;
  }

  it("carries a run change from the paragraph half to the character half", () => {
    const doc = linkedDoc();
    expect(modifyStyle(doc, "Lead", { run: { italic: true, color: "C00000" } })).toBe(true);
    for (const id of ["Lead", "LeadChar"]) {
      expect(definition(doc, id), id).toContain(`<w:i/>`);
      expect(definition(doc, id), id).toContain(`<w:color w:val="C00000"/>`);
    }
  });

  it("carries it the other way too", () => {
    const doc = linkedDoc();
    expect(modifyStyle(doc, "LeadChar", { run: { fontSizePt: 18 } })).toBe(true);
    expect(definition(doc, "Lead")).toContain(`<w:sz w:val="36"/>`);
  });

  it("does NOT carry a rename or a parent change", () => {
    const doc = linkedDoc();
    modifyStyle(doc, "Lead", { name: "Standfirst", basedOn: "Heading1" });
    // Only the run properties are shared; the companion keeps its own name
    // and its own place in the character cascade.
    expect(definition(doc, "LeadChar")).toContain(`<w:name w:val="Lead Char"/>`);
    expect(definition(doc, "LeadChar")).toContain(`<w:basedOn w:val="DefaultParagraphFont"/>`);
  });

  it("leaves no dangling w:link when half the pair is deleted", () => {
    const doc = linkedDoc();
    expect(deleteStyle(doc, "Lead")).toBe(true);
    expect(stylesXml(doc)).not.toContain(`w:styleId="Lead"`);
    // The survivor's w:link would otherwise name a style styles.xml no longer
    // declares — the dangling reference deleteStyle exists to avoid.
    expect(definition(doc, "LeadChar")).not.toContain("<w:link");
  });
});

describe("createStyle: table and numbering styles", () => {
  it("creates a table style based on TableNormal, with a grid", () => {
    const doc = load(para("body"));
    expect(
      createStyle(doc, {
        styleId: "Ledger",
        type: "table",
        name: "Ledger",
        table: {
          borders: {
            top: { style: "single", sz: 8, color: "4472C4" },
            insideH: { style: "dotted", sz: 4 },
          },
        },
        run: { bold: true },
      }),
    ).toBe(true);
    const style = definition(doc, "Ledger");
    expect(style).toContain(`w:type="table"`);
    // Word's default parent for a table style.
    expect(style).toContain(`<w:basedOn w:val="TableNormal"/>`);
    expect(style).toContain(`<w:top w:val="single" w:sz="8" w:space="0" w:color="4472C4"/>`);
    expect(style).toContain(`<w:insideH w:val="dotted" w:sz="4" w:space="0" w:color="auto"/>`);
    // A table style's w:rPr is the default for every run in the table.
    expect(style).toContain(`<w:b/>`);
    // …and the toolbar's list offers it, which is what makes it applyable.
    expect(listStyles(doc).some((entry) => entry.id === "Ledger" && entry.type === "table")).toBe(true);
  });

  it("creates a numbering style naming a numbering definition", () => {
    const doc = load(para("body"));
    expect(
      createStyle(doc, { styleId: "Steps", type: "numbering", name: "Steps", numbering: { numId: 4 } }),
    ).toBe(true);
    const style = definition(doc, "Steps");
    expect(style).toContain(`w:type="numbering"`);
    expect(style).toContain(`<w:pPr><w:numPr><w:numId w:val="4"/></w:numPr></w:pPr>`);
    // Word writes no parent for a numbering style.
    expect(style).not.toContain("<w:basedOn");
  });

  it("refuses a numbering style with nothing to number", () => {
    const doc = load(para("body"));
    expect(createStyle(doc, { styleId: "Steps", type: "numbering", name: "Steps" })).toBe(false);
    expect(stylesXml(doc)).not.toContain("Steps");
  });
});
