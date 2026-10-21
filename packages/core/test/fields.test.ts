import { describe, expect, it } from "vitest";
import { resolveField, type FieldContext } from "../src/layout/inline.js";
import { DocxDocument } from "../src/docx.js";
import { layoutDocument, layoutDocumentStream } from "../src/layout/engine.js";
import { ApproxMeasurer } from "../src/layout/measure.js";
import { openStreamingDocument } from "../src/stream/body.js";
import { bytesSource } from "../src/stream/source.js";
import { makeDocx, wrapDocument, W_NS } from "./helpers.js";

const measurer = new ApproxMeasurer();

/** A minimal FieldContext; individual tests override the members they exercise. */
function ctx(over: Partial<FieldContext> = {}): FieldContext {
  return {
    pageNumber: () => 1,
    totalPages: () => 2,
    formatPageNumber: (n) => String(n),
    ...over,
  };
}

/** A legacy complex field: begin / instrText / separate / cached result / end. */
function field(instr: string, cached: string): string {
  return (
    `<w:r><w:fldChar w:fldCharType="begin"/></w:r>` +
    `<w:r><w:instrText xml:space="preserve"> ${instr} </w:instrText></w:r>` +
    `<w:r><w:fldChar w:fldCharType="separate"/></w:r>` +
    `<w:r><w:t xml:space="preserve">${cached}</w:t></w:r>` +
    `<w:r><w:fldChar w:fldCharType="end"/></w:r>`
  );
}

function nestedField(before: string, nested: string, after: string, cached: string): string {
  return (
    `<w:r><w:fldChar w:fldCharType="begin"/></w:r>` +
    `<w:r><w:instrText xml:space="preserve">${before}</w:instrText></w:r>` +
    nested +
    `<w:r><w:instrText xml:space="preserve">${after}</w:instrText></w:r>` +
    `<w:r><w:fldChar w:fldCharType="separate"/></w:r>` +
    `<w:r><w:t xml:space="preserve">${cached}</w:t></w:r>` +
    `<w:r><w:fldChar w:fldCharType="end"/></w:r>`
  );
}

function bodyText(body: string): string {
  const doc = DocxDocument.load(makeDocx({ "word/document.xml": wrapDocument(body) }));
  const result = layoutDocument(doc, { measurer });
  return result.pages
    .flatMap((pg) => pg.items.filter((i) => i.kind === "text").map((i) => (i.kind === "text" ? i.text : "")))
    .join("");
}

describe("nested complex fields", () => {
  it.each(["1", "9"])("hides SET and evaluates its nested SEQ with cache %s", (cached) => {
    const set = nestedField("SET x ", field("SEQ table", cached), "", cached);
    expect(bodyText(`<w:p>${set}${field("SEQ table", "9")}</w:p>`)).toBe("2");
  });

  it("REF reads the value assigned by SET", () => {
    const set = nestedField("SET x ", field("SEQ table", "9"), "", "9");
    expect(bodyText(`<w:p>${set}${field("REF x", "stale")}${field("SEQ table", "9")}</w:p>`)).toBe("12");
  });

  it("prints the stored IF result with a nested PAGE in its instruction", () => {
    const conditional = nestedField("IF ", field("PAGE", "9"), ' = 1 "first" "other"', "first");
    expect(bodyText(`<w:p>${conditional}</w:p>`)).toBe("first");
  });

  it("recomputes the outer REF from its evaluated instruction", () => {
    const assignments = field('SET target "value"', "stale") + field('SET name "target"', "stale");
    const reference = nestedField("REF ", field("REF name", "stale"), "", "stale");
    expect(bodyText(`<w:p>${assignments}${reference}</w:p>`)).toBe("value");
  });

  it("keeps nested results in a hyperlink instruction out of body text", () => {
    const link = nestedField('HYPERLINK "https://example.com/', field("SEQ table", "9"), '"', "link");
    expect(bodyText(`<w:p>${link}${field("SEQ table", "9")}</w:p>`)).toBe("link2");
  });

  it("builds the hyperlink target from the evaluated nested field", () => {
    const link = nestedField('HYPERLINK "https://example.com/', field("SEQ table", "9"), '"', "link");
    const doc = DocxDocument.load(makeDocx({ "word/document.xml": wrapDocument(`<w:p>${link}</w:p>`) }));
    const items = layoutDocument(doc, { measurer }).pages.flatMap((page) => page.items);
    expect(items.filter((item) => item.kind === "text").map((item) => item.href)).toEqual(["https://example.com/1"]);
  });

  it("preserves instruction text on both sides of a nested result", () => {
    const set = nestedField('SET x "table ', field("SEQ table", "9"), ' caption"', "stale");
    expect(bodyText(`<w:p>${set}${field("REF x", "stale")}</w:p>`)).toBe("table 1 caption");
  });

  it("keeps a nested hyperlink result inside the outer instruction", () => {
    const set = nestedField('SET x "', field('HYPERLINK "https://example.com"', "value"), '"', "stale");
    expect(bodyText(`<w:p>${set}${field("REF x", "stale")}</w:p>`)).toBe("value");
  });

  it("handles field markers and text that share a run", () => {
    const set = nestedField("SET x ", field("SEQ table", "9"), "", "9");
    const runs = (set + field("SEQ table", "9")).replaceAll("</w:r><w:r>", "");
    expect(bodyText(`<w:p>${runs}</w:p>`)).toBe("2");
  });

  it.each([false, true])("preserves text after a nested field in a DATE result (shared run: %s)", (sharedRun) => {
    const date = 'DATE \\@ "d/M/yyyy"';
    let result = field(date, "").replace(
      '<w:r><w:t xml:space="preserve"></w:t></w:r>',
      field("NUMPAGES", "2") + '<w:r><w:t>romeo eight ten oscar quebec</w:t></w:r>',
    );
    if (sharedRun) {
      result = result.replace(
        '<w:fldChar w:fldCharType="separate"/></w:r><w:r><w:fldChar w:fldCharType="begin"/>',
        '<w:fldChar w:fldCharType="separate"/><w:fldChar w:fldCharType="begin"/>',
      );
    }
    const currentDate = bodyText(`<w:p>${field(date, "")}</w:p>`);
    expect(bodyText(`<w:p>${result}</w:p>`)).toBe(`${currentDate}1romeo eight ten oscar quebec`);
  });

  it("keeps nested SEQ assignments stable across streamed layout passes", async () => {
    const set = nestedField("SET x ", field("SEQ table", "9"), "", "9");
    const body = `<w:p>${set}${field("REF x", "stale")}${field("SEQ table", "9")}</w:p>`;
    const opened = await openStreamingDocument(bytesSource(makeDocx({ "word/document.xml": wrapDocument(body) })));
    for (let pass = 0; pass < 2; pass++) {
      let text = "";
      await layoutDocumentStream(opened.doc, {
        measurer,
        prescan: opened.prescan,
        onPage: (page) => {
          text += page.items.map((item) => item.kind === "text" ? item.text : "").join("");
        },
      });
      expect(text).toBe("12");
    }
  });

  it("handles multiple levels of nesting and instruction text after them", () => {
    const quote = nestedField("QUOTE ", field("SEQ table", "9"), "", "value");
    const set = nestedField("SET x ", quote, "", "value");
    expect(bodyText(`<w:p>${set}${field("REF x", "stale")}${field("SEQ table", "9")}</w:p>`)).toBe("value2");
  });

  it("keeps plain complex and simple fields unchanged", () => {
    const simple = '<w:fldSimple w:instr="SEQ table"><w:r><w:t>9</w:t></w:r></w:fldSimple>';
    expect(bodyText(`<w:p>${field("SEQ table", "9")}${simple}${field("PAGE", "9")}${field("UNKNOWN", "cached")}</w:p>`)).toBe("121cached");
  });

  it("hides a simple SET and exposes its value through REF", () => {
    const set = '<w:fldSimple w:instr="SET x &quot;value&quot;"><w:r><w:t>stale</w:t></w:r></w:fldSimple>';
    expect(bodyText(`<w:p>${set}${field("REF x", "stale")}</w:p>`)).toBe("value");
  });

  it("hides a SET result containing a nested field and trailing text", () => {
    const set = field('SET x "value"', "stale").replace(
      '<w:r><w:t xml:space="preserve">stale</w:t></w:r>',
      field("NUMPAGES", "2") + '<w:r><w:t>hidden</w:t></w:r>',
    );
    expect(bodyText(`<w:p>${set}${field("REF x", "stale")}</w:p>`)).toBe("value");
  });
});

describe("SEQ switches", () => {
  it.each(["\\r0", "\\r 0"])("%s resets the counter to zero", (reset) => {
    const instructions = ["SEQ Figure", "SEQ Figure", `SEQ Figure ${reset}`, "SEQ Figure"];
    expect(bodyText(instructions.map((instr) => `<w:p>${field(instr, "9")}</w:p>`).join(""))).toBe("1201");
  });

  it.each([
    ["ALPHABETIC", "AB"],
    ["alphabetic", "ab"],
    ["ROMAN", "III"],
    ["roman", "iii"],
  ])("preserves the case of %s with either switch spacing", (format, expected) => {
    for (const space of ["", " "]) {
      const body = `<w:p>${field(`SEQ Figure \\*${space}${format}`, "9")}</w:p>`;
      expect(bodyText(body + body)).toBe(expected);
    }
  });

  it("\\h hides the value and advances the counter", () => {
    const instructions = ["SEQ Figure", "SEQ Figure \\h", "SEQ Figure"];
    expect(bodyText(instructions.map((instr) => `<w:p>${field(instr, "9")}</w:p>`).join(""))).toBe("13");
  });

  it("\\h \\r0 resets the counter while hiding the value", () => {
    const instructions = [
      "seq L \\*ALPHABETIC",
      "seq L \\*ALPHABETIC",
      "seq L \\h \\r0",
      "seq L \\* ALPHABETIC",
      "seq L \\*arabic",
    ];
    expect(bodyText(instructions.map((instr) => `<w:p>${field(instr, "9")}</w:p>`).join(""))).toBe("ABA2");
  });

  it.each(["\\s1", "\\s 1"])("accepts %s alongside a format switch", (restart) => {
    const body = `<w:p>${field(`SEQ Figure ${restart} \\*ALPHABETIC`, "9")}</w:p>`;
    expect(bodyText(body + body)).toBe("AB");
  });

  it("\\c repeats the current value", () => {
    const instructions = ["SEQ Figure", "SEQ Figure \\c", "SEQ Figure"];
    expect(bodyText(instructions.map((instr) => `<w:p>${field(instr, "9")}</w:p>`).join(""))).toBe("112");
  });
});

describe("PAGE/NUMPAGES/SECTIONPAGES general-formatting (\\*) switch", () => {
  it.each(["PAGE", "NUMPAGES", "SECTIONPAGES"])("%s preserves format case with either switch spacing", (keyword) => {
    for (const space of ["", " "]) {
      for (const [format, expected] of [["ALPHABETIC", "B"], ["alphabetic", "b"], ["ROMAN", "II"], ["roman", "ii"]]) {
        expect(resolveField(`${keyword} \\*${space}${format}`, "9", ctx({ pageNumber: () => 2 }))).toBe(expected);
      }
    }
  });

  it("PAGE \\* roman overrides the section decimal format", () => {
    // pageNumber 1, section formats decimal — the \* switch forces roman.
    expect(resolveField("PAGE \\* roman", "9", ctx({ pageNumber: () => 1 }))).toBe("i");
  });
  it("PAGE \\* ArabicDash wraps the number in hyphens", () => {
    expect(resolveField("PAGE \\* ArabicDash", "x", ctx({ pageNumber: () => 1 }))).toBe("- 1 -");
  });
  it("plain PAGE keeps the section format", () => {
    expect(resolveField("PAGE", "x", ctx({ pageNumber: () => 3, formatPageNumber: (n) => `p${n}` }))).toBe("p3");
  });
  it("NUMPAGES \\* roman formats the total", () => {
    expect(resolveField("NUMPAGES \\* roman", "9", ctx({ totalPages: () => 2 }))).toBe("ii");
  });
  it("NUMPAGES with a non-numeric switch keeps the plain total", () => {
    expect(resolveField("NUMPAGES \\* MERGEFORMAT", "9", ctx({ totalPages: () => 4 }))).toBe("4");
  });
});

describe("REF recompute", () => {
  it("REF \\h to an empty (captured) bookmark shows nothing, not the stale cache", () => {
    expect(resolveField("REF bk \\h", "Table 1", ctx({ refText: () => "" }))).toBe("");
  });
  it("REF \\h to a known bookmark re-renders its text over the cache", () => {
    expect(resolveField("REF bk \\h", "STALE", ctx({ refText: () => "Introduction" }))).toBe("Introduction");
  });
  it("REF to an uncaptured bookmark keeps its cache", () => {
    expect(resolveField("REF bk \\h", "cached", ctx({ refText: () => undefined }))).toBe("cached");
  });
  it("REF \\p paints the relative position, not the bookmark text", () => {
    const key = {};
    expect(
      resolveField("REF bk \\h \\p", "Table 1 below", ctx({ refPosition: () => "above" }), key),
    ).toBe("above");
  });
  it("REF \\r paints the paragraph number", () => {
    const key = {};
    expect(resolveField("REF bk \\r", "1", ctx({ refParaNumber: () => "0" }), key)).toBe("0");
  });
  it("a \\p ref with no recorded position keeps its cache", () => {
    expect(resolveField("REF bk \\p", "below", ctx({ refPosition: () => undefined }), {})).toBe("below");
  });
});

describe("STYLEREF recompute", () => {
  it("plain STYLEREF resolves the first-on-page style text over the cache", () => {
    expect(
      resolveField('STYLEREF "Heading 1" \\* MERGEFORMAT', "Chapter One", ctx({ styleRef: (n, last) => (n === "Heading 1" && !last ? "Chapter Four" : undefined) })),
    ).toBe("Chapter Four");
  });
  it("STYLEREF \\l asks for the last-on-page paragraph", () => {
    expect(
      resolveField('STYLEREF "Heading 1" \\l \\* MERGEFORMAT', "Chapter One", ctx({ styleRef: (_n, last) => (last ? "Chapter Five" : "Chapter Four") })),
    ).toBe("Chapter Five");
  });
  it("STYLEREF keeps the cache when the resolver returns undefined (no match)", () => {
    expect(resolveField('STYLEREF "Heading 1"', "cached", ctx({ styleRef: () => undefined }))).toBe("cached");
  });
  it("STYLEREF with no resolver (body field) keeps the cache", () => {
    expect(resolveField('STYLEREF "Heading 1"', "cached", ctx())).toBe("cached");
  });
});

describe("index-xrefs style REF fields end-to-end", () => {
  // A zero-length bookmark (start immediately followed by end) sitting before
  // an unnumbered caption, referenced later by \h / \r / \p. Word recomputes
  // all three: empty text, paragraph number 0, position "above".
  const body =
    `<w:p><w:bookmarkStart w:id="1" w:name="bk1"/><w:bookmarkEnd w:id="1"/>` +
    `<w:r><w:t xml:space="preserve">Caption</w:t></w:r></w:p>` +
    `<w:p>` +
    `<w:r><w:t xml:space="preserve">[</w:t></w:r>` +
    field("REF bk1 \\h", "STALEREF") +
    `<w:r><w:t xml:space="preserve">|</w:t></w:r>` +
    field("REF bk1 \\r", "7") +
    `<w:r><w:t xml:space="preserve">|</w:t></w:r>` +
    field("REF bk1 \\p", "below") +
    `<w:r><w:t xml:space="preserve">]</w:t></w:r>` +
    `</w:p>`;

  it("empty \\h => '', \\r => '0', \\p => 'above'", () => {
    expect(bodyText(body)).toBe("Caption[|0|above]");
  });
});

describe("checkbox ballot glyph routing", () => {
  it("routes ballot glyphs to MS Gothic via paintFamily (Latin metrics kept)", () => {
    const body =
      `<w:p><w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/></w:rPr>` +
      `<w:t xml:space="preserve">A☒B</w:t></w:r></w:p>`;
    const doc = DocxDocument.load(makeDocx({ "word/document.xml": wrapDocument(body) }));
    const result = layoutDocument(doc, { measurer });
    const items = result.pages[0].items.filter((i) => i.kind === "text");
    const ballot = items.find((i) => i.kind === "text" && i.text.includes("☒"));
    const latin = items.find((i) => i.kind === "text" && i.text === "A");
    expect(ballot && ballot.kind === "text" ? ballot.font.paintFamily : undefined).toBe("MS Gothic");
    // The Latin neighbour is NOT rerouted.
    expect(latin && latin.kind === "text" ? latin.font.paintFamily : undefined).toBeUndefined();
    // Metrics stay on the run's own family so the line box is not inflated.
    expect(ballot && ballot.kind === "text" ? ballot.font.family : "").toMatch(/calibri/i);
  });
});

describe("DATE/TIME re-evaluation (Word recomputes on open; cache is stale)", () => {
  it("DATE \\@ ISO picture renders the CURRENT date, not the cache", () => {
    // Inject a fixed clock so the expectation and resolveField evaluate the
    // same instant (avoids a midnight race between two new Date() calls).
    const now = new Date();
    // formatDatePicture uses the machine's local time, as Word does.
    const iso = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}-${String(now.getDate()).padStart(2, "0")}`;
    expect(resolveField('DATE \\@ "yyyy-MM-dd"', "1999-01-01", ctx({ now }))).toBe(iso);
  });
  it.each([[2026, "26"], [2006, "06"]])("DATE single-y picture uses two digits for %i", (year, expected) => {
    const now = new Date(year, 9, 20, 12);
    expect(resolveField('DATE \\@ "MMMM d, y"', "October 20, y", ctx({ now }))).toBe(`October 20, ${expected}`);
    expect(resolveField("DATE \\@ \"'y' y yy yyyy\"", "", ctx({ now }))).toBe(`y ${expected} ${expected} ${year}`);
  });
  it("TIME am/pm token renders the designator UPPERCASE like Word", () => {
    expect(resolveField('TIME \\@ "h:mm am/pm"', "3:07 pm", ctx())).toMatch(/^\d{1,2}:\d{2} [AP]M$/);
  });
  it("month (M) and minute (m) stay case-distinguished in mixed pictures", () => {
    const now = new Date();
    const out = resolveField('DATE \\@ "M/d/yy H:mm"', "x", ctx({ now }));
    expect(out.startsWith(`${now.getMonth() + 1}/${now.getDate()}/`)).toBe(true);
    expect(out).toMatch(/ \d{1,2}:\d{2}$/);
  });
  it("CREATEDATE keeps its stored cache (a document moment, not the clock)", () => {
    expect(resolveField("CREATEDATE", "May 5, 2020", ctx())).toBe("May 5, 2020");
  });
});
