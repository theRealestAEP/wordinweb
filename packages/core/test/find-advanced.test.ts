import { describe, expect, it } from "vitest";
import { DocxDocument } from "../src/docx.js";
import { findAll, replaceAll } from "../src/edit/find.js";
import { serializeXml } from "../src/xml.js";
import { makeDocx, p, wrapDocument } from "./helpers.js";

function load(body: string): DocxDocument {
  return DocxDocument.load(makeDocx({ "word/document.xml": wrapDocument(body) }));
}

/** The matched text of each hit, read back through its source ranges. */
function matchTexts(doc: DocxDocument, query: string, opts?: Parameters<typeof findAll>[2]): string[] {
  return findAll(doc, query, opts).map((m) =>
    m.ranges.map((r) => r.t.text.slice(r.start, r.end)).join(""),
  );
}

describe("wildcard find (Word's pattern subset)", () => {
  const doc = () => load(p("cat cot cut coat CAT") + p("lot loot looot"));

  it("? matches any single character, * a lazy run", () => {
    expect(matchTexts(doc(), "c?t", { wildcards: true })).toEqual(["cat", "cot", "cut"]);
    expect(matchTexts(doc(), "c*t", { wildcards: true })).toEqual(["cat", "cot", "cut", "coat"]);
  });

  it("is always case-sensitive, Word's wildcard rule", () => {
    expect(matchTexts(doc(), "CAT", { wildcards: true })).toEqual(["CAT"]);
    expect(matchTexts(doc(), "CAT", { wildcards: true, matchCase: false })).toEqual(["CAT"]);
  });

  it("sets, negated sets, and ranges", () => {
    expect(matchTexts(doc(), "c[ao]t", { wildcards: true })).toEqual(["cat", "cot"]);
    expect(matchTexts(doc(), "c[!ao]t", { wildcards: true })).toEqual(["cut"]);
    expect(matchTexts(doc(), "c[a-u]t", { wildcards: true })).toEqual(["cat", "cot", "cut"]);
  });

  it("@ repeats the previous element one or more times", () => {
    expect(matchTexts(doc(), "lo@t", { wildcards: true })).toEqual(["lot", "loot", "looot"]);
    expect(matchTexts(doc(), "lo@@t", { wildcards: true })).toEqual([]); // @ with a quantifier before it: malformed
  });

  it("< and > anchor word boundaries", () => {
    const d = load(p("scatter cat concat"));
    expect(matchTexts(d, "<cat", { wildcards: true })).toEqual(["cat"]);
    expect(matchTexts(d, "cat>", { wildcards: true })).toEqual(["cat", "cat"]); // "concat" ends a word too
    expect(matchTexts(d, "<cat>", { wildcards: true })).toEqual(["cat"]);
  });

  it("\\ escapes wildcards; ^13 and ^9 name the paragraph mark and tab", () => {
    const d = load(p("really?") + p("next"));
    expect(matchTexts(d, "really\\?", { wildcards: true })).toEqual(["really?"]);
    // A match across the paragraph mark keeps the mark (null ref); the
    // covered TEXT is what the ranges hold.
    expect(matchTexts(d, "really\\?^13next", { wildcards: true })).toEqual(["really?next"]);
  });

  it("* never crosses a paragraph mark", () => {
    const d = load(p("alpha") + p("beta"));
    expect(matchTexts(d, "al*ta", { wildcards: true })).toEqual([]);
  });

  it("refuses malformed and over-large patterns with zero matches", () => {
    expect(matchTexts(doc(), "[abc", { wildcards: true })).toEqual([]); // unterminated set
    expect(matchTexts(doc(), "@x", { wildcards: true })).toEqual([]); // nothing to repeat
    expect(matchTexts(doc(), "^12", { wildcards: true })).toEqual([]); // unmodeled code
    expect(matchTexts(doc(), "a".repeat(300), { wildcards: true })).toEqual([]); // over the cap
    expect(matchTexts(doc(), "*a*a*a*a*a*a*a*a*a", { wildcards: true })).toEqual([]); // too many quantifiers
  });

  it("replaceAll works over wildcard matches", () => {
    const d = load(p("cat cot cut"));
    const result = replaceAll(d, "c[ao]t", "dog", { wildcards: true });
    expect(result.total).toBe(2);
    expect(serializeXml(d.docRoot)).toContain("dog dog cut");
  });
});

describe("special-character escapes in literal mode", () => {
  it("^p matches the paragraph mark join", () => {
    const d = load(p("one") + p("two"));
    expect(matchTexts(d, "one^ptwo")).toEqual(["onetwo"]); // the mark itself has no w:t
    expect(findAll(d, "one^ptwo").length).toBe(1);
  });

  it("^t matches a real w:tab", () => {
    const d = load(`<w:p><w:r><w:t xml:space="preserve">a</w:t><w:tab/><w:t xml:space="preserve">b</w:t></w:r></w:p>`);
    expect(findAll(d, "a^tb").length).toBe(1);
    expect(findAll(d, "a b").length).toBe(0); // the tab is not a space
  });

  it("^# ^$ ^? ^w and ^^ translate as documented", () => {
    const d = load(p("Room 42 is open ^ up"));
    expect(matchTexts(d, "Room ^#^#")).toEqual(["Room 42"]);
    expect(matchTexts(d, "^$^$^$^$ 42")).toEqual(["Room 42"]);
    expect(matchTexts(d, "4^?")).toEqual(["42"]);
    expect(matchTexts(d, "is^wopen")).toEqual(["is open"]);
    expect(matchTexts(d, "^^ up")).toEqual(["^ up"]);
  });

  it("keeps case-insensitivity and wholeWord in literal escape mode", () => {
    const d = load(p("Item 7 items 8"));
    expect(matchTexts(d, "item ^#")).toEqual(["Item 7"]); // case-insensitive; "items" fails the space
    expect(matchTexts(d, "item", { wholeWord: true })).toEqual(["Item"]);
    expect(matchTexts(d, "item^w^#", { wholeWord: true })).toEqual(["Item 7"]);
  });

  it("an unknown escape stays literal text", () => {
    const d = load(p("x^qy"));
    expect(matchTexts(d, "x^qy")).toEqual(["x^qy"]);
  });

  it("a match made only of unaddressable characters reports nothing", () => {
    const d = load(p("one") + p("two"));
    expect(findAll(d, "^p")).toEqual([]);
  });
});
