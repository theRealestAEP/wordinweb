import { describe, expect, it } from "vitest";
import { DocxDocument } from "../src/docx.js";
import { TOC_EMPTY_TEXT, findTocFields, insertToc, rebuildToc } from "../src/edit/toc.js";
import { updateFields } from "../src/edit/update-fields.js";
import { layoutDocument } from "../src/layout/engine.js";
import { ApproxMeasurer } from "../src/layout/measure.js";
import { XmlElement, localName, serializeXml } from "../src/xml.js";
import { makeDocx, wrapDocument } from "./helpers.js";

const measurer = new ApproxMeasurer();

const heading = (level: number, text: string) =>
  `<w:p><w:pPr><w:pStyle w:val="Heading${level}"/></w:pPr>` +
  `<w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;

const anchor = `<w:p><w:r><w:t xml:space="preserve">TOC goes here</w:t></w:r></w:p>`;

/** Heading1-3 carry the outline levels a TOC selects on. */
const STYLES_XML =
  `<?xml version="1.0"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
  [1, 2, 3]
    .map(
      (n) =>
        `<w:style w:type="paragraph" w:styleId="Heading${n}"><w:name w:val="heading ${n}"/>` +
        `<w:pPr><w:outlineLvl w:val="${n - 1}"/></w:pPr></w:style>`,
    )
    .join("") +
  `</w:styles>`;

function load(body: string): DocxDocument {
  return DocxDocument.load(
    makeDocx({ "word/document.xml": wrapDocument(body), "word/styles.xml": STYLES_XML }),
  );
}

/** Insert a TOC at the anchor paragraph, then fill in real page numbers the
 * way a host does: lay the document out and run the update pass. */
function withToc(body: string, options = {}): DocxDocument {
  const doc = load(body);
  const caretT = firstTextIn(doc.docRoot)!;
  expect(insertToc(doc, caretT, options)).toBe(true);
  updateFields(doc, { layout: layoutDocument(doc, { measurer }) });
  return doc;
}

function firstTextIn(el: XmlElement): XmlElement | undefined {
  if (localName(el.name) === "t") return el;
  for (const c of el.children) {
    const found = firstTextIn(c);
    if (found) return found;
  }
  return undefined;
}

function xml(doc: DocxDocument): string {
  return serializeXml(doc.docRoot);
}

/** The paragraphs of the document body, as XML strings. */
function bodyParagraphs(doc: DocxDocument): string[] {
  const body = doc.docRoot.children.find((c) => localName(c.name) === "body")!;
  return body.children.filter((c) => localName(c.name) === "p").map((p) => serializeXml(p));
}

describe("insertToc: the field structure Word requires", () => {
  const body = anchor + heading(1, "Introduction") + heading(2, "Background") + heading(1, "Method");

  it("opens the field in the FIRST entry paragraph: begin, instrText, separate", () => {
    const paras = bodyParagraphs(withToc(body));
    // paras[0] is the anchor; paras[1] is the first entry and carries the field.
    expect(paras[1]).toContain(`<w:fldChar w:fldCharType="begin"/>`);
    expect(paras[1]).toContain(`<w:instrText xml:space="preserve"> TOC \\o "1-3" \\h \\z \\u </w:instrText>`);
    expect(paras[1]).toContain(`<w:fldChar w:fldCharType="separate"/>`);
    // In that order, and all three before the entry's hyperlink.
    const begin = paras[1].indexOf(`fldCharType="begin"`);
    const instr = paras[1].indexOf(`<w:instrText`);
    const sep = paras[1].indexOf(`fldCharType="separate"`);
    const link = paras[1].indexOf(`<w:hyperlink`);
    expect(begin).toBeLessThan(instr);
    expect(instr).toBeLessThan(sep);
    expect(sep).toBeLessThan(link);
  });

  it("closes the field with an end fldChar in a paragraph of its own", () => {
    const paras = bodyParagraphs(withToc(body));
    const closing = paras.filter((p) => p.includes(`fldCharType="end"`) && !p.includes("PAGEREF"));
    expect(closing).toEqual([`<w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>`]);
    // It sits directly after the last entry, ahead of the source headings.
    const lastEntry = paras.map((p) => p.includes("<w:hyperlink")).lastIndexOf(true);
    expect(paras[lastEntry + 1]).toBe(closing[0]);
  });

  it("has exactly one TOC begin and one matching end", () => {
    const doc = withToc(body);
    const all = xml(doc);
    expect(findTocFields(doc)).toHaveLength(1);
    // Three PAGEREF fields plus the TOC field itself: four begins, four ends.
    expect(all.match(/fldCharType="begin"/g)).toHaveLength(4);
    expect(all.match(/fldCharType="end"/g)).toHaveLength(4);
  });

  it("gives each entry a TOC style matching its outline level", () => {
    const paras = bodyParagraphs(withToc(body));
    expect(paras[1]).toContain(`<w:pStyle w:val="TOC1"/>`);
    expect(paras[2]).toContain(`<w:pStyle w:val="TOC2"/>`);
    expect(paras[3]).toContain(`<w:pStyle w:val="TOC1"/>`);
  });

  it("gives each entry a right-aligned dot-leader tab stop", () => {
    const paras = bodyParagraphs(withToc(body));
    expect(paras[1]).toMatch(/<w:tab w:val="right" w:leader="dot" w:pos="\d+"\/>/);
  });

  it("honours an explicit tab leader", () => {
    const paras = bodyParagraphs(withToc(body, { leader: "hyphen" }));
    expect(paras[1]).toContain(`w:leader="hyphen"`);
  });

  it("wraps each entry in a hyperlink to the heading's _Toc bookmark", () => {
    const paras = bodyParagraphs(withToc(body));
    const m = /<w:hyperlink w:anchor="(_Toc\d+)" w:history="1">/.exec(paras[1]);
    expect(m).not.toBeNull();
    // The same bookmark names the entry's PAGEREF target.
    expect(paras[1]).toContain(`<w:instrText xml:space="preserve"> PAGEREF ${m![1]} \\h </w:instrText>`);
  });

  it("lays an entry out as title, tab, then a PAGEREF field", () => {
    const paras = bodyParagraphs(withToc(body));
    const entry = paras[2];
    expect(entry.indexOf("Background")).toBeLessThan(entry.indexOf("<w:tab/>"));
    expect(entry.indexOf("<w:tab/>")).toBeLessThan(entry.indexOf("PAGEREF"));
    // The PAGEREF is a complete complex field.
    expect(entry).toMatch(
      /fldCharType="begin".*<w:instrText[^>]*> PAGEREF _Toc\d+ \\h <\/w:instrText>.*fldCharType="separate".*<w:t[^>]*>\d+<\/w:t>.*fldCharType="end"/s,
    );
  });

  it("marks generated runs noProof, and the leader and page number webHidden", () => {
    const paras = bodyParagraphs(withToc(body));
    expect(paras[1]).toContain(`<w:rStyle w:val="Hyperlink"/><w:noProof/>`);
    expect(paras[1]).toContain(`<w:noProof/><w:webHidden/></w:rPr><w:tab/>`);
  });

  it("bookmarks each heading so Ctrl-click navigation lands on it", () => {
    const paras = bodyParagraphs(withToc(body));
    const introduction = paras.find((p) => p.includes("Introduction") && p.includes("Heading1"))!;
    expect(introduction).toMatch(/<w:bookmarkStart w:id="\d+" w:name="_Toc\d+"\/><w:r>/);
    expect(introduction).toMatch(/<w:bookmarkEnd w:id="\d+"\/><\/w:p>$/);
  });

  it("declares the TOC paragraph styles it references", () => {
    const doc = withToc(body);
    expect(doc.styles.byId.get("TOC1")?.name).toBe("toc 1");
    expect(doc.styles.byId.get("TOC2")?.name).toBe("toc 2");
  });
});

describe("insertToc: level selection", () => {
  const body = anchor + heading(1, "One") + heading(2, "Two") + heading(3, "Three");

  it("includes levels 1-3 by default", () => {
    const paras = bodyParagraphs(withToc(body));
    expect(paras.filter((p) => p.includes("<w:hyperlink"))).toHaveLength(3);
  });

  it("\\o \"1-1\" takes only the top level", () => {
    const doc = withToc(body, { levels: [1, 1] });
    const paras = bodyParagraphs(doc);
    expect(paras.filter((p) => p.includes("<w:hyperlink"))).toHaveLength(1);
    expect(xml(doc)).toContain(`> TOC \\o "1-1" \\h \\z \\u <`);
  });

  it("writes Word's placeholder when nothing qualifies", () => {
    const doc = withToc(anchor + `<w:p><w:r><w:t xml:space="preserve">Body</w:t></w:r></w:p>`);
    expect(xml(doc)).toContain(TOC_EMPTY_TEXT);
    // Still a real field, so Word's Update Table can fill it in later.
    expect(findTocFields(doc)).toHaveLength(1);
  });

  it("rejects a level range outside 1-9", () => {
    const doc = load(anchor + heading(1, "One"));
    expect(insertToc(doc, firstTextIn(doc.docRoot)!, { levels: [0, 3] })).toBe(false);
    expect(insertToc(doc, firstTextIn(doc.docRoot)!, { levels: [3, 1] })).toBe(false);
    expect(insertToc(doc, firstTextIn(doc.docRoot)!, { levels: [1, 12] })).toBe(false);
  });
});

describe("TOC page numbers come from the update pass", () => {
  const pageBreak = `<w:p><w:r><w:br w:type="page"/></w:r></w:p>`;
  const body = anchor + heading(1, "First") + pageBreak + heading(1, "Second");

  it("each entry's PAGEREF caches the page its heading landed on", () => {
    const doc = withToc(body);
    const pages = [...xml(doc).matchAll(/<w:webHidden\/><\/w:rPr><w:t xml:space="preserve">(\d+)<\/w:t>/g)].map(
      (m) => m[1],
    );
    expect(pages).toEqual(["1", "2"]);
  });

  it("repages after content moves a heading, with no TOC-specific code", () => {
    const doc = withToc(body);
    // Push the second heading onto a third page with another break. The TOC's
    // own entry for it comes first in document order, so match the heading by
    // its style, not just its text.
    const bodyEl = doc.docRoot.children.find((c) => localName(c.name) === "body")!;
    const secondHeading = bodyEl.children.findIndex(
      (c) => serializeXml(c).includes("Second") && serializeXml(c).includes("Heading1"),
    );
    bodyEl.children.splice(secondHeading, 0, {
      name: "w:p",
      attrs: {},
      children: [{ name: "w:r", attrs: {}, children: [{ name: "w:br", attrs: { "w:type": "page" }, children: [], text: "" }], text: "" }],
      text: "",
    });
    doc.refresh();
    expect(updateFields(doc, { layout: layoutDocument(doc, { measurer }) })).toBe(true);
    const pages = [...xml(doc).matchAll(/<w:webHidden\/><\/w:rPr><w:t xml:space="preserve">(\d+)<\/w:t>/g)].map(
      (m) => m[1],
    );
    expect(pages).toEqual(["1", "3"]);
  });
});

describe("rebuildToc", () => {
  const body = anchor + heading(1, "First") + heading(1, "Second");

  it("picks up a heading added after the table was built", () => {
    const doc = withToc(body);
    const bodyEl = doc.docRoot.children.find((c) => localName(c.name) === "body")!;
    bodyEl.children.push(...load(heading(1, "Third")).docRoot.children[0].children);
    doc.refresh();

    expect(rebuildToc(doc, findTocFields(doc)[0])).toBe(true);
    const paras = bodyParagraphs(doc);
    const entries = paras.filter((p) => p.includes("<w:hyperlink"));
    expect(entries).toHaveLength(3);
    expect(entries[2]).toContain("Third");
  });

  it("keeps the bookmark a heading already had, so existing links still resolve", () => {
    const doc = withToc(body);
    const before = [...xml(doc).matchAll(/w:name="(_Toc\d+)"/g)].map((m) => m[1]);
    rebuildToc(doc, findTocFields(doc)[0]);
    const after = [...xml(doc).matchAll(/w:name="(_Toc\d+)"/g)].map((m) => m[1]);
    expect(after).toEqual(before);
  });

  it("leaves exactly one field behind: the rebuild replaces, never nests", () => {
    const doc = withToc(body);
    rebuildToc(doc, findTocFields(doc)[0]);
    rebuildToc(doc, findTocFields(doc)[0]);
    expect(findTocFields(doc)).toHaveLength(1);
    const all = xml(doc);
    expect(all.match(/fldCharType="begin"/g)).toHaveLength(3); // TOC + two PAGEREFs
    expect(all.match(/fldCharType="end"/g)).toHaveLength(3);
  });
});

describe("a generated TOC renders", () => {
  it("paints every entry's title and page number", () => {
    const doc = withToc(anchor + heading(1, "Introduction") + heading(2, "Background"));
    const text = layoutDocument(doc, { measurer })
      .pages.flatMap((p) => p.items.filter((i) => i.kind === "text").map((i) => (i.kind === "text" ? i.text : "")))
      .join("");
    expect(text).toContain("Introduction");
    expect(text).toContain("Background");
    // The instruction itself never paints.
    expect(text).not.toContain("PAGEREF");
    expect(text).not.toContain("TOC \\o");
  });
});
