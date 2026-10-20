import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { docToDocxFiles, openDoc } from "../src/doc/convert.js";
import { plainText, readDocReview, revisionsAt } from "../src/doc/review.js";
import { commentsExtendedXml, commentsXml, joinRevisionWrappers, reviewHooks } from "../src/doc/review-xml.js";

// Word 2016 "save as Word 97-2003" of three of the render eval's synthetic
// documents (comments with replies, insertions and deletions, format
// changes, a paragraph property change, inserted and deleted table rows).
const fixture = (name: string): Uint8Array => new Uint8Array(readFileSync(new URL(`./fixtures/doc/${name}.doc`, import.meta.url)));
const open = (name: string) => {
  const bytes = fixture(name);
  const reader = openDoc(bytes);
  return { bytes, reader, review: readDocReview(reader) };
};
const part = (files: Record<string, Uint8Array>, name: string): string => new TextDecoder().decode(files[name]);

describe("doc review: comments", () => {
  const { reader, review } = open("review-comments");

  it("reads every comment with its author, initials, date and text", () => {
    expect(review.comments).toHaveLength(8);
    expect(review.commentAuthors).toEqual(["Counsel C", "Editor B"]);
    const first = review.comments[1];
    expect(first.author).toBe("Editor B");
    expect(first.initials).toBe("EB");
    expect(first.date).toBe("2026-01-06T14:30:00Z");
    expect(first.text).toBe("Agreed, leaving as is.");
    expect(first.paragraphs).toHaveLength(1);
    expect(first.paragraphs[0].start).toBe(first.textStart);
    expect(first.paragraphs[0].end).toBe(first.textEnd);
    expect(first.textStart).toBeGreaterThanOrEqual(reader.parts.comments);
  });

  it("gives the commented range and the reference mark as main-document CPs", () => {
    const c = review.comments[1];
    expect(reader.pieces.text(c.rangeStart, c.rangeEnd)).toBe("razokoba Juqozuzi kipuq ");
    expect(reader.pieces.text(c.referenceCp, c.referenceCp + 1)).toBe("\u0005");
    expect(c.rangeEnd).toBe(c.referenceCp);
  });

  it("links a reply to its parent from AtrdExtra", () => {
    const reply = review.comments[2];
    expect(reply.depth).toBe(1);
    expect(reply.parentId).toBe(1);
    expect(review.comments.filter((c) => c.parentId !== undefined)).toHaveLength(4);
    expect(review.comments[0].parentId).toBeUndefined();
  });

  it("writes comments.xml and commentsExtended.xml", () => {
    const xml = commentsXml(review.comments);
    expect(xml).toContain('<w:comment w:id="1" w:author="Editor B" w:date="2026-01-06T14:30:00Z" w:initials="EB">');
    expect(xml).toContain("<w:annotationRef/>");
    expect(xml).toContain("Agreed, leaving as is.");
    const extended = commentsExtendedXml(review.comments)!;
    expect(extended).toContain('w15:paraId="10000002" w15:paraIdParent="10000001" w15:done="0"');
  });
});

describe("doc review: tracked changes", () => {
  const { reader, review } = open("review-comments");

  it("reads insertions and deletions with authors and dates", () => {
    expect(review.authors).toEqual(["Unknown", "Counsel C", "Editor B", "Reviewer A"]);
    const ins = review.revisions.filter((r) => r.kind === "ins");
    const del = review.revisions.filter((r) => r.kind === "del");
    expect(ins).toHaveLength(7);
    expect(del).toHaveLength(3);
    expect(reader.pieces.text(ins[0].start, ins[0].end)).toBe(" meridian saffron 65");
    expect(ins[0].author).toBe("Counsel C");
    expect(ins[0].date).toBe("2026-01-07T09:15:00Z");
    expect(reader.pieces.text(del[0].start, del[0].end)).toBe("- Juqozuzi siliqaj ");
    expect(del[0].author).toBe("Editor B");
  });

  it("covers paragraph marks, which revisionsAt finds", () => {
    const marks = review.revisions.filter((r) => reader.pieces.text(r.start, r.end) === "\r");
    expect(marks.map((r) => r.kind)).toEqual(["ins", "del"]);
    expect(revisionsAt(review, marks[0].start)).toEqual([marks[0]]);
    expect(revisionsAt(review, marks[0].start - 1)).toEqual([]);
  });

  it("reads a run's format change with the old properties", () => {
    const changes = review.propertyRevisions.filter((p) => p.kind === "rPrChange");
    expect(changes).toHaveLength(2);
    expect(changes[0].author).toBe("Counsel C");
    expect(changes[0].date).toBe("2026-01-06T14:30:00Z");
    expect(reader.pieces.text(changes[0].start, changes[0].end)).toBe("Har 7494 - Zevuboj hugerogus tilozuhu talece");
    // the sprms before sprmCWall: bold, size 13pt, the color
    expect(changes[0].old.length).toBeGreaterThan(0);
    const files = docToDocxFiles(fixture("review-comments"), reviewHooks(reader, review));
    const body = part(files, "word/document.xml");
    expect(body).toMatch(/<w:rPrChange w:id="\d+" w:author="Counsel C" w:date="2026-01-06T14:30:00Z"><w:rPr>(?:(?!<\/w:rPr>).)*<w:b\/>(?:(?!<\/w:rPr>).)*<w:color w:val="454042"\/>(?:(?!<\/w:rPr>).)*<w:sz w:val="26"\/>/);
  });
});

describe("doc review: the converter hooks", () => {
  it("wraps runs, marks the comment ranges and references, and adds the parts", () => {
    const { bytes, reader, review } = open("review-comments");
    const files = docToDocxFiles(bytes, reviewHooks(reader, review));
    const body = part(files, "word/document.xml");
    expect(body).toMatch(/<w:ins w:id="\d+" w:author="Counsel C" w:date="2026-01-07T09:15:00Z"><w:r>/);
    expect(body).toMatch(/<w:del w:id="\d+" w:author="Editor B" w:date="2026-01-06T14:30:00Z"><w:r><w:rPr>.*?<\/w:rPr><w:delText xml:space="preserve">- Juqozuzi siliqaj <\/w:delText><\/w:r><\/w:del>/);
    // the paragraph marks' revisions sit in the paragraph's rPr
    expect(body).toMatch(/<w:pPr>(?:(?!<\/w:pPr>).)*<w:rPr>(?:(?!<\/w:rPr>).)*<w:ins w:id="\d+" w:author="Counsel C" w:date="2026-01-05T10:00:00Z"\/>/);
    expect(body).toContain('<w:commentRangeStart w:id="1"/>');
    expect(body).toContain('<w:commentRangeEnd w:id="1"/>');
    expect(body).toMatch(/<w:commentRangeEnd w:id="1"\/><w:r>(?:<w:rPr>.*?<\/w:rPr>)?<w:commentReference w:id="1"\/><\/w:r>/);
    expect(part(files, "word/comments.xml")).toContain('<w:comment w:id="1" w:author="Editor B"');
    expect(part(files, "word/comments.xml")).toContain("<w:annotationRef/>");
    expect(part(files, "word/commentsExtended.xml")).toContain("w15:paraIdParent");
    expect(part(files, "word/_rels/document.xml.rels")).toContain('Target="comments.xml"');
    expect(part(files, "[Content_Types].xml")).toContain("/word/commentsExtended.xml");
    // every w:id of a revision is its own
    const ids = [...body.matchAll(/<w:(?:ins|del|rPrChange|pPrChange|trPrChange) w:id="(\d+)"/g)].map((m) => m[1]);
    expect(new Set(ids).size).toBe(ids.length);
  });

  it("writes a paragraph property change with the old paragraph properties", () => {
    const { bytes, reader, review } = open("review-pprchange");
    const change = review.propertyRevisions.find((p) => p.kind === "pPrChange")!;
    expect(change.author).toBe("Reviewer A");
    const body = part(docToDocxFiles(bytes, reviewHooks(reader, review)), "word/document.xml");
    expect(body).toMatch(/<w:pPrChange w:id="\d+" w:author="Reviewer A" w:date="2026-01-06T14:30:00Z"><w:pPr><w:spacing w:line="480" w:lineRule="exact"\/><\/w:pPr><\/w:pPrChange>/);
  });

  it("marks inserted and deleted table rows in their trPr", () => {
    const { bytes, reader, review } = open("review-rows");
    expect(review.rowMarks).toHaveLength(5);
    const rowRevisions = review.rowMarks.flatMap((cp) => revisionsAt(review, cp));
    expect(rowRevisions.map((r) => r.kind)).toEqual(["del", "ins"]);
    const body = part(docToDocxFiles(bytes, reviewHooks(reader, review)), "word/document.xml");
    expect(body).toMatch(/<w:trPr>(?:(?!<\/w:trPr>).)*<w:del w:id="\d+" w:author="Editor B" w:date="2026-01-06T14:30:00Z"\/>/);
    expect(body).toMatch(/<w:trPr>(?:(?!<\/w:trPr>).)*<w:ins w:id="\d+" w:author="Editor B" w:date="2026-01-05T10:00:00Z"\/>/);
  });
});

describe("doc review: joinRevisionWrappers", () => {
  const A = 'w:author="A" w:date="2026-01-05T10:00:00Z"';
  const B = 'w:author="B" w:date="2026-01-05T10:00:00Z"';
  it("joins adjacent wrappers of one author and date, level by level", () => {
    expect(joinRevisionWrappers(`<w:p><w:del w:id="1" ${A}>x</w:del><w:del w:id="2" ${A}>y</w:del></w:p>`)).toBe(`<w:p><w:del w:id="1" ${A}>xy</w:del></w:p>`);
    expect(joinRevisionWrappers(`<w:ins w:id="1" ${A}><w:del w:id="2" ${B}>x</w:del></w:ins><w:ins w:id="3" ${A}><w:del w:id="4" ${B}>y</w:del></w:ins>`)).toBe(
      `<w:ins w:id="1" ${A}><w:del w:id="2" ${B}>xy</w:del></w:ins>`,
    );
  });
  it("takes the bookmark and comment range marks between two wrappers inside", () => {
    expect(joinRevisionWrappers(`<w:del w:id="1" ${A}>x</w:del><w:bookmarkStart w:id="7" w:name="b"/><w:commentRangeEnd w:id="2"/><w:del w:id="2" ${A}>y</w:del>`)).toBe(
      `<w:del w:id="1" ${A}>x<w:bookmarkStart w:id="7" w:name="b"/><w:commentRangeEnd w:id="2"/>y</w:del>`,
    );
  });
  it("keeps wrappers apart across other content, another author or a mark's revision", () => {
    const apart = `<w:del w:id="1" ${A}>x</w:del><w:r/><w:del w:id="2" ${A}>y</w:del><w:del w:id="3" ${B}>z</w:del>`;
    expect(joinRevisionWrappers(apart)).toBe(apart);
    const mark = `<w:pPr><w:rPr><w:del w:id="1" ${A}/></w:rPr></w:pPr><w:del w:id="2" ${A}>x</w:del>`;
    expect(joinRevisionWrappers(mark)).toBe(mark);
  });
});

describe("doc review: plainText", () => {
  it("keeps field results and drops codes and marks", () => {
    expect(plainText("\u0005see \u0013 PAGE \u0014 3 \u0015 here\u0001\r")).toBe("see  3  here\n");
    expect(plainText("a\u001eb\u001fc\tx")).toBe("a‑bc\tx");
  });
});
