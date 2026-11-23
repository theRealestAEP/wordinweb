import { describe, expect, it } from "vitest";
import { DocxDocument } from "../src/docx.js";
import { readWordDate, wordDate } from "../src/word-date.js";
import { makeDocx, p, wrapDocument } from "./helpers.js";

// Word 16 on macOS in California at 11:01 local wrote w:date="2026-11-10T11:01:00Z"
// in comments.xml and w16cex:dateUtc="2026-11-10T19:01:00Z" in commentsExtensible.xml.
describe("Word comment and revision dates", () => {
  it("writes local wall-clock time with Z and whole seconds, as Word does", () => {
    expect(wordDate(new Date(2026, 10, 10, 11, 1, 7, 450))).toBe("2026-11-10T11:01:07Z");
  });

  it("reads w:date as local time when the file has no UTC time", () => {
    const when = readWordDate("2026-11-10T11:01:00Z")!;
    expect([when.getFullYear(), when.getMonth(), when.getDate(), when.getHours(), when.getMinutes()]).toEqual([2026, 10, 10, 11, 1]);
    expect(readWordDate(wordDate(new Date(2026, 0, 2, 3, 4, 5)))).toEqual(new Date(2026, 0, 2, 3, 4, 5));
  });

  it("prefers the true UTC time", () => {
    expect(readWordDate("2026-11-10T11:01:00Z", "2026-11-10T19:01:00Z")!.toISOString()).toBe("2026-11-10T19:01:00.000Z");
    expect(readWordDate(undefined)).toBeNull();
    expect(readWordDate("not a date")).toBeNull();
  });

  it("takes a comment's UTC time from commentsExtensible through commentsIds", () => {
    const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"';
    const doc = DocxDocument.load(makeDocx({
      "word/document.xml": wrapDocument(p("Comment time test.")),
      "word/comments.xml": `<w:comments ${W}><w:comment w:id="0" w:author="Ada" w:date="2026-11-10T11:01:00Z"><w:p w14:paraId="1A2B3C4D"><w:r><w:t>Test comment</w:t></w:r></w:p></w:comment></w:comments>`,
      "word/commentsIds.xml": `<w16cid:commentsIds xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid"><w16cid:commentId w16cid:paraId="1A2B3C4D" w16cid:durableId="5E6F7A8B"/></w16cid:commentsIds>`,
      "word/commentsExtensible.xml": `<w16cex:commentsExtensible xmlns:w16cex="http://schemas.microsoft.com/office/word/2018/wordml/cex"><w16cex:commentExtensible w16cex:durableId="5E6F7A8B" w16cex:dateUtc="2026-11-10T19:01:00Z"/></w16cex:commentsExtensible>`,
    }));
    expect(doc.comments[0]).toMatchObject({ date: "2026-11-10T11:01:00Z", dateUtc: "2026-11-10T19:01:00Z" });
  });
});
