import { describe, expect, it } from "vitest";
import { zipSync, strToU8 } from "fflate";
import { DocxDocument, serializeXml, localName, type Paragraph, type Run, type XmlElement } from "@wordinweb/core";
import { DocumentSession } from "../src/session.js";

/**
 * Tool-depth wave, lane B: the header/footer page-variant toggles and the
 * page-number format op (registered operations), and the comment lifecycle
 * intents (resolveComment / editComment), through the canonical wire apply.
 */

function makeDoc(text = "body"): DocxDocument {
  const body = `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>`;
  const documentXml = `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}</w:body></w:document>`;
  return DocxDocument.load(zipSync({
    "[Content_Types].xml": strToU8(`<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`),
    "_rels/.rels": strToU8(`<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`),
    "word/document.xml": strToU8(documentXml),
  }));
}

function firstRunId(s: DocumentSession): number {
  const para = s.doc.sections[0].blocks[0] as Paragraph;
  const run = para.children[0] as Run;
  return s.ids.idOf(run.src!)!;
}

const seq = (n: number, base = 700) => Array.from({ length: n }, (_, i) => base + i);

function allRootsXml(doc: DocxDocument): string {
  return doc.editableRoots().map((r) => serializeXml(r)).join("\n----\n");
}

function bodySectPr(doc: DocxDocument): XmlElement {
  const found: XmlElement[] = [];
  const walk = (e: XmlElement): void => {
    if (localName(e.name) === "sectPr") found.push(e);
    else for (const c of e.children) walk(c);
  };
  walk(doc.docRoot);
  return found[found.length - 1];
}

describe("setTitlePage / setEvenOddHeaders on the wire", () => {
  it("enables titlePg, creates the first-page parts on carried ids, and honest no-ops when already on", () => {
    const s = new DocumentSession(makeDoc());
    const e = s.submit({ kind: "setTitlePage", clientId: "a", clientSeq: 1, base: 0, enabled: true, nodeIds: seq(4) });
    expect(e.kind).toBe("applied");
    const sectPr = bodySectPr(s.doc);
    expect(sectPr.children.some((c) => localName(c.name) === "titlePg")).toBe(true);
    expect(s.doc.sections[0].props.headerRefs.first).toBeTruthy();
    expect(s.doc.sections[0].props.footerRefs.first).toBeTruthy();
    // The created header paragraph + run took the carried ids, in doc order.
    expect(s.ids.elOf(700)).toBeTruthy();
    expect(localName(s.ids.elOf(700)!.name)).toBe("p");
    // Already enabled: a clean rejection everywhere.
    expect(s.submit({ kind: "setTitlePage", clientId: "a", clientSeq: 2, base: s.seq, enabled: true, nodeIds: seq(4, 800) }).kind).toBe("rejected");
    // Disable keeps parts, drops the toggle.
    expect(s.submit({ kind: "setTitlePage", clientId: "a", clientSeq: 3, base: s.seq, enabled: false, nodeIds: seq(4, 900) }).kind).toBe("applied");
    expect(bodySectPr(s.doc).children.some((c) => localName(c.name) === "titlePg")).toBe(false);
    expect(s.doc.sections[0].props.headerRefs.first).toBeTruthy();
  });

  it("two sessions converge byte-identically on the whole package", () => {
    const build = () => {
      const s = new DocumentSession(makeDoc());
      s.submit({ kind: "setTitlePage", clientId: "a", clientSeq: 1, base: 0, enabled: true, nodeIds: seq(4) });
      s.submit({ kind: "setEvenOddHeaders", clientId: "a", clientSeq: 2, base: s.seq, enabled: true, nodeIds: seq(4, 800) });
      return allRootsXml(s.doc);
    };
    expect(build()).toBe(build());
  });

  it("enables w:evenAndOddHeaders in settings.xml and creates the even parts", () => {
    const s = new DocumentSession(makeDoc());
    const e = s.submit({ kind: "setEvenOddHeaders", clientId: "a", clientSeq: 1, base: 0, enabled: true, nodeIds: seq(4) });
    expect(e.kind).toBe("applied");
    expect(s.doc.evenAndOddHeaders).toBe(true);
    expect(s.doc.sections[0].props.headerRefs.even).toBeTruthy();
    expect(s.doc.sections[0].props.footerRefs.even).toBeTruthy();
    expect(s.submit({ kind: "setEvenOddHeaders", clientId: "a", clientSeq: 2, base: s.seq, enabled: true, nodeIds: seq(4, 800) }).kind).toBe("rejected");
    expect(s.submit({ kind: "setEvenOddHeaders", clientId: "a", clientSeq: 3, base: s.seq, enabled: false, nodeIds: seq(4, 900) }).kind).toBe("applied");
    expect(s.doc.evenAndOddHeaders).toBe(false);
  });

  it("rejects a malformed toggle payload", () => {
    const s = new DocumentSession(makeDoc());
    expect(s.submit({ kind: "setTitlePage", clientId: "a", clientSeq: 1, base: 0, enabled: "yes" as never, nodeIds: seq(4) }).kind).toBe("rejected");
  });
});

describe("setPageNumberFormat on the wire", () => {
  it("writes w:pgNumType fmt/start; a no-change patch rejects", () => {
    const s = new DocumentSession(makeDoc());
    const e = s.submit({ kind: "setPageNumberFormat", clientId: "a", clientSeq: 1, base: 0, fmt: "lowerRoman", start: 1 });
    expect(e.kind).toBe("applied");
    expect(s.doc.sections[0].props.pageNumberFormat).toBe("lowerRoman");
    expect(s.doc.sections[0].props.pageNumberStart).toBe(1);
    expect(s.submit({ kind: "setPageNumberFormat", clientId: "a", clientSeq: 2, base: s.seq, fmt: "lowerRoman" }).kind).toBe("rejected");
    expect(s.submit({ kind: "setPageNumberFormat", clientId: "a", clientSeq: 3, base: s.seq, fmt: null, start: null }).kind).toBe("applied");
    expect(bodySectPr(s.doc).children.some((c) => localName(c.name) === "pgNumType")).toBe(false);
  });

  it("rejects malformed payloads before sequencing", () => {
    const s = new DocumentSession(makeDoc());
    expect(s.submit({ kind: "setPageNumberFormat", clientId: "a", clientSeq: 1, base: 0 } as never).kind).toBe("rejected");
    expect(s.submit({ kind: "setPageNumberFormat", clientId: "a", clientSeq: 2, base: 0, fmt: "ordinal" as never }).kind).toBe("rejected");
    expect(s.submit({ kind: "setPageNumberFormat", clientId: "a", clientSeq: 3, base: 0, start: -1 }).kind).toBe("rejected");
  });
});

describe("resolveComment / editComment on the wire", () => {
  const withComment = (): { s: DocumentSession; id: string } => {
    const s = new DocumentSession(makeDoc("commented text"));
    expect(s.submit({ kind: "commentRun", clientId: "a", clientSeq: 1, base: 0, runId: firstRunId(s), text: "Note", author: "Reviewer", date: "2026-01-01T00:00:00Z", paraId: "11111111" }).kind).toBe("applied");
    return { s, id: s.doc.comments[0].id };
  };

  it("resolves and reopens a thread; the resolved flag derives from w15:done", () => {
    const { s, id } = withComment();
    expect(s.submit({ kind: "resolveComment", clientId: "b", clientSeq: 1, base: s.seq, commentId: id, resolved: true, paraId: "22222222" }).kind).toBe("applied");
    expect(s.doc.comments[0].resolved).toBe(true);
    // Same state again: honest no-op, identical rejection everywhere.
    expect(s.submit({ kind: "resolveComment", clientId: "b", clientSeq: 2, base: s.seq, commentId: id, resolved: true, paraId: "33333333" }).kind).toBe("rejected");
    expect(s.submit({ kind: "resolveComment", clientId: "b", clientSeq: 3, base: s.seq, commentId: id, resolved: false, paraId: "44444444" }).kind).toBe("applied");
    expect(s.doc.comments[0].resolved).toBeUndefined();
    expect(s.submit({ kind: "resolveComment", clientId: "b", clientSeq: 4, base: s.seq, commentId: "ghost", resolved: true, paraId: "55555555" }).kind).toBe("rejected");
  });

  it("two sessions resolving with the same carried paraId converge byte-identically", () => {
    const build = () => {
      const { s, id } = withComment();
      s.submit({ kind: "resolveComment", clientId: "b", clientSeq: 1, base: s.seq, commentId: id, resolved: true, paraId: "22222222" });
      return allRootsXml(s.doc);
    };
    expect(build()).toBe(build());
  });

  it("edits a comment's text; unchanged or unknown edits reject", () => {
    const { s, id } = withComment();
    expect(s.submit({ kind: "editComment", clientId: "a", clientSeq: 2, base: s.seq, commentId: id, text: "Sharper note" }).kind).toBe("applied");
    expect(s.doc.comments[0].text).toBe("Sharper note");
    expect(s.submit({ kind: "editComment", clientId: "a", clientSeq: 3, base: s.seq, commentId: id, text: "Sharper note" }).kind).toBe("rejected");
    expect(s.submit({ kind: "editComment", clientId: "a", clientSeq: 4, base: s.seq, commentId: "ghost", text: "x" }).kind).toBe("rejected");
    expect(s.submit({ kind: "editComment", clientId: "a", clientSeq: 5, base: s.seq, commentId: id, text: "" }).kind).toBe("rejected");
    expect(s.submit({ kind: "editComment", clientId: "a", clientSeq: 6, base: s.seq, commentId: id, text: "x".repeat(20_001) }).kind).toBe("rejected");
  });
});
