import { describe, expect, it } from "vitest";
import { zipSync, strToU8 } from "fflate";
import { DocxDocument, serializeXml, localName, type Paragraph, type Run, type XmlElement } from "@wordinweb/core";
import { DocumentSession } from "../src/session.js";
import { transformPosition } from "../src/transform.js";
import { FormatRangeIntent, InsertTextIntent, Position } from "../src/intents.js";

function makeDoc(text: string): DocxDocument {
  const body = `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
  const documentXml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}</w:body></w:document>`;
  return DocxDocument.load(zipSync({
    "[Content_Types].xml": strToU8(`<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`),
    "_rels/.rels": strToU8(`<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`),
    "word/document.xml": strToU8(documentXml),
  }));
}

function addr(s: DocumentSession) {
  const para = s.doc.sections[0].blocks[0] as Paragraph;
  const run = para.children[0] as Run;
  const t = run.content.find((c) => c.kind === "text")!.srcT!;
  return { blockId: s.ids.idOf(para.src!)!, runId: s.ids.idOf(run.src!)!, len: t.text.length };
}

function bodyText(doc: DocxDocument): string {
  const body = doc.docRoot.children.find((c) => localName(c.name) === "body")!;
  const collectT = (el: XmlElement): string => {
    let s = localName(el.name) === "t" ? el.text : "";
    for (const c of el.children) s += collectT(c);
    return s;
  };
  return body.children.filter((c) => localName(c.name) === "p").map(collectT).join("\n");
}

const pos = (runId: number, offset: number): Position => ({ blockId: 1, runId, offset });

describe("formatRange transform (run-split remapping, F3)", () => {
  function fr(runId: number, start: number, end: number, before: number | undefined, middle: number, after: number | undefined): FormatRangeIntent {
    return { kind: "formatRange", clientId: "a", clientSeq: 1, base: 0, blockId: 1, runId, start, end, patch: { bold: true }, beforeId: before, middleId: middle, afterId: after };
  }
  it("remaps a position before the range into the before-piece", () => {
    // run 2 split at [3,6): before=10, middle=11, after=12.
    expect(transformPosition(pos(2, 1), [fr(2, 3, 6, 10, 11, 12)])).toEqual(pos(10, 1));
  });
  it("remaps a position inside the range into the middle-piece", () => {
    expect(transformPosition(pos(2, 4), [fr(2, 3, 6, 10, 11, 12)])).toEqual(pos(11, 1));
  });
  it("remaps a position after the range into the after-piece", () => {
    expect(transformPosition(pos(2, 8), [fr(2, 3, 6, 10, 11, 12)])).toEqual(pos(12, 2));
  });
  it("no before-piece when start==0: an early position lands in middle", () => {
    expect(transformPosition(pos(2, 0), [fr(2, 0, 4, undefined, 11, 12)])).toEqual(pos(11, 0));
  });
});

describe("DocumentSession formatRange apply", () => {
  it("bolds a middle sub-range, splitting the run into three addressable pieces", () => {
    const s = new DocumentSession(makeDoc("HelloWorld"));
    const a = addr(s);
    const e = s.submit({ kind: "formatRange", clientId: "a", clientSeq: 1, base: 0, blockId: a.blockId, runId: a.runId, start: 5, end: 10, patch: { bold: true }, beforeId: 100, middleId: 101, afterId: undefined });
    expect(e.kind).toBe("applied");
    expect(bodyText(s.doc)).toBe("HelloWorld");
    const xml = serializeXml(s.doc.docRoot);
    expect(xml).toContain("<w:b/>"); // middle piece bold
    // The middle piece is addressable by its carried id for a follow-up edit.
    const ins = s.submit({ kind: "insertText", clientId: "a", clientSeq: 2, base: s.seq, at: { blockId: a.blockId, runId: 101, offset: 5 }, text: "!" });
    expect(ins.kind).toBe("applied");
    expect(bodyText(s.doc)).toBe("HelloWorld!");
  });

  it("converges: a concurrent insert in a run that gets sub-range formatted lands in the right piece", () => {
    const s = new DocumentSession(makeDoc("abcdef"));
    const a = addr(s);
    // A bolds [2,4) (splitting into before=abc..., middle, after); B inserts "X"
    // at offset 5 (in the 'after' piece), both base 0.
    const format: FormatRangeIntent = { kind: "formatRange", clientId: "a", clientSeq: 1, base: 0, blockId: a.blockId, runId: a.runId, start: 2, end: 4, patch: { italic: true }, beforeId: 200, middleId: 201, afterId: 202 };
    const insert: InsertTextIntent = { kind: "insertText", clientId: "b", clientSeq: 1, base: 0, at: { blockId: a.blockId, runId: a.runId, offset: 5 }, text: "X" };
    s.submit(format);
    const e = s.submit(insert); // B's offset 5 remaps into the after-piece (202) at offset 1
    expect(e.kind).toBe("applied");
    expect(bodyText(s.doc)).toBe("abcdeXf");
    expect(serializeXml(s.doc.docRoot)).toContain("<w:i/>");
  });

  it("determinism across two sessions", () => {
    const build = () => {
      const s = new DocumentSession(makeDoc("formatme"));
      const a = addr(s);
      s.submit({ kind: "formatRange", clientId: "a", clientSeq: 1, base: 0, blockId: a.blockId, runId: a.runId, start: 0, end: 6, patch: { bold: true }, middleId: 300, afterId: 301 });
      return serializeXml(s.doc.docRoot);
    };
    expect(build()).toBe(build());
  });
});
