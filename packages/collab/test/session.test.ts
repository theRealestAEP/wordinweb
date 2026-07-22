import { describe, expect, it } from "vitest";
import { zipSync, strToU8 } from "fflate";
import { DocxDocument, serializeXml, localName, type Paragraph, type Run, type XmlElement } from "@wordinweb/core";
import { DocumentSession } from "../src/session.js";
import { InsertTextIntent, SplitParagraphIntent, DeleteTextIntent } from "../src/intents.js";

function makeDoc(paras: string[]): DocxDocument {
  const body = paras
    .map((t) => `<w:p><w:r><w:t xml:space="preserve">${t}</w:t></w:r></w:p>`)
    .join("");
  const documentXml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}</w:body></w:document>`;
  const zip = zipSync({
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
  });
  return DocxDocument.load(zip);
}

/** Read paragraph text straight from the XML tree (the source of truth), so
 * it reflects text inserts that skip a full model refresh (the session's
 * text-only fast path). */
function docText(doc: DocxDocument): string {
  const body = doc.docRoot.children.find((c) => localName(c.name) === "body");
  if (!body) return "";
  const out: string[] = [];
  const collectT = (el: XmlElement): string => {
    let s = "";
    if (localName(el.name) === "t") s += el.text;
    for (const c of el.children) s += collectT(c);
    return s;
  };
  for (const child of body.children) {
    if (localName(child.name) === "p") out.push(collectT(child));
  }
  return out.join("\n");
}

function runIdOf(session: DocumentSession, paraIdx: number): { blockId: number; runId: number } {
  const para = session.doc.sections[0].blocks[paraIdx] as Paragraph;
  const run = para.children[0] as Run;
  return { blockId: session.ids.idOf(para.src!)!, runId: session.ids.idOf(run.src!)! };
}

describe("DocumentSession convergence", () => {
  it("sequences a single client's inserts and applies them to the authoritative tree", () => {
    const session = new DocumentSession(makeDoc(["Hello"]));
    const { blockId, runId } = runIdOf(session, 0);
    const i1: InsertTextIntent = { kind: "insertText", clientId: "a", clientSeq: 1, base: 0, at: { blockId, runId, offset: 5 }, text: " world" };
    const e1 = session.submit(i1);
    expect(e1.kind).toBe("applied");
    expect(session.seq).toBe(1);
    expect(docText(session.doc)).toBe("Hello world");
  });

  it("resolves two concurrent inserts in the same run via the canonical transform", () => {
    const session = new DocumentSession(makeDoc(["Hello"]));
    const { blockId, runId } = runIdOf(session, 0);
    // Both clients produced their intent at base 0. Client A inserts "X" at 0,
    // client B inserts "Y" at 5 — B's intent is transformed against A's.
    const a: InsertTextIntent = { kind: "insertText", clientId: "a", clientSeq: 1, base: 0, at: { blockId, runId, offset: 0 }, text: "X" };
    const b: InsertTextIntent = { kind: "insertText", clientId: "b", clientSeq: 1, base: 0, at: { blockId, runId, offset: 5 }, text: "Y" };
    session.submit(a); // seq 1: "XHello"
    session.submit(b); // seq 2: B's offset 5 shifts to 6 → "XHelloY"
    expect(docText(session.doc)).toBe("XHelloY");
  });

  it("both clients' intended text survives when they type at their own carets (intention check)", () => {
    // A types "AAA" at the start, B types "BBB" at the end, both from base 0.
    const session = new DocumentSession(makeDoc(["mid"]));
    const { blockId, runId } = runIdOf(session, 0);
    session.submit({ kind: "insertText", clientId: "a", clientSeq: 1, base: 0, at: { blockId, runId, offset: 0 }, text: "AAA" });
    session.submit({ kind: "insertText", clientId: "b", clientSeq: 1, base: 0, at: { blockId, runId, offset: 3 }, text: "BBB" });
    const text = docText(session.doc);
    expect(text).toContain("AAA");
    expect(text).toContain("BBB");
    expect(text).toBe("AAAmidBBB");
  });

  it("deduplicates a re-sent intent by idempotency key", () => {
    const session = new DocumentSession(makeDoc(["Hi"]));
    const { blockId, runId } = runIdOf(session, 0);
    const i: InsertTextIntent = { kind: "insertText", clientId: "a", clientSeq: 7, base: 0, at: { blockId, runId, offset: 2 }, text: "!" };
    const first = session.submit(i);
    const second = session.submit(i); // re-send
    expect(second).toBe(first);
    expect(session.seq).toBe(1);
    expect(docText(session.doc)).toBe("Hi!");
  });

  it("records a rejection (sequenced no-op) for an unresolvable position", () => {
    const session = new DocumentSession(makeDoc(["x"]));
    const bad: DeleteTextIntent = { kind: "deleteText", clientId: "a", clientSeq: 1, base: 0, blockId: 1, runId: 9999, start: 0, end: 1 };
    const e = session.submit(bad);
    expect(e.kind).toBe("rejected");
    expect(session.seq).toBe(1); // rejection occupies a sequence position
  });

  it("splits a paragraph and keeps subsequent intents addressable via carried ids", () => {
    const session = new DocumentSession(makeDoc(["HelloWorld"]));
    const { blockId, runId } = runIdOf(session, 0);
    const split: SplitParagraphIntent = {
      kind: "splitParagraph", clientId: "a", clientSeq: 1, base: 0,
      at: { blockId, runId, offset: 5 }, newBlockId: 5000, newRunId: 5001,
    };
    session.submit(split);
    expect(docText(session.doc)).toBe("Hello\nWorld");
    // The moved-tail run is addressable by its carried id.
    const insertInSecond: InsertTextIntent = {
      kind: "insertText", clientId: "a", clientSeq: 2, base: session.seq,
      at: { blockId: 5000, runId: 5001, offset: 5 }, text: "!",
    };
    const e = session.submit(insertInSecond);
    expect(e.kind).toBe("applied");
    expect(docText(session.doc)).toBe("Hello\nWorld!");
  });

  it("produces identical serialized XML for two sessions given the same intent stream (determinism)", () => {
    const build = () => {
      const s = new DocumentSession(makeDoc(["abc"]));
      const { blockId, runId } = runIdOf(s, 0);
      s.submit({ kind: "insertText", clientId: "a", clientSeq: 1, base: 0, at: { blockId, runId, offset: 3 }, text: "def" });
      s.submit({ kind: "insertText", clientId: "b", clientSeq: 1, base: 0, at: { blockId, runId, offset: 0 }, text: "ZZ" });
      return serializeXml(s.doc.docRoot);
    };
    expect(build()).toBe(build());
  });
});

describe("DocumentSession deleteText", () => {
  it("applies a delete and transforms a concurrent delete in the same run", () => {
    const session = new DocumentSession(makeDoc(["abcdefgh"]));
    const { blockId, runId } = runIdOf(session, 0);
    // Delete "cd" (2..4). Then a concurrent delete of "gh" (6..8) from base 0
    // must shift left by the 2 chars the first delete removed → 4..6.
    session.submit({ kind: "deleteText", clientId: "a", clientSeq: 1, base: 0, blockId, runId, start: 2, end: 4 });
    session.submit({ kind: "deleteText", clientId: "b", clientSeq: 1, base: 0, blockId, runId, start: 6, end: 8 });
    expect(docText(session.doc)).toBe("abef");
  });

  it("rejects an out-of-range delete", () => {
    const session = new DocumentSession(makeDoc(["hi"]));
    const { blockId, runId } = runIdOf(session, 0);
    const e = session.submit({ kind: "deleteText", clientId: "a", clientSeq: 1, base: 0, blockId, runId, start: 0, end: 99 });
    expect(e.kind).toBe("rejected");
  });
});

describe("DocumentSession formatRun (whole-run character formatting)", () => {
  function isBold(session: DocumentSession, paraIdx: number): boolean {
    const para = session.doc.sections[paraIdx].blocks ? undefined : undefined;
    void para;
    // Read bold from the serialized run properties.
    return serializeXml(session.doc.docRoot).includes("<w:b/>");
  }

  it("applies bold to a whole run and preserves the run id", () => {
    const session = new DocumentSession(makeDoc(["word"]));
    const { blockId, runId } = runIdOf(session, 0);
    const e = session.submit({ kind: "formatRun", clientId: "a", clientSeq: 1, base: 0, blockId, runId, patch: { bold: true } });
    expect(e.kind).toBe("applied");
    expect(isBold(session, 0)).toBe(true);
    // The run id survives (whole-run format is in-place).
    const para = session.doc.sections[0].blocks[0] as Paragraph;
    expect(session.ids.idOf((para.children[0] as Run).src!)).toBe(runId);
  });

  it("a concurrent insert in the run being formatted is unaffected by the format transform", () => {
    const session = new DocumentSession(makeDoc(["abc"]));
    const { blockId, runId } = runIdOf(session, 0);
    // Format the run bold, and concurrently (base 0) insert at offset 3.
    session.submit({ kind: "formatRun", clientId: "a", clientSeq: 1, base: 0, blockId, runId, patch: { italic: true } });
    session.submit({ kind: "insertText", clientId: "b", clientSeq: 1, base: 0, at: { blockId, runId, offset: 3 }, text: "d" });
    expect(docText(session.doc)).toBe("abcd");
    expect(serializeXml(session.doc.docRoot)).toContain("<w:i/>");
  });

  it("determinism: same format+edit stream yields identical XML twice", () => {
    const build = () => {
      const s = new DocumentSession(makeDoc(["xy"]));
      const { blockId, runId } = runIdOf(s, 0);
      s.submit({ kind: "formatRun", clientId: "a", clientSeq: 1, base: 0, blockId, runId, patch: { bold: true, underline: true } });
      s.submit({ kind: "insertText", clientId: "a", clientSeq: 2, base: s.seq, at: { blockId, runId, offset: 2 }, text: "z" });
      return serializeXml(s.doc.docRoot);
    };
    expect(build()).toBe(build());
  });
});
