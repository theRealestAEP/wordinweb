import { describe, expect, it } from "vitest";
import { zipSync, strToU8 } from "fflate";
import { DocxDocument, serializeXml, type Paragraph, type Run } from "@wordinweb/core";
import { DocumentSession } from "../src/session.js";

/**
 * Wire offset basis (checkpoint B1 rev 2): cumulative within the run, with
 * each inline separator (w:tab / w:br / w:cr) counting as ONE wire unit. The
 * separator unit makes encoding ONE-TO-ONE: "end of the w:t before a tab"
 * and "start of the w:t after it" are different physical positions and get
 * different wire offsets — under the earlier separator-blind basis both
 * collapsed to the same number, so boundary intents applied into the wrong
 * w:t remotely (originator-vs-canonical divergence) and boundary-starting
 * ranges always rejected (adversarial-review bugs 1a/1b).
 *
 * These tests drive WIRE intents (the canonical DocumentSession.submit path)
 * into a <w:t>HELLO</w:t><w:tab/><w:t>WORLD</w:t> run: wire offsets 0-5
 * cover HELLO, 5→6 is the tab, 6-11 cover WORLD.
 */

function makeTabDoc(): DocxDocument {
  const body = `<w:p><w:r><w:t xml:space="preserve">HELLO</w:t><w:tab/><w:t xml:space="preserve">WORLD</w:t></w:r></w:p>`;
  const documentXml = `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}</w:body></w:document>`;
  return DocxDocument.load(zipSync({
    "[Content_Types].xml": strToU8(`<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`),
    "_rels/.rels": strToU8(`<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`),
    "word/document.xml": strToU8(documentXml),
  }));
}

function session(): { s: DocumentSession; blockId: number; runId: number } {
  const s = new DocumentSession(makeTabDoc());
  const para = s.doc.sections[0].blocks[0] as Paragraph;
  const run = para.children[0] as Run;
  return { s, blockId: s.ids.idOf(para.src!)!, runId: s.ids.idOf(run.src!)! };
}

/** All w:t text contents of the doc's Nth paragraph's runs, in doc order. */
function paraTexts(s: DocumentSession, paraIndex = 0): string[] {
  const para = s.doc.sections[0].blocks[paraIndex] as Paragraph;
  const out: string[] = [];
  const walk = (el: { name: string; text: string; children: unknown[] }): void => {
    if (el.name.endsWith(":t")) out.push(el.text);
    for (const c of el.children) walk(c as never);
  };
  walk(para.src as never);
  return out;
}

describe("wire offset basis: cumulative with separator units (B1 rev 2)", () => {
  it("insertText into the SECOND w:t (wire 8 = 5 text + 1 tab + 2)", () => {
    const { s, blockId, runId } = session();
    const e = s.submit({ kind: "insertText", clientId: "a", clientSeq: 1, base: 0, at: { blockId, runId, offset: 8 }, text: "x" });
    expect(e.kind).toBe("applied");
    expect(paraTexts(s)).toEqual(["HELLO", "WOxRLD"]);
  });

  it("BOUNDARY is one-to-one: wire 5 = end of HELLO (before the tab), wire 6 = start of WORLD (after it)", () => {
    const a = session();
    const ea = a.s.submit({ kind: "insertText", clientId: "a", clientSeq: 1, base: 0, at: { blockId: a.blockId, runId: a.runId, offset: 5 }, text: "x" });
    expect(ea.kind).toBe("applied");
    expect(paraTexts(a.s)).toEqual(["HELLOx", "WORLD"]);

    const b = session();
    const eb = b.s.submit({ kind: "insertText", clientId: "a", clientSeq: 1, base: 0, at: { blockId: b.blockId, runId: b.runId, offset: 6 }, text: "x" });
    expect(eb.kind).toBe("applied");
    expect(paraTexts(b.s)).toEqual(["HELLO", "xWORLD"]);
  });

  it("splitParagraph disambiguates the tab side: wire 5 keeps the tab in the NEW paragraph, wire 6 keeps it in the OLD one", () => {
    // Enter at (HELLO, 5) — caret BEFORE the tab: the tab travels with the
    // moved tail into the new paragraph.
    const a = session();
    const ea = a.s.submit({ kind: "splitParagraph", clientId: "a", clientSeq: 1, base: 0, at: { blockId: a.blockId, runId: a.runId, offset: 5 }, newBlockId: 900, newRunId: 901 });
    expect(ea.kind).toBe("applied");
    expect(serializeXml((a.s.doc.sections[0].blocks[0] as Paragraph).src!)).not.toContain("w:tab");
    expect(serializeXml((a.s.doc.sections[0].blocks[1] as Paragraph).src!)).toContain("w:tab");

    // Enter at (WORLD, 0) — caret AFTER the tab: the tab stays behind.
    const b = session();
    const eb = b.s.submit({ kind: "splitParagraph", clientId: "a", clientSeq: 1, base: 0, at: { blockId: b.blockId, runId: b.runId, offset: 6 }, newBlockId: 900, newRunId: 901 });
    expect(eb.kind).toBe("applied");
    expect(serializeXml((b.s.doc.sections[0].blocks[0] as Paragraph).src!)).toContain("w:tab");
    expect(serializeXml((b.s.doc.sections[0].blocks[1] as Paragraph).src!)).not.toContain("w:tab");
  });

  it("deleteText STARTING at the post-tab boundary applies (review bug 1b: it always rejected)", () => {
    const { s, blockId, runId } = session();
    // Backspace over "W": wire [6,7).
    const e = s.submit({ kind: "deleteText", clientId: "a", clientSeq: 1, base: 0, blockId, runId, start: 6, end: 7 });
    expect(e.kind).toBe("applied");
    expect(paraTexts(s)).toEqual(["HELLO", "ORLD"]);
  });

  it("deleteText inside the SECOND w:t splices locally (wire [7,9) = local [1,3))", () => {
    const { s, blockId, runId } = session();
    const e = s.submit({ kind: "deleteText", clientId: "a", clientSeq: 1, base: 0, blockId, runId, start: 7, end: 9 });
    expect(e.kind).toBe("applied");
    expect(paraTexts(s)).toEqual(["HELLO", "WLD"]);
  });

  it("deleteText spanning the tab is a clean rejection (no wire form for separator removal)", () => {
    const { s, blockId, runId } = session();
    // [4,7) covers "O", the tab, and "W".
    const e = s.submit({ kind: "deleteText", clientId: "a", clientSeq: 1, base: 0, blockId, runId, start: 4, end: 7 });
    expect(e.kind).toBe("rejected");
    expect(paraTexts(s)).toEqual(["HELLO", "WORLD"]);
  });

  it("deleteText past the run's wire length is rejected", () => {
    const { s, blockId, runId } = session();
    const e = s.submit({ kind: "deleteText", clientId: "a", clientSeq: 1, base: 0, blockId, runId, start: 11, end: 13 });
    expect(e.kind).toBe("rejected");
    expect(paraTexts(s)).toEqual(["HELLO", "WORLD"]);
  });

  it("suggestRevision strikes a wire range in the SECOND w:t, including one starting at the boundary", () => {
    const { s, blockId, runId } = session();
    const e = s.submit({
      kind: "suggestRevision", clientId: "a", clientSeq: 1, base: 0,
      ranges: [{ blockId, runId, start: 6, end: 8 }],
      suggest: { author: "R", date: "2020-01-01T00:00:00Z" },
    });
    expect(e.kind).toBe("applied");
    const xml = serializeXml(s.doc.docRoot);
    expect(xml).toContain("w:del");
    expect(xml).toMatch(/<w:del[^>]*>(?:(?!<\/w:del>).)*WO(?:(?!<\/w:del>).)*<\/w:del>/s);
    expect(xml).not.toMatch(/<w:del[^>]*>(?:(?!<\/w:del>).)*HELLO/s);
  });

  it("formatRange with a wire range inside the SECOND w:t formats only it", () => {
    const { s, blockId, runId } = session();
    // Wire [8,10) = "RL" of WORLD. Pieces: before(HELLO+tab+WO), middle(RL), after(D).
    const e = s.submit({
      kind: "formatRange", clientId: "a", clientSeq: 1, base: 0,
      blockId, runId, start: 8, end: 10, patch: { bold: true },
      beforeId: 101, middleId: 102, afterId: 103,
    });
    expect(e.kind).toBe("applied");
    const middle = s.ids.elOf(102)!;
    expect(middle).toBeDefined();
    const middleTexts: string[] = [];
    const walk = (el: { name: string; text: string; children: unknown[] }): void => {
      if (el.name.endsWith(":t")) middleTexts.push(el.text);
      for (const c of el.children) walk(c as never);
    };
    walk(middle as never);
    expect(middleTexts.join("")).toBe("RL");
    const flat = serializeXml(s.doc.docRoot).replace(/<[^>]+>/g, "");
    expect(flat).toContain("HELLO");
    expect(flat).toContain("WO");
    expect(flat).toContain("D");
  });

  it("two replicas applying the same multi-w:t wire intents converge byte-identically", () => {
    const a = session();
    const b = session();
    const intents = [
      { kind: "insertText", clientId: "c", clientSeq: 1, base: 0, at: { blockId: a.blockId, runId: a.runId, offset: 8 }, text: "x" },
      { kind: "deleteText", clientId: "c", clientSeq: 2, base: 1, blockId: a.blockId, runId: a.runId, start: 6, end: 7 },
      { kind: "insertText", clientId: "c", clientSeq: 3, base: 2, at: { blockId: a.blockId, runId: a.runId, offset: 3 }, text: "y" },
    ] as const;
    for (const i of intents) {
      const ea = a.s.submit(i as never);
      const eb = b.s.submit(i as never);
      expect(ea.kind).toBe(eb.kind);
    }
    expect(serializeXml(a.s.doc.docRoot)).toBe(serializeXml(b.s.doc.docRoot));
  });
});
