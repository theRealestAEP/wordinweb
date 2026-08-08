import { describe, expect, it } from "vitest";
import { zipSync, strToU8 } from "fflate";
import { DocxDocument, operationBody, serializeXml, localName } from "@wordinweb/core";
import { DocumentSession } from "../src/session.js";
import { ClientReplica } from "../src/replica.js";
import type { Intent } from "../src/intents.js";

function docBytes(paras: string[]): Uint8Array {
  const body = paras.map((t) => `<w:p><w:r><w:t xml:space="preserve">${t}</w:t></w:r></w:p>`).join("");
  const documentXml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}</w:body></w:document>`;
  return zipSync({
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
}

function blockIdOf(session: DocumentSession, idx: number): number {
  const body = session.doc.docRoot.children.find((c) => localName(c.name) === "body")!;
  const p = body.children.filter((c) => localName(c.name) === "p")[idx];
  return session.ids.idOf(p)!;
}

describe("setTabStops on the wire (registered operation)", () => {
  it("applies on every replica byte-identically, tracked form included", () => {
    const initial = docBytes(["alpha", "beta"]);
    const server = new DocumentSession(DocxDocument.load(initial));
    const a = new ClientReplica(initial);
    const b = new ClientReplica(initial);

    const set: Intent = {
      ...operationBody("setTabStops", blockIdOf(server, 0), {
        stops: [
          { posPt: 72, align: "left" as const, leader: "none" as const },
          { posPt: 396, align: "right" as const, leader: "dot" as const },
        ],
      }),
      clientId: "a", clientSeq: 1, base: 0,
    } as Intent;
    a.submitLocal(set);
    const e1 = server.submit(set);
    a.receive([e1]);
    b.receive([e1]);

    // Suggested edit on the second paragraph records a w:pPrChange.
    const suggest: Intent = {
      ...operationBody("setTabStops", blockIdOf(server, 1), {
        stops: [{ posPt: 36, align: "center" as const, leader: "underscore" as const }],
        suggest: { author: "Reviewer", date: "2026-08-08T00:00:00Z" },
      }),
      clientId: "b", clientSeq: 1, base: server.seq,
    } as Intent;
    b.submitLocal(suggest);
    const e2 = server.submit(suggest);
    a.receive([e2]);
    b.receive([e2]);

    const serverXml = serializeXml(server.doc.docRoot);
    expect(serializeXml(a.doc.docRoot)).toBe(serverXml);
    expect(serializeXml(b.doc.docRoot)).toBe(serverXml);
    expect(serverXml).toContain(`<w:tab w:val="right" w:leader="dot" w:pos="7920"/>`);
    expect(serverXml).toContain("w:pPrChange");

    // An empty list clears w:tabs — and still converges.
    const clear: Intent = {
      ...operationBody("setTabStops", blockIdOf(server, 0), { stops: [] }),
      clientId: "a", clientSeq: 2, base: server.seq,
    } as Intent;
    a.submitLocal(clear);
    const e3 = server.submit(clear);
    a.receive([e3]);
    b.receive([e3]);
    expect(serializeXml(a.doc.docRoot)).toBe(serializeXml(server.doc.docRoot));
    expect(serializeXml(server.doc.docRoot)).not.toContain(`w:pos="1440"`);
  });

  it("rejects malformed payloads at validation", async () => {
    const { validateIntent } = await import("../src/validate.js");
    const base = { clientId: "a", clientSeq: 1, base: 0 };
    expect(validateIntent({ kind: "setTabStops", blockId: 1, stops: [{ posPt: 99999, align: "left", leader: "none" }], ...base } as Intent)).toContain("bad pos");
    expect(validateIntent({ kind: "setTabStops", blockId: 1, stops: [{ posPt: 10, align: "wavy", leader: "none" }], ...base } as Intent)).toContain("bad align");
    expect(validateIntent({ kind: "setTabStops", blockId: 1, stops: [{ posPt: 10, align: "left", leader: "none" }], ...base } as Intent)).toBeNull();
  });
});

describe("setParagraphBorders on the wire (registered operation)", () => {
  it("applies on every replica byte-identically", () => {
    const initial = docBytes(["one", "two"]);
    const server = new DocumentSession(DocxDocument.load(initial));
    const a = new ClientReplica(initial);
    const b = new ClientReplica(initial);
    const rule = { style: "single" as const, sz: 4, color: "auto" };

    const set: Intent = {
      ...operationBody("setParagraphBorders", blockIdOf(server, 0), {
        patch: { borders: { top: rule, bottom: rule, left: rule, right: rule }, shading: "FFF2CC" },
      }),
      clientId: "a", clientSeq: 1, base: 0,
    } as Intent;
    a.submitLocal(set);
    const e1 = server.submit(set);
    a.receive([e1]);
    b.receive([e1]);

    const serverXml = serializeXml(server.doc.docRoot);
    expect(serializeXml(a.doc.docRoot)).toBe(serverXml);
    expect(serializeXml(b.doc.docRoot)).toBe(serverXml);
    expect(serverXml).toContain("<w:pBdr>");
    expect(serverXml).toContain(`w:fill="FFF2CC"`);

    const clear: Intent = {
      ...operationBody("setParagraphBorders", blockIdOf(server, 0), {
        patch: { borders: { top: null, bottom: null, left: null, right: null }, shading: null },
      }),
      clientId: "b", clientSeq: 1, base: server.seq,
    } as Intent;
    b.submitLocal(clear);
    const e2 = server.submit(clear);
    a.receive([e2]);
    b.receive([e2]);
    expect(serializeXml(a.doc.docRoot)).toBe(serializeXml(server.doc.docRoot));
    expect(serializeXml(server.doc.docRoot)).not.toContain("pBdr");
  });

  it("rejects malformed payloads at validation", async () => {
    const { validateIntent } = await import("../src/validate.js");
    const base = { clientId: "a", clientSeq: 1, base: 0 };
    expect(validateIntent({ kind: "setParagraphBorders", blockId: 1, patch: {}, ...base } as Intent)).toContain("empty patch");
    expect(validateIntent({ kind: "setParagraphBorders", blockId: 1, patch: { borders: { diag: { style: "single" } } }, ...base } as Intent)).toContain("bad edge");
    expect(validateIntent({ kind: "setParagraphBorders", blockId: 1, patch: { shading: "not-a-color" }, ...base } as Intent)).toContain("bad shading");
    expect(validateIntent({ kind: "setParagraphBorders", blockId: 1, patch: { borders: { top: { style: "single" } } }, ...base } as Intent)).toBeNull();
  });
});
