import { describe, expect, it } from "vitest";
import { zipSync, strToU8 } from "fflate";
import {
  DocxDocument,
  compileReplaceAll,
  replaceAll,
  serializeXml,
  localName,
  type XmlElement,
} from "@wordinweb/core";
import { DocumentSession } from "../src/session.js";
import { ClientReplica } from "../src/replica.js";
import type { Intent } from "../src/intents.js";

/** Body paragraphs, each given as its inner run XML. */
function docBytes(paras: string[]): Uint8Array {
  const body = paras.map((runs) => `<w:p>${runs}</w:p>`).join("");
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

const t = (s: string) => `<w:r><w:t xml:space="preserve">${s}</w:t></w:r>`;

function paraText(doc: DocxDocument): string {
  const body = doc.docRoot.children.find((c) => localName(c.name) === "body")!;
  const collect = (el: XmlElement): string => {
    let s = localName(el.name) === "t" || localName(el.name) === "delText" ? el.text : "";
    for (const c of el.children) s += collect(c);
    return s;
  };
  return body.children.filter((c) => localName(c.name) === "p").map(collect).join("\n");
}

/** Live (undeleted) text only — what the document reads after the strikes. */
function liveText(doc: DocxDocument): string {
  const body = doc.docRoot.children.find((c) => localName(c.name) === "body")!;
  const collect = (el: XmlElement): string => {
    if (localName(el.name) === "del") return "";
    let s = localName(el.name) === "t" ? el.text : "";
    for (const c of el.children) s += collect(c);
    return s;
  };
  return body.children.filter((c) => localName(c.name) === "p").map(collect).join("\n");
}

/** Compile on `author`'s replica and drive every intent through the server
 * one-in-flight; every replica receives each canonical broadcast. */
function replaceAllOnWire(
  server: DocumentSession,
  author: ClientReplica,
  others: ClientReplica[],
  clientId: string,
  query: string,
  replacement: string,
  suggest?: { author: string; date: string },
): ReturnType<typeof compileReplaceAll>["result"] {
  const { intents, result } = compileReplaceAll(author.doc, query, replacement, undefined, suggest);
  let seq = 1;
  for (const body of intents) {
    const full = { ...body, clientId, clientSeq: seq++, base: server.seq } as Intent;
    author.submitLocal(full);
    const entry = server.submit(full);
    author.receive([entry]);
    for (const c of others) c.receive([entry]);
  }
  return result;
}

describe("replace-all on the collab wire (#112)", () => {
  it("plain replace-all converges byte-identically and matches the local mutation", () => {
    const initial = docBytes([t("the cat sat"), t("cat catalog cat")]);
    const server = new DocumentSession(DocxDocument.load(initial));
    const a = new ClientReplica(initial);
    const b = new ClientReplica(initial);

    const result = replaceAllOnWire(server, a, [b], "a", "cat", "kitten");
    expect(result.total).toBe(4);
    expect(result.byStory).toEqual({ body: 4 });

    const serverXml = serializeXml(server.doc.docRoot);
    expect(serializeXml(a.doc.docRoot)).toBe(serverXml);
    expect(serializeXml(b.doc.docRoot)).toBe(serverXml);
    expect(paraText(server.doc)).toBe("the kitten sat\nkitten kittenalog kitten");

    // The wire form lands the exact splices the local replaceAll performs.
    const control = DocxDocument.load(initial);
    const local = replaceAll(control, "cat", "kitten");
    expect(local.total).toBe(4);
    expect(serializeXml(control.docRoot)).toBe(serializeXml(server.doc.docRoot));
  });

  it("a match spanning split runs compiles per-range deletes and converges", () => {
    // "Hel" + "lo world Hel" + "lo" — two matches for "Hello", each crossing
    // a run boundary.
    const initial = docBytes([t("Hel") + t("lo world Hel") + t("lo")]);
    const server = new DocumentSession(DocxDocument.load(initial));
    const a = new ClientReplica(initial);
    const b = new ClientReplica(initial);

    const result = replaceAllOnWire(server, a, [b], "a", "Hello", "Bye");
    expect(result.total).toBe(2);

    const serverXml = serializeXml(server.doc.docRoot);
    expect(serializeXml(a.doc.docRoot)).toBe(serverXml);
    expect(serializeXml(b.doc.docRoot)).toBe(serverXml);
    expect(paraText(server.doc)).toBe("Bye world Bye");
  });

  it("shorter and longer replacements over several matches in ONE run keep offsets valid", () => {
    const initial = docBytes([t("cat cat cat")]);
    const server = new DocumentSession(DocxDocument.load(initial));
    const a = new ClientReplica(initial);
    const b = new ClientReplica(initial);

    replaceAllOnWire(server, a, [b], "a", "cat", "dragonfly");
    expect(paraText(server.doc)).toBe("dragonfly dragonfly dragonfly");
    const serverXml = serializeXml(server.doc.docRoot);
    expect(serializeXml(a.doc.docRoot)).toBe(serverXml);
    expect(serializeXml(b.doc.docRoot)).toBe(serverXml);
  });

  it("an empty replacement is a pure delete-all and converges", () => {
    const initial = docBytes([t("xx a xx b xx")]);
    const server = new DocumentSession(DocxDocument.load(initial));
    const a = new ClientReplica(initial);
    const b = new ClientReplica(initial);

    const result = replaceAllOnWire(server, a, [b], "a", "xx ", "");
    expect(result.total).toBe(2);
    expect(paraText(server.doc)).toBe("a b xx");
    expect(serializeXml(a.doc.docRoot)).toBe(serializeXml(server.doc.docRoot));
    expect(serializeXml(b.doc.docRoot)).toBe(serializeXml(server.doc.docRoot));
  });

  it("suggesting mode compiles strike-then-insert: struck text kept, replacement tracked, replicas byte-identical", () => {
    const initial = docBytes([t("cat cat cat"), t("the cat mat")]);
    const server = new DocumentSession(DocxDocument.load(initial));
    const a = new ClientReplica(initial);
    const b = new ClientReplica(initial);
    const meta = { author: "Reviewer", date: "2026-08-08T00:00:00Z" };

    const result = replaceAllOnWire(server, a, [b], "a", "cat", "dog", meta);
    expect(result.total).toBe(4);

    const serverXml = serializeXml(server.doc.docRoot);
    expect(serializeXml(a.doc.docRoot)).toBe(serverXml);
    expect(serializeXml(b.doc.docRoot)).toBe(serverXml);

    // Nothing removed: the struck originals survive as w:del/w:delText,
    // the replacements ride w:ins — Word's side-by-side replacement pair.
    expect(liveText(server.doc)).toBe("dog dog dog\nthe dog mat");
    expect(paraText(server.doc)).toBe("dogcat dogcat dogcat\nthe dogcat mat");
    expect(serverXml).toContain("w:delText");
    expect((serverXml.match(/<w:ins /g) ?? []).length).toBe(4);
    expect((serverXml.match(/<w:del /g) ?? []).length).toBe(4);
    expect(serverXml).toContain(`w:author="Reviewer"`);
  });

  it("a concurrent remote insert at base 0 still converges", () => {
    const initial = docBytes([t("one cat two")]);
    const server = new DocumentSession(DocxDocument.load(initial));
    const a = new ClientReplica(initial);
    const b = new ClientReplica(initial);

    // B types at the end of the same run, sequenced FIRST at the server.
    const para = server.doc.sections[0].blocks[0];
    const pEl = (para as { src?: XmlElement }).src!;
    const blockId = server.ids.idOf(pEl)!;
    const runEl = pEl.children.find((c) => localName(c.name) === "r")!;
    const runId = server.ids.idOf(runEl)!;
    const bIns: Intent = {
      kind: "insertText", clientId: "b", clientSeq: 1, base: 0,
      at: { blockId, runId, offset: 11 }, text: "!",
    };
    b.submitLocal(bIns);
    const eB = server.submit(bIns);

    // A compiled its replace against base 0 (no knowledge of B's insert).
    const { intents } = compileReplaceAll(a.doc, "cat", "feline");
    let seq = 1;
    const entries = [eB];
    for (const body of intents) {
      const full = { ...body, clientId: "a", clientSeq: seq++, base: 0 } as Intent;
      a.submitLocal(full);
      entries.push(server.submit(full));
    }
    a.receive(entries);
    b.receive(entries);

    const serverXml = serializeXml(server.doc.docRoot);
    expect(serializeXml(a.doc.docRoot)).toBe(serverXml);
    expect(serializeXml(b.doc.docRoot)).toBe(serverXml);
    expect(paraText(server.doc)).toBe("one feline two!");
  });
});
