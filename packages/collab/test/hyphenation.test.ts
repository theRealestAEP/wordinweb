import { describe, expect, it } from "vitest";
import { zipSync, strToU8 } from "fflate";
import { DocxDocument, documentOperationBody, serializeXml } from "@wordinweb/core";
import { DocumentSession } from "../src/session.js";
import { ClientReplica } from "../src/replica.js";
import type { Intent } from "../src/intents.js";

function docBytes(): Uint8Array {
  const documentXml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>` +
    `<w:p><w:r><w:t xml:space="preserve">hello</w:t></w:r></w:p></w:body></w:document>`;
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

describe("setHyphenation on the wire (registered operation)", () => {
  it("writes the settings cluster byte-identically on every replica", () => {
    const initial = docBytes();
    const server = new DocumentSession(DocxDocument.load(initial));
    const a = new ClientReplica(initial);
    const b = new ClientReplica(initial);

    const set: Intent = {
      ...documentOperationBody("setHyphenation", { auto: true, zonePt: 18, noCaps: true }),
      clientId: "a", clientSeq: 1, base: 0,
    } as Intent;
    a.submitLocal(set);
    const e1 = server.submit(set);
    a.receive([e1]);
    b.receive([e1]);

    const settingsXml = (d: { settingsRoot: unknown }) => serializeXml(d.settingsRoot as never);
    const serverXml = settingsXml(server.doc as never);
    expect(settingsXml(a.doc as never)).toBe(serverXml);
    expect(settingsXml(b.doc as never)).toBe(serverXml);
    expect(serverXml).toContain("<w:autoHyphenation/>");
    expect(serverXml).toContain('<w:hyphenationZone w:val="360"/>'); // 18pt = 360 twips
    expect(serverXml).toContain("<w:doNotHyphenateCaps/>");

    // The same values again: honest no-op, every replica agrees.
    const again: Intent = {
      ...documentOperationBody("setHyphenation", { auto: true }),
      clientId: "b", clientSeq: 1, base: server.seq,
    } as Intent;
    b.submitLocal(again);
    const e2 = server.submit(again);
    expect((e2 as { kind: string }).kind).toBe("rejected");
  });

  it("rejects malformed payloads at validation", async () => {
    const { validateIntent } = await import("../src/validate.js");
    const base = { clientId: "a", clientSeq: 1, base: 0 };
    expect(validateIntent({ kind: "setHyphenation", ...base } as Intent)).toContain("empty patch");
    expect(validateIntent({ kind: "setHyphenation", zonePt: -1, ...base } as Intent)).toContain("bad zonePt");
    expect(validateIntent({ kind: "setHyphenation", auto: "yes", ...base } as Intent)).toContain("bad auto");
    expect(validateIntent({ kind: "setHyphenation", auto: true, zonePt: null, ...base } as Intent)).toBeNull();
  });
});
