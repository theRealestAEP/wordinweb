import { describe, expect, it } from "vitest";
import { zipSync, strToU8 } from "fflate";
import { serializeXml, localName, type XmlElement } from "@wordinweb/core";
import { CollabConnection } from "../src/connection.js";
import { CollabHubLoopback } from "./loopback.js";

function docBytes(text: string): Uint8Array {
  const documentXml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>` +
    `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p></w:body></w:document>`;
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

function text(doc: { docRoot: XmlElement } | null): string {
  if (!doc) return "";
  const body = doc.docRoot.children.find((c) => localName(c.name) === "body")!;
  const collectT = (el: XmlElement): string => {
    let s = localName(el.name) === "t" ? el.text : "";
    for (const c of el.children) s += collectT(c);
    return s;
  };
  return body.children.filter((c) => localName(c.name) === "p").map(collectT).join("\n");
}

describe("CollabConnection over an in-process hub", () => {
  it("joins, receives a welcome, and applies a local edit optimistically", () => {
    const hub = new CollabHubLoopback(() => docBytes("hi"));
    const conn = new CollabConnection(hub.connect(), "a");
    conn.join("d");
    expect(conn.ready).toBe(true);
    // The doc's paragraph/run ids in parse order: block 1, run 2.
    conn.submit({ kind: "insertText", at: { blockId: 1, runId: 2, offset: 2 }, text: "!" } as never);
    expect(text(conn.doc)).toBe("hi!");
  });

  it("propagates one client's edit to another client", () => {
    const hub = new CollabHubLoopback(() => docBytes("ab"));
    const a = new CollabConnection(hub.connect(), "a");
    const b = new CollabConnection(hub.connect(), "b");
    a.join("d");
    b.join("d");
    a.submit({ kind: "insertText", at: { blockId: 1, runId: 2, offset: 2 }, text: "C" } as never);
    // b receives the broadcast synchronously through the loopback.
    expect(text(b.doc)).toBe(text(a.doc));
    expect(text(b.doc)).toBe("abC");
  });

  it("delivers presence to the other participant, not the sender", () => {
    const hub = new CollabHubLoopback(() => docBytes("x"));
    const seen: { participant: string; pos: unknown }[] = [];
    const a = new CollabConnection(hub.connect(), "a");
    const b = new CollabConnection(hub.connect(), "b", {
      onPresence: (participant, position) => seen.push({ participant, pos: position }),
    });
    a.join("d");
    b.join("d");
    a.setPresence({ anchor: { blockId: 1, runId: 2, offset: 1 } });
    expect(seen).toHaveLength(1);
    expect(seen[0].participant).toBeTypeOf("string");
  });

  it("surfaces a version refusal when the transport advertises a bad version", () => {
    const hub = new CollabHubLoopback(() => docBytes("x"));
    let refused = "";
    const transport = hub.connect();
    new CollabConnection(transport, "a", { onRefused: (r) => (refused = r) });
    // Send a hello with a deliberately wrong protocol version.
    transport.send({ t: "hello", protocolVersion: 999, docId: "d", sinceSeq: 0 });
    expect(refused).toBe("version-mismatch");
    void serializeXml; // keep import used across builds
  });
});
