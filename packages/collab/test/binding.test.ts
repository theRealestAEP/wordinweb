import { describe, expect, it } from "vitest";
import { zipSync, strToU8 } from "fflate";
import { DocxDocument, localName, type XmlElement } from "@wordinweb/core";
import { CollabConnection } from "../src/connection.js";
import { bindEditor, EditorBridge } from "../src/binding.js";
import { Intent } from "../src/intents.js";
import { PresencePosition } from "../src/protocol.js";
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

function text(doc: DocxDocument | null): string {
  if (!doc) return "";
  const body = doc.docRoot.children.find((c) => localName(c.name) === "body")!;
  const collectT = (el: XmlElement): string => {
    let s = localName(el.name) === "t" ? el.text : "";
    for (const c of el.children) s += collectT(c);
    return s;
  };
  return body.children.filter((c) => localName(c.name) === "p").map(collectT).join("\n");
}

/** A DOM-free EditorBridge double: records the doc it's told to render and
 * exposes a way to fire a local intent, as the real DocxEditor would. */
class FakeEditor implements EditorBridge {
  doc: DocxDocument | null = null;
  private intentHandler: ((i: Omit<Intent, "clientId" | "clientSeq" | "base">) => void) | null = null;
  private presenceHandler: ((p: PresencePosition | null) => void) | null = null;
  remotePresence: { participant: string; pos: PresencePosition | null }[] = [];

  setDocument(doc: DocxDocument): void {
    this.doc = doc;
  }
  onLocalIntent(handler: (i: Omit<Intent, "clientId" | "clientSeq" | "base">) => void): () => void {
    this.intentHandler = handler;
    return () => (this.intentHandler = null);
  }
  onLocalPresence(handler: (p: PresencePosition | null) => void): () => void {
    this.presenceHandler = handler;
    return () => (this.presenceHandler = null);
  }
  setRemotePresence(participant: string, pos: PresencePosition | null): void {
    this.remotePresence.push({ participant, pos });
  }
  // Test drivers:
  fireIntent(i: Omit<Intent, "clientId" | "clientSeq" | "base">): void {
    this.intentHandler?.(i);
  }
  firePresence(p: PresencePosition | null): void {
    this.presenceHandler?.(p);
  }
}

describe("bindEditor", () => {
  it("pushes the joined document into the editor and applies a local intent", () => {
    const hub = new CollabHubLoopback(() => docBytes("hi"));
    const conn = new CollabConnection(hub.connect(), "a");
    const ed = new FakeEditor();
    conn.join("d");
    bindEditor(conn, ed);
    expect(text(ed.doc)).toBe("hi"); // welcome pushed

    ed.fireIntent({ kind: "insertText", at: { blockId: 1, runId: 2, offset: 2 }, text: "!" } as never);
    expect(text(ed.doc)).toBe("hi!"); // optimistic apply reflected in the pushed doc
  });

  it("propagates one editor's edit to a second bound editor", () => {
    const hub = new CollabHubLoopback(() => docBytes("ab"));
    const a = new CollabConnection(hub.connect(), "a");
    const b = new CollabConnection(hub.connect(), "b");
    const ea = new FakeEditor();
    const eb = new FakeEditor();
    a.join("d");
    b.join("d");
    bindEditor(a, ea);
    bindEditor(b, eb);

    ea.fireIntent({ kind: "insertText", at: { blockId: 1, runId: 2, offset: 2 }, text: "C" } as never);
    expect(text(eb.doc)).toBe("abC"); // b's editor received the broadcast
    expect(text(ea.doc)).toBe(text(eb.doc));
  });

  it("routes local presence to the wire and remote presence to the editor", () => {
    const hub = new CollabHubLoopback(() => docBytes("x"));
    const a = new CollabConnection(hub.connect(), "a");
    const b = new CollabConnection(hub.connect(), "b");
    const ea = new FakeEditor();
    const eb = new FakeEditor();
    a.join("d");
    b.join("d");
    bindEditor(a, ea);
    bindEditor(b, eb);

    ea.firePresence({ anchor: { blockId: 1, runId: 2, offset: 1 } });
    expect(eb.remotePresence).toHaveLength(1);
    expect(eb.remotePresence[0].pos).toEqual({ anchor: { blockId: 1, runId: 2, offset: 1 } });
  });

  it("dispose unsubscribes local intent handling", () => {
    const hub = new CollabHubLoopback(() => docBytes("hi"));
    const conn = new CollabConnection(hub.connect(), "a");
    const ed = new FakeEditor();
    conn.join("d");
    const binding = bindEditor(conn, ed);
    binding.dispose();
    ed.fireIntent({ kind: "insertText", at: { blockId: 1, runId: 2, offset: 2 }, text: "!" } as never);
    expect(text(ed.doc)).toBe("hi"); // no-op after dispose
  });
});
