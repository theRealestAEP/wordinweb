import { describe, expect, it } from "vitest";
import { CollabHub, type Connection, type DocProvider, type ServerMessage } from "../src/index.js";
import { MetricsObservability, serializeError } from "../src/observability.js";
import { PROTOCOL_VERSION } from "@wordinweb/collab/server";
import { zipSync, strToU8 } from "fflate";

/**
 * LOG HYGIENE — the zero-custody invariant as a PROPERTY, not a comment.
 *
 * observability.ts has always promised "shapes and counts, never content".
 * A promise in a doc comment survives exactly until someone adds a field to
 * help debug something. This drives a real session through the hub with
 * every secret set to a distinctive marker, then greps EVERY line the sink
 * emitted, at the most verbose level, for every marker.
 *
 * Deliberately blunt, for the same reason the presence-blindness pin is: a
 * narrower assertion ("the docId field is absent") only proves the field
 * someone thought to check. The question worth answering is whether the
 * string appears AT ALL.
 */

const SECRETS = {
  docId: "SECRETdocIdMARKER",
  clientId: "SECRETclientIdMARKER",
  docText: "SECRETdocumentTEXTmarker",
  ownerToken: "SECRETownerTokenMARKER",
  shareCode: "SECRETshareCodeMARKER",
  docKey: "SECRETdocKeyMARKER",
};

function docBytes(text: string): Uint8Array {
  const xml =
    `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
    `<w:body><w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p></w:body></w:document>`;
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
    "word/document.xml": strToU8(xml),
  });
}

class FakeConn implements Connection {
  received: ServerMessage[] = [];
  constructor(public id: string) {}
  send(m: ServerMessage): void {
    this.received.push(m);
  }
}

describe("structured logs never carry content or identifiers", () => {
  it("emits nothing containing a doc id, client id, document text, token, code or key", async () => {
    const lines: string[] = [];
    // The MOST verbose level: if a marker can reach a log line at all, debug
    // is where it happens.
    const obs = new MetricsObservability({ level: "debug", out: (l) => lines.push(l) });
    const provider: DocProvider = { load: () => docBytes(SECRETS.docText) };
    const hub = new CollabHub(provider, undefined, undefined, undefined, undefined, obs);

    const a = new FakeConn("cA");
    const b = new FakeConn("cB");

    // A full-shaped session: two joins, a takeover kick, a submit, a forged
    // submit (refused), presence, and a disconnect — every path that logs.
    await hub.handle(a, {
      t: "hello", protocolVersion: PROTOCOL_VERSION, docId: SECRETS.docId,
      clientId: SECRETS.clientId, sinceSeq: 0, ownerToken: SECRETS.ownerToken,
      codeProof: SECRETS.shareCode,
    } as never);
    await hub.handle(b, {
      t: "hello", protocolVersion: PROTOCOL_VERSION, docId: SECRETS.docId,
      clientId: SECRETS.clientId, sinceSeq: 0, takeover: true,
    } as never);
    await hub.handle(b, {
      t: "submit",
      intent: {
        kind: "insertText", clientId: SECRETS.clientId, clientSeq: 1, base: 0,
        at: { blockId: 1, runId: 2, offset: 0 }, text: SECRETS.docText,
      },
    } as never);
    await hub.handle(b, {
      t: "submit",
      intent: {
        kind: "insertText", clientId: "forged-" + SECRETS.clientId, clientSeq: 2, base: 0,
        at: { blockId: 1, runId: 2, offset: 0 }, text: SECRETS.docText,
      },
    } as never);
    await hub.handle(b, { t: "presence", position: { anchor: { blockId: 1, runId: 2, offset: 0 } } } as never);
    hub.disconnect(b);

    expect(lines.length, "the session did produce log lines").toBeGreaterThan(0);
    const all = lines.join("\n");
    for (const [name, secret] of Object.entries(SECRETS)) {
      expect(all, `${name} must never appear in a log line`).not.toContain(secret);
    }
    // …and the safe vocabulary IS there, so this is not passing by logging
    // nothing at all.
    expect(all).toContain("taken-over");
    expect(all).toContain("hello");
  });

  it("an error line carries a stack but still no session data", () => {
    const lines: string[] = [];
    const obs = new MetricsObservability({ level: "debug", out: (l) => lines.push(l) });
    // The realistic hazard: a message that INTERPOLATED a secret. The sink
    // cannot sanitize that (it is opaque text by then) — so what this pins is
    // that the sink adds nothing of its own, and that callers must not build
    // messages from session data. `where` is the fixed label discipline.
    obs.error("ws.handle", new Error("connection reset by peer"), { msgType: "submit" });
    const rec = JSON.parse(lines[0]) as Record<string, unknown>;
    expect(rec.level).toBe("error");
    expect(rec.where).toBe("ws.handle");
    expect(rec.msgType).toBe("submit");
    expect(typeof rec.stack).toBe("string");
    expect(rec.stack as string).toContain("Error");
    expect(lines.join("\n")).not.toContain(SECRETS.docId);
  });
});

describe("serializeError", () => {
  it("flattens a real Error with its stack", () => {
    const e = serializeError(new TypeError("boom"));
    expect(e.name).toBe("TypeError");
    expect(e.message).toBe("boom");
    expect(e.stack).toContain("TypeError");
  });

  it("survives non-Error throws — a string, null, undefined", () => {
    // `catch` gives you `unknown`, and a rejected promise with no reason is
    // real. A serializer that assumes Error is one more silent failure.
    expect(serializeError("just a string").message).toBe("just a string");
    expect(serializeError(null).message).toBe("null");
    expect(serializeError(undefined).message).toBe("undefined");
    expect(() => serializeError({ weird: true })).not.toThrow();
  });

  it("walks a cause chain, capped", () => {
    const root = new Error("root");
    const mid = new Error("mid", { cause: root });
    const top = new Error("top", { cause: mid });
    const e = serializeError(top);
    expect(e.message).toBe("top");
    expect(e.causes).toEqual(["Error: mid", "Error: root"]);
  });

  it("caps a pathological stack so one throw cannot flood a collector", () => {
    const e = new Error("x");
    e.stack = "F".repeat(100_000);
    expect((serializeError(e).stack ?? "").length).toBeLessThanOrEqual(4000);
  });
});
