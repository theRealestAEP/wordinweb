import { describe, expect, it } from "vitest";
import { zipSync, strToU8 } from "fflate";
import { localName, type XmlElement } from "@wordinweb/core";
import { CollabConnection } from "../src/connection.js";
import { CollabHubLoopback } from "./loopback.js";
import type { ClientMessage, ServerMessage } from "../src/protocol.js";
import type { ClientTransport } from "../src/connection.js";

/**
 * SUBMIT BEFORE THE WELCOME IS SILENTLY LOST (B13).
 *
 * Both connections open their submit path with a readiness guard —
 * `if (!this.replica) return;` in the plaintext connection, plus keys and
 * genesis in the encrypted one. The guard is correct (there is nothing to
 * apply against and nothing to seal with) but it returns SILENTLY: the caller
 * gets no error, no rejection, no counter moves, and the intent is not queued
 * for later. It simply ceases to exist.
 *
 * Why this matters beyond tidiness: the window is not theoretical. A client
 * joining a large encrypted document spends real time deriving epoch keys and
 * opening the checkpoint before `replica` exists, and the editor's submit path
 * is reachable throughout. Anything submitted in that window — a keystroke, a
 * toolbar command, a test harness's burst — is gone, and the only evidence is
 * a document that quietly lacks it.
 *
 * This is the shape B13 measured in the swarm: 403 submitOp calls, 172
 * envelopes on the wire, 172 tokens in the document, ZERO refusals and ZERO
 * rejections server-side. These tests pin the loss channel itself; whether it
 * accounts for all of B13's 231 is a separate question the swarm must answer.
 */

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
  const collect = (el: XmlElement): string => {
    let s = localName(el.name) === "t" ? el.text : "";
    for (const c of el.children) s += collect(c);
    return s;
  };
  return body.children.filter((c) => localName(c.name) === "p").map(collect).join("\n");
}

/** A transport that records everything sent and can hold the welcome back,
 * reproducing the join window a slow (large / encrypted) document creates. */
function deferrableTransport(inner: ClientTransport): {
  transport: ClientTransport;
  sent: ClientMessage[];
  release: () => void;
} {
  const sent: ClientMessage[] = [];
  let held: ServerMessage[] | null = [];
  let deliver: ((m: ServerMessage) => void) | null = null;
  const transport: ClientTransport = {
    send: (msg) => {
      sent.push(msg);
      inner.send(msg);
    },
    onMessage: (handler) => {
      deliver = handler;
      inner.onMessage((msg) => {
        if (held) held.push(msg);
        else handler(msg);
      });
    },
  };
  return {
    transport,
    sent,
    release: () => {
      const queued = held ?? [];
      held = null;
      for (const m of queued) deliver?.(m);
    },
  };
}

describe("a submit before the connection is ready", () => {
  it("is dropped — not applied, not sent, not queued — but LOUDLY", () => {
    const hub = new CollabHubLoopback(() => docBytes("hi"));
    const { transport, sent, release } = deferrableTransport(hub.connect());
    let changes = 0;
    const dropped: string[] = [];
    const conn = new CollabConnection(transport, "a", {
      onChange: () => changes++,
      onSubmitDropped: (reason) => dropped.push(reason),
    });

    conn.join("d"); // the welcome is held by the transport
    expect(conn.ready, "the welcome has not arrived yet").toBe(false);

    conn.submit({ kind: "insertText", at: { blockId: 1, runId: 2, offset: 2 }, text: "!" } as never);

    // The intent still cannot be honoured — there is nothing to apply it to.
    expect(sent.filter((m) => m.t === "submit"), "nothing reached the wire").toEqual([]);
    expect(changes, "no change callback fired").toBe(0);
    expect(conn.doc, "nothing was applied locally").toBeNull();
    // …but the caller is TOLD, and it is counted. This is the difference
    // between a bug you can see and one that presents as data loss weeks
    // later in a swarm run.
    expect(dropped, "the caller was told, with a reason").toEqual(["not-ready"]);
    expect(conn.droppedPreReady, "and it is counted for telemetry").toBe(1);

    // Still not replayed once ready — dropping is the contract, silence was
    // the defect.
    release();
    expect(conn.ready).toBe(true);
    expect(text(conn.doc), "the edit is not resurrected").toBe("hi");
    expect(sent.filter((m) => m.t === "submit"), "still nothing on the wire").toEqual([]);
    expect(conn.droppedPreReady, "the counter is not reset by becoming ready").toBe(1);
  });

  it("counts every submit in the window, and the ones after it survive", () => {
    const hub = new CollabHubLoopback(() => docBytes("hi"));
    const { transport, sent, release } = deferrableTransport(hub.connect());
    const dropped: string[] = [];
    const conn = new CollabConnection(transport, "a", { onSubmitDropped: (r) => dropped.push(r) });
    conn.join("d");

    // Five submits during the join window — the shape of a client that starts
    // editing (or a harness that starts bursting) the moment the page is up.
    for (let i = 0; i < 5; i++) {
      conn.submit({ kind: "insertText", at: { blockId: 1, runId: 2, offset: 2 }, text: `${i}` } as never);
    }
    release();

    // Ready now: a submit from here on behaves.
    conn.submit({ kind: "insertText", at: { blockId: 1, runId: 2, offset: 2 }, text: "X" } as never);

    expect(sent.filter((m) => m.t === "submit").length, "only the post-ready submit was sent").toBe(1);
    expect(text(conn.doc), "the five are absent; the sixth is present").toBe("hiX");
    // THE NUMBER THAT MAKES THE NEXT FAILURE DIAGNOSABLE. A swarm run that
    // loses N intents can now compare N against this counter: equal means
    // they all died at this guard, fewer means there is a second channel.
    expect(conn.droppedPreReady, "all five are accounted for").toBe(5);
    expect(dropped, "and each was reported as it happened").toEqual(
      ["not-ready", "not-ready", "not-ready", "not-ready", "not-ready"],
    );
  });

  it("`ready` is the only signal a caller has, and it is honest", () => {
    // Not a defect — the contract that makes the loss avoidable. Anything
    // gating submits on `ready` is safe; the problem is that nothing forces a
    // caller to, and the failure is invisible when they don't.
    const hub = new CollabHubLoopback(() => docBytes("hi"));
    const { transport, release } = deferrableTransport(hub.connect());
    const conn = new CollabConnection(transport, "a");
    conn.join("d");
    expect(conn.ready).toBe(false);
    release();
    expect(conn.ready).toBe(true);
  });
});
