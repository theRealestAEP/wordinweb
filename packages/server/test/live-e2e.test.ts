import { describe, expect, it } from "vitest";
import { WebSocketServer, WebSocket } from "ws";
import { CollabHub } from "../src/hub.js";
import { attachWebSocketServer } from "../src/ws.js";
import { blankProvider } from "../src/blank.js";
import { InMemoryStorage } from "../src/storage.js";
import { CollabConnection, createWebSocketTransport } from "@wordinweb/collab/client";
import { serializeXml, localName, type XmlElement } from "@wordinweb/core";

function paraText(doc: { docRoot: XmlElement } | null): string {
  if (!doc) return "";
  const body = doc.docRoot.children.find((c) => localName(c.name) === "body")!;
  const walk = (el: XmlElement): string => { let t = localName(el.name) === "t" ? el.text : ""; el.children.forEach((c) => (t += walk(c))); return t; };
  return walk(body);
}
const wait = (ms: number) => new Promise((r) => setTimeout(r, ms));
async function until(pred: () => boolean, label: string, ms = 4000) {
  const start = Date.now();
  while (!pred() && Date.now() - start < ms) await wait(15);
  if (!pred()) throw new Error(`timeout: ${label}`);
}
function client(url: string, id: string) {
  const sock = new WebSocket(url);
  const conn = new CollabConnection(createWebSocketTransport(sock as never), id);
  return conn;
}

describe("LIVE end-to-end over a real WebSocket network", () => {
  it("two real clients connect over the wire, edit, and converge", async () => {
    const hub = new CollabHub(blankProvider, new InMemoryStorage());
    const wss = new WebSocketServer({ port: 0 });
    attachWebSocketServer(wss, hub);
    await new Promise<void>((r) => wss.on("listening", () => r()));
    const port = (wss.address() as { port: number }).port;
    const url = `ws://127.0.0.1:${port}`;

    const connA = client(url, "alice");
    const connB = client(url, "bob");
    connA.join("doc-42");
    connB.join("doc-42");
    await until(() => connA.ready && connB.ready, "both ready");

    connA.submit({ kind: "insertText", at: { blockId: 1, runId: 2, offset: 0 }, text: "hello" } as never);
    await until(() => paraText(connB.doc) === "hello", "bob sees hello");
    connB.submit({ kind: "insertText", at: { blockId: 1, runId: 2, offset: 5 }, text: " world" } as never);
    await until(() => paraText(connA.doc) === "hello world", "alice sees world");

    expect(serializeXml(connA.doc!.docRoot)).toBe(serializeXml(connB.doc!.docRoot));

    const connC = client(url, "carol");
    connC.join("doc-42");
    await until(() => paraText(connC.doc) === "hello world", "carol late-joins");
    expect(serializeXml(connC.doc!.docRoot)).toBe(serializeXml(connA.doc!.docRoot));
    wss.close();
  }, 15000);
});

describe("LIVE persistence over the wire", () => {
  it("a NEW server on the same storage rehydrates the document over a real socket", async () => {
    const storage = new InMemoryStorage();
    const wait = (ms: number) => new Promise((r) => setTimeout(r, ms));
    const upd = (doc: { docRoot: XmlElement } | null): string => {
      if (!doc) return "";
      const body = doc.docRoot.children.find((c) => localName(c.name) === "body")!;
      const walk = (el: XmlElement): string => { let t = localName(el.name) === "t" ? el.text : ""; el.children.forEach((c) => (t += walk(c))); return t; };
      return walk(body);
    };
    const untilT = async (pred: () => boolean, label: string) => { const s = Date.now(); while (!pred() && Date.now() - s < 4000) await wait(15); if (!pred()) throw new Error("timeout " + label); };

    // Server 1: a client edits over the wire.
    const hub1 = new CollabHub(blankProvider, storage);
    const wss1 = new WebSocketServer({ port: 0 });
    attachWebSocketServer(wss1, hub1);
    await new Promise<void>((r) => wss1.on("listening", () => r()));
    const p1 = (wss1.address() as { port: number }).port;
    const c1 = new CollabConnection(createWebSocketTransport(new WebSocket(`ws://127.0.0.1:${p1}`) as never), "a");
    c1.join("persist");
    await untilT(() => c1.ready, "c1 ready");
    c1.submit({ kind: "insertText", at: { blockId: 1, runId: 2, offset: 0 }, text: "durable" } as never);
    await untilT(() => upd(c1.doc) === "durable", "c1 edit");
    await wait(60); // durable append settles
    wss1.close();

    // Server 2: brand-new hub on the SAME storage — rehydrates from the log.
    const hub2 = new CollabHub(blankProvider, storage);
    const wss2 = new WebSocketServer({ port: 0 });
    attachWebSocketServer(wss2, hub2);
    await new Promise<void>((r) => wss2.on("listening", () => r()));
    const p2 = (wss2.address() as { port: number }).port;
    const c2 = new CollabConnection(createWebSocketTransport(new WebSocket(`ws://127.0.0.1:${p2}`) as never), "b");
    c2.join("persist");
    await untilT(() => upd(c2.doc) === "durable", "c2 rehydrated");
    expect(upd(c2.doc)).toBe("durable");
    wss2.close();
  }, 15000);
});
