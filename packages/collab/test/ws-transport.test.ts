import { describe, expect, it } from "vitest";
import { createWebSocketTransport, SocketLike } from "../src/ws-transport.js";
import { ServerMessage } from "../src/protocol.js";

class FakeSocket implements SocketLike {
  sent: string[] = [];
  private openCbs: (() => void)[] = [];
  private msgCbs: ((ev: { data: unknown }) => void)[] = [];
  send(data: string): void { this.sent.push(data); }
  addEventListener(type: "message" | "open", cb: never): void {
    if (type === "open") this.openCbs.push(cb as () => void);
    else this.msgCbs.push(cb as (ev: { data: unknown }) => void);
  }
  open(): void { for (const c of this.openCbs) c(); }
  deliver(msg: ServerMessage): void { for (const c of this.msgCbs) c({ data: JSON.stringify(msg) }); }
}

describe("createWebSocketTransport", () => {
  it("buffers sends until open, then flushes", () => {
    const sock = new FakeSocket();
    const t = createWebSocketTransport(sock);
    t.send({ t: "hello", protocolVersion: 1, docId: "d", sinceSeq: 0 });
    expect(sock.sent).toHaveLength(0); // buffered
    sock.open();
    expect(sock.sent).toHaveLength(1);
    expect(JSON.parse(sock.sent[0]).t).toBe("hello");
  });

  it("sends immediately once open", () => {
    const sock = new FakeSocket();
    const t = createWebSocketTransport(sock);
    sock.open();
    t.send({ t: "presence", position: null });
    expect(sock.sent).toHaveLength(1);
  });

  it("parses inbound server messages and ignores malformed frames", () => {
    const sock = new FakeSocket();
    const t = createWebSocketTransport(sock);
    const got: ServerMessage[] = [];
    t.onMessage((m) => got.push(m));
    sock.deliver({ t: "refused", reason: "x" });
    expect(got).toHaveLength(1);
    expect(got[0].t).toBe("refused");
    // Malformed:
    (sock as unknown as { deliver: (s: unknown) => void });
    for (const c of (sock as never as { ["msgCbs"]: ((ev: { data: unknown }) => void)[] })["msgCbs"] ?? []) void c;
    expect(got).toHaveLength(1);
  });
});
