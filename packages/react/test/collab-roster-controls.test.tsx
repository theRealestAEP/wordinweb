// @vitest-environment jsdom
import { afterEach, describe, expect, it } from "vitest";
import { createElement, act } from "react";
import { createRoot, type Root } from "react-dom/client";
import { App } from "../../../examples/anon-share/src/app";
import { CollabHub, blankDocxBytes, type Connection, type ServerMessage } from "@wordinweb/server";
import { CollabConnection, createWebSocketTransport, InMemoryBundleStore } from "@wordinweb/collab/client";

let mounted: { root: Root; host: HTMLElement }[] = [];
const previousSocket = globalThis.WebSocket;

afterEach(() => {
  for (const { root, host } of mounted) {
    act(() => { root.unmount(); });
    host.remove();
  }
  mounted = [];
  (globalThis as { WebSocket: unknown }).WebSocket = previousSocket;
});

async function tick(ms = 5) {
  await act(async () => { await new Promise<void>((resolve) => setTimeout(resolve, ms)); });
}

async function until(check: () => boolean, label: string) {
  for (let i = 0; i < 200; i++) {
    if (check()) return;
    await tick();
  }
  throw new Error(`timeout: ${label}`);
}

function socketClass(hub: CollabHub, sockets: { close: () => void }[]) {
  let sequence = 0;
  return class HubSocket {
    onclose: (() => void) | null = null;
    onerror: (() => void) | null = null;
    private closed = false;
    private listeners: ((event: { data: unknown }) => void)[] = [];
    private conn: Connection;

    constructor(_url: string) {
      this.conn = {
        id: `roster-control-${sequence++}`,
        send: (message: ServerMessage) => {
          setTimeout(() => this.listeners.forEach((listener) => listener({ data: JSON.stringify(message) })), 1);
        },
        close: () => this.close(),
      };
      sockets.push(this);
    }

    send(data: string) {
      setTimeout(() => { void hub.handle(this.conn, JSON.parse(data)); }, 1);
    }

    close = () => {
      if (this.closed) return;
      this.closed = true;
      hub.disconnect(this.conn);
      this.onclose?.();
    };

    addEventListener(type: string, callback: never) {
      if (type === "message") this.listeners.push(callback as never);
      if (type === "open") (callback as () => void)();
    }
  };
}

function click(element: Element) {
  act(() => { element.dispatchEvent(new MouseEvent("click", { bubbles: true })); });
}

describe("collaborator controls", () => {
  it("toggles view-only, explains both actions, and removes inactive participants", async () => {
    const hub = new CollabHub(null);
    const seeded = hub.seed("roster-controls", blankDocxBytes());
    if (!seeded.ok) throw new Error("seed failed");

    const sockets: { close: () => void }[] = [];
    const Socket = socketClass(hub, sockets);
    (globalThis as { WebSocket: unknown }).WebSocket = Socket;

    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);
    await act(async () => {
      root.render(createElement(App, {
        url: "ws://loopback/collab",
        httpBase: "http://loopback",
        docId: "roster-controls",
        clientId: "owner",
        name: "Owner",
        ownerToken: seeded.ownerToken,
        store: new InMemoryBundleStore(),
      }));
    });
    mounted.push({ root, host });

    const bea = new CollabConnection(
      createWebSocketTransport(new Socket("ws://peer") as never),
      "bea",
    );
    bea.join("roster-controls", undefined, { profile: { name: "Bea", color: "" } });

    await until(() => host.querySelectorAll('[data-testid="roster-chip"]').length === 2, "Bea joins");
    const peerChip = () => [...host.querySelectorAll<HTMLElement>('[data-testid="roster-chip"]')]
      .find((entry) => entry.textContent?.includes("Bea"));
    const roleButton = () => peerChip()?.querySelector<HTMLButtonElement>('[data-testid="roster-role"]');

    expect(roleButton()?.getAttribute("aria-pressed")).toBe("false");
    expect(roleButton()?.title).toContain("Make Bea view-only");
    act(() => { roleButton()!.dispatchEvent(new MouseEvent("mouseover", { bubbles: true })); });
    expect(host.querySelector('[data-testid="control-tooltip"]')?.textContent).toContain("They can still read and download");

    click(roleButton()!);
    await until(() => roleButton()?.getAttribute("aria-pressed") === "true", "Bea becomes view-only");
    expect(roleButton()?.title).toContain("Allow Bea to edit");

    click(roleButton()!);
    await until(() => roleButton()?.getAttribute("aria-pressed") === "false", "Bea can edit again");

    const kickButton = peerChip()!.querySelector<HTMLButtonElement>('[data-testid="roster-kick"]')!;
    expect(kickButton.title).toContain("Remove Bea from this session");
    click(kickButton);
    await until(() => !peerChip(), "kicked participant disappears");

    const cal = new CollabConnection(
      createWebSocketTransport(new Socket("ws://peer-2") as never),
      "cal",
    );
    cal.join("roster-controls", undefined, { profile: { name: "Cal", color: "" } });
    await until(() => host.textContent?.includes("Cal") ?? false, "Cal joins");
    await until(() => cal.ready, "Cal receives the document");
    cal.setPresence({ anchor: { blockId: 1, runId: 2, offset: 0 } });
    await until(
      () => !!host.querySelector('.dxw-presence-caret[data-participant="cal"]'),
      "Cal's caret appears",
    );
    sockets[sockets.length - 1].close();
    await until(
      () => ![...host.querySelectorAll<HTMLElement>('[data-testid="roster-chip"]')]
        .some((entry) => entry.textContent?.includes("Cal")),
      "closed browser disappears from the active roster",
    );
    await until(
      () => !host.querySelector('.dxw-presence-caret[data-participant="cal"]'),
      "closed browser's caret disappears",
    );
  });

  it("gives a removed participant a useful local-copy action", async () => {
    const hub = new CollabHub(null);
    const seeded = hub.seed("kicked-copy", blankDocxBytes());
    if (!seeded.ok) throw new Error("seed failed");

    const sockets: { close: () => void }[] = [];
    const Socket = socketClass(hub, sockets);
    (globalThis as { WebSocket: unknown }).WebSocket = Socket;
    const mount = async (props: Parameters<typeof App>[0]) => {
      const host = document.createElement("div");
      document.body.appendChild(host);
      const root = createRoot(host);
      await act(async () => { root.render(createElement(App, props)); });
      mounted.push({ root, host });
      return host;
    };
    const owner = await mount({
      url: "ws://loopback/collab",
      httpBase: "http://loopback",
      docId: "kicked-copy",
      clientId: "owner",
      name: "Owner",
      ownerToken: seeded.ownerToken,
      store: new InMemoryBundleStore(),
    });
    let localCopy: Uint8Array | null | undefined;
    const peer = await mount({
      url: "ws://loopback/collab",
      httpBase: "http://loopback",
      docId: "kicked-copy",
      clientId: "peer",
      name: "Peer",
      onDisconnect: (bytes) => { localCopy = bytes; },
      store: new InMemoryBundleStore(),
    });

    await until(() => owner.querySelectorAll('[data-testid="roster-chip"]').length === 2, "peer joins");
    const peerChip = [...owner.querySelectorAll<HTMLElement>('[data-testid="roster-chip"]')]
      .find((entry) => entry.textContent?.includes("Peer"));
    click(peerChip!.querySelector('[data-testid="roster-kick"]')!);

    await until(() => !!peer.querySelector('[data-testid="kicked-keep-local"]'), "removed screen appears");
    expect(peer.textContent).toContain("You were removed from this session");
    expect(peer.textContent).not.toContain("server refused");
    expect(peer.textContent).not.toContain("Reload this tab");

    click(peer.querySelector('[data-testid="kicked-keep-local"]')!);
    expect(localCopy).toBeInstanceOf(Uint8Array);
    expect(localCopy!.byteLength).toBeGreaterThan(0);
  });
});
