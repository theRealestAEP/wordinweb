// @vitest-environment jsdom
/**
 * The shared-room disconnect boundary.
 *
 * Disconnect offers an explicit room rejoin or a linked offline workspace.
 * The room lock applies to the owner and peers while the admin control stays live.
 */
import { afterEach, describe, expect, it } from "vitest";
import { createElement, act } from "react";
import { createRoot, type Root } from "react-dom/client";
import { CollabHub, blankDocxBytes, type Connection, type ServerMessage } from "@wordinweb/server";
import { CollabConnection, createWebSocketTransport, InMemoryBundleStore } from "@wordinweb/collab/client";
import { App } from "../src/app";
import type { CollabSession } from "../../../packages/react/src/collab";

let mounted: { root: Root; host: HTMLElement }[] = [];
const previousSocket = globalThis.WebSocket;

afterEach(() => {
  for (const { root, host } of mounted) {
    act(() => root.unmount());
    host.remove();
  }
  mounted = [];
  (globalThis as { WebSocket: unknown }).WebSocket = previousSocket;
  delete (window as unknown as { __ww?: unknown }).__ww;
  localStorage.removeItem("wordinweb-offline-offline-rejoin");
});

async function tick(ms = 5) {
  await act(async () => { await new Promise<void>((resolve) => setTimeout(resolve, ms)); });
}

async function until(check: () => boolean, label: string, turns = 400) {
  for (let i = 0; i < turns; i++) {
    if (check()) return;
    await tick();
  }
  throw new Error(`timeout: ${label}`);
}

async function untilAsync(check: () => Promise<boolean>, label: string, turns = 400) {
  for (let i = 0; i < turns; i++) {
    if (await check()) return;
    await tick();
  }
  throw new Error(`timeout: ${label}`);
}

function socketClass(hub: CollabHub) {
  let sequence = 0;
  return class HubSocket {
    onclose: (() => void) | null = null;
    onerror: (() => void) | null = null;
    private closed = false;
    private listeners: ((event: { data: unknown }) => void)[] = [];
    private conn: Connection;

    constructor(_url: string) {
      this.conn = {
        id: `offline-rejoin-${sequence++}`,
        send: (message: ServerMessage) => {
          setTimeout(() => this.listeners.forEach((listener) => listener({ data: JSON.stringify(message) })), 1);
        },
        close: () => this.close(),
      };
    }

    send(data: string) {
      const message = JSON.parse(data);
      setTimeout(() => { void hub.handle(this.conn, message); }, 1);
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
  act(() => element.dispatchEvent(new MouseEvent("click", { bubbles: true })));
}

function enterValue(element: HTMLInputElement, value: string) {
  act(() => {
    const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, "value")!.set!;
    setter.call(element, value);
    element.dispatchEvent(new Event("input", { bubbles: true }));
  });
}

async function placeCaret(host: HTMLElement) {
  const page = host.querySelector<HTMLElement>(".dxw-page")!;
  const target = page.querySelector("span") ?? page;
  await act(async () => {
    const event = { bubbles: true, cancelable: true, clientX: 10, clientY: 10, button: 0 };
    target.dispatchEvent(new MouseEvent("mousedown", event));
    target.dispatchEvent(new MouseEvent("mouseup", event));
  });
  await tick();
}

async function keys(host: HTMLElement, values: string[]) {
  const target = host.querySelector<HTMLElement>("textarea") ?? host;
  await act(async () => {
    for (const key of values) {
      target.dispatchEvent(new KeyboardEvent("keydown", { key, bubbles: true, cancelable: true }));
      await new Promise((resolve) => setTimeout(resolve, 2));
    }
  });
}

function session(): CollabSession {
  return (window as unknown as { __ww: { _session: CollabSession } }).__ww._session;
}

function textOf(doc: { docRoot: { name: string; text: string; children: unknown[] } }): string {
  const walk = (element: { name: string; text: string; children: unknown[] }): string =>
    (element.name.endsWith(":t") ? element.text : "") + (element.children as never[]).map(walk).join("");
  return walk(doc.docRoot);
}

describe("disconnect choice", () => {
  it("keeps the stable room identity and checks offline changes before merge", async () => {
    const hub = new CollabHub(null);
    const seeded = hub.seed("offline-rejoin", blankDocxBytes());
    if (!seeded.ok) throw new Error("seed failed");

    const Socket = socketClass(hub);
    (globalThis as { WebSocket: unknown }).WebSocket = Socket;
    const store = new InMemoryBundleStore();
    let localBytes: Uint8Array | null | undefined;
    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);
    await act(async () => {
      root.render(createElement(App, {
        url: "ws://loopback/collab",
        httpBase: "http://loopback",
        docId: "offline-rejoin",
        clientId: "owner",
        name: "Owner",
        ownerToken: seeded.ownerToken,
        store,
        onShareLink: () => {},
        onDisconnect: (bytes) => { localBytes = bytes; },
      }));
    });
    mounted.push({ root, host });

    const peer = new CollabConnection(createWebSocketTransport(new Socket("ws://peer") as never), "peer");
    peer.join("offline-rejoin", undefined, { profile: { name: "Peer", color: "" } });
    await until(() => !!host.querySelector(".dxw-page") && peer.ready, "both clients join");
    await placeCaret(host);
    await keys(host, [..."Client brief"]);
    await until(() => textOf(peer.doc!).includes("Client brief"), "title source reaches the room");
    peer.setPresence({ anchor: { blockId: 1, runId: 2, offset: 0 } });
    await until(
      () => !!host.querySelector('.dxw-presence-caret[data-participant="peer"]'),
      "peer cursor appears while the room is live",
    );

    click(host.querySelector('[data-testid="disconnect"]')!);
    await until(() => session().offline !== null, "owner enters intentional offline mode");
    await until(() => !!host.querySelector('[data-testid="disconnect-modal"]'), "disconnect choice opens");
    await until(
      () => !host.querySelector('.dxw-presence-caret[data-participant="peer"]'),
      "peer cursor clears while offline",
    );
    expect(Object.keys(session().presence)).toHaveLength(0);
    expect(host.querySelector('[data-testid="disconnect"]')?.textContent).toBe("Rejoin");
    expect(host.querySelector('[data-testid="edit-offline"]')?.textContent).toBe("Edit offline");
    expect(host.querySelector('[data-testid="disconnect-modal"]')?.textContent).toContain("Continue editing offline");
    expect(host.querySelector('[data-testid="disconnect-modal"]')?.textContent).toContain("stay linked to this room");
    expect(host.querySelector<HTMLInputElement>('[data-testid="offline-document-title"]')?.value).toBe("Client brief");
    expect(localBytes, "Disconnect waits for a choice").toBeUndefined();
    expect(host.querySelector('[data-testid="toolbar"]')?.getAttribute("data-room-state")).toBe("offline");
    expect(host.querySelector('[data-testid="offline-roster"]')?.textContent).toContain("unavailable while offline");
    expect(host.querySelectorAll('[data-testid="roster-chip"]')).toHaveLength(0);
    expect(host.querySelector('[data-testid="readonly-toggle"]')).toBeNull();
    expect(host.querySelector('[data-testid="share-collab"]')?.textContent).toBe("Copy room link");

    click(host.querySelector('[data-testid="disconnect-rejoin"]')!);
    await until(() => session().connection === "live" && session().offline === null, "owner rejoins the room");
    await until(() => !host.querySelector('[data-testid="disconnect-modal"]'), "choice closes after rejoin");
    expect(host.querySelector('[data-testid="readonly-toggle"]'), "owner controls survive rejoin").toBeTruthy();

    click(host.querySelector('[data-testid="readonly-toggle"]')!);
    await until(
      () => session().writesBlocked && session().writeStatus === "owner-lock",
      "the room lock reaches the owner editor",
    );
    await until(() => peer.roster.find((entry) => entry.clientId === "peer")?.write === "owner-lock", "the room lock reaches the peer");
    expect(host.querySelector('[data-testid="readonly-toggle"]')?.getAttribute("aria-pressed")).toBe("true");
    expect(host.querySelector('[data-testid="suggest-toggle"]')?.textContent).toBe("View only");
    expect(host.querySelector('[data-testid="readonly-banner"]')?.textContent).toContain("paused for everyone");

    const before = textOf(session().doc!);
    await keys(host, ["B", "L", "O", "C", "K", "E", "D"]);
    expect(textOf(session().doc!)).toBe(before);

    click(host.querySelector('[data-testid="readonly-toggle"]')!);
    await until(() => !session().writesBlocked && session().writeStatus === "allowed", "owner lifts the room lock");

    click(host.querySelector('[data-testid="disconnect"]')!);
    await until(() => !!host.querySelector('[data-testid="disconnect-modal"]'), "disconnect choice opens again");
    const title = host.querySelector<HTMLInputElement>('[data-testid="offline-document-title"]')!;
    expect(title).toBeTruthy();
    expect(title.value).toBe("Client brief");
    enterValue(title, "");
    expect(host.querySelector<HTMLButtonElement>('[data-testid="disconnect-continue-offline"]')!.disabled).toBe(true);
    enterValue(title, "Case filing");
    click(host.querySelector('[data-testid="disconnect-continue-offline"]')!);
    await until(() => !!host.querySelector('[data-testid="offline-workspace-status"]'), "local-looking offline workspace opens");
    await until(() => localStorage.getItem("wordinweb-offline-offline-rejoin") === "1", "offline choice persists");
    expect((await store.get("offline-rejoin"))?.title).toBe("Case filing");
    expect(host.querySelector('[data-testid="file-menu"]')).toBeTruthy();
    expect(host.querySelector('[data-testid="saved-documents"]')).toBeTruthy();
    expect(host.querySelector('[data-testid="offline-workspace-rejoin"]')?.textContent).toBe("Rejoin room");
    expect(localBytes, "offline editing keeps the room identity and stable URL").toBeUndefined();

    await placeCaret(host);
    await keys(host, ["m", "i", "n", "e"]);
    expect(session().offline?.editsHeld).toBe(1);
    expect(host.querySelector('[data-testid="offline-workspace-status"]')?.textContent).toContain("1 saved change");

    await untilAsync(
      async () => Boolean((await store.get("offline-rejoin"))?.offlineTail?.length),
      "offline edit persists before refresh",
    );
    await act(async () => { root.render(null); });
    await tick(20);
    await act(async () => {
      root.render(createElement(App, {
        url: "ws://loopback/collab",
        httpBase: "http://loopback",
        docId: "offline-rejoin",
        clientId: "owner",
        name: "Owner",
        ownerToken: seeded.ownerToken,
        store,
        onShareLink: () => {},
        onDisconnect: (bytes) => { localBytes = bytes; },
      }));
    });
    await until(
      () => session().connection === "lost" && !!session().doc && textOf(session().doc!).includes("mine") &&
        !!host.querySelector('[data-testid="offline-workspace-status"]'),
      "refresh restores the offline workspace",
    );
    expect(textOf(session().doc!)).toContain("mine");
    expect(host.querySelector('[data-testid="disconnect-modal"]')).toBeNull();

    click(host.querySelector('[data-testid="offline-workspace-rejoin"]')!);
    await until(() => !!host.querySelector('[data-testid="arrival-banner"]'), "same room offers an explicit merge check");
    expect(host.querySelector('[data-testid="arrival-banner"]')?.textContent).toContain("no newer changes");
    click(host.querySelector('[data-testid="merge-offline-changes"]')!);
    await until(() => textOf(peer.doc!).includes("mine"), "same-room offline text merges through the room");

    await placeCaret(host);
    click(host.querySelector('[data-testid="disconnect"]')!);
    await until(() => !!host.querySelector('[data-testid="disconnect-modal"]'), "disconnect choice opens for a changed-room check");
    click(host.querySelector('[data-testid="disconnect-continue-offline"]')!);
    await until(() => !host.querySelector('[data-testid="disconnect-modal"]'), "offline title saves before editing resumes");
    await placeCaret(host);
    await keys(host, ["l", "o", "c", "a", "l"]);
    expect(session().offline?.editsHeld).toBe(1);
    peer.submit({ kind: "insertText", at: { blockId: 1, runId: 2, offset: 0 }, text: "ROOM " });
    await until(() => peer.pendingCount === 0 && textOf(peer.doc!).includes("ROOM"), "the room changes while the owner edits offline");

    click(host.querySelector('[data-testid="offline-workspace-rejoin"]')!);
    await until(() => !!host.querySelector('[data-testid="arrival-banner"]'), "changed room offers reconciliation");
    expect(host.querySelector('[data-testid="arrival-banner"]')?.textContent).toContain("Add my changes as suggestions");
    click([...host.querySelectorAll("button")].find((button) => button.textContent === "Add my changes as suggestions")!);
    await until(() => textOf(peer.doc!).includes("local"), "changed-room offline text reconciles as suggestions");
  });
});
