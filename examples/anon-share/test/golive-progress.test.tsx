// @vitest-environment jsdom
/**
 * Going collaborative on a big document (examples/anon-share): the pipeline
 * (serialise → seal → PUT) is real seconds of main-thread work at 500 pages,
 * so the UI must stay honest — the progress overlay is PAINTED before the
 * synchronous serialise starts (the fea7e44 two-frame lesson), the phases
 * report as they happen, a refusal surfaces with its reason, and the phase
 * timings are emitted as STRESS-METRIC lines so this stays measurable.
 */
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { createElement, act } from "react";
import { createRoot, type Root } from "react-dom/client";
import { blankDocxBytes } from "@wordinweb/server";
import { InMemoryBundleStore } from "@wordinweb/collab/client";
import { LocalEditor } from "../src/local-editor";
import { goLiveEncrypted, GoLiveError, type GoLivePhase } from "../src/e2ee-flows";

/* --------------------------- goLiveEncrypted ----------------------------- */

const prevFetch = globalThis.fetch;
afterEach(() => { globalThis.fetch = prevFetch; vi.restoreAllMocks(); });

describe("goLiveEncrypted", () => {
  it("reports encrypt → upload in order and emits the phase timings as a STRESS-METRIC line", async () => {
    globalThis.fetch = vi.fn(async () =>
      new Response(JSON.stringify({ genesisId: "g_x", ownerToken: "tok" }), { status: 201 }),
    ) as unknown as typeof fetch;
    const log = vi.spyOn(console, "log");
    const phases: GoLivePhase[] = [];

    const out = await goLiveEncrypted("http://seedhost", blankDocxBytes(), "redwood", (p) => phases.push(p));

    expect(phases).toEqual(["encrypt", "upload"]);
    expect(out.docId).toMatch(/^d_/);
    expect(out.ownerToken).toBe("tok");
    const metric = log.mock.calls.map((c) => String(c[0])).find((l) => l.includes("STRESS-METRIC golive-seal"));
    expect(metric, "the phase timings must be emitted for perf-report ingestion").toBeTruthy();
    for (const key of ["docxBytes=", "stretchMs=", "parseMs=", "hashMs=", "sealMs=", "uploadMs=", "totalMs="]) {
      expect(metric).toContain(key);
    }
  });

  it("surfaces a 413 as a plain too-large error instead of joining a refused session", async () => {
    globalThis.fetch = vi.fn(async () =>
      new Response(JSON.stringify({ error: "too-large", maxBytes: 10 * 1024 * 1024 }), { status: 413 }),
    ) as unknown as typeof fetch;
    const err = await goLiveEncrypted("http://seedhost", blankDocxBytes(), "redwood").then(
      () => null,
      (e: unknown) => e,
    );
    expect(err, "a refused seed must reject, not resolve").toBeInstanceOf(GoLiveError);
    expect((err as GoLiveError).status).toBe(413);
    expect((err as GoLiveError).message).toContain("too large");
  });
});

/* ------------------------------ LocalEditor ------------------------------ */

let mounted: { root: Root; host: HTMLElement }[] = [];
function render(node: ReturnType<typeof createElement>) {
  const host = document.createElement("div");
  document.body.appendChild(host);
  const root = createRoot(host);
  act(() => { root.render(node); });
  mounted.push({ root, host });
  return host;
}
afterEach(() => {
  for (const { root, host } of mounted) {
    act(() => { root.unmount(); });
    host.remove();
  }
  mounted = [];
});

beforeEach(() => {
  globalThis.fetch = vi.fn(async () =>
    new Response(blankDocxBytes() as unknown as BodyInit, { status: 200 }),
  ) as unknown as typeof fetch;
});

const byId = (host: HTMLElement, id: string) => host.querySelector<HTMLElement>(`[data-testid="${id}"]`);
function click(el: HTMLElement | null) {
  expect(el).toBeTruthy();
  act(() => { el!.dispatchEvent(new MouseEvent("click", { bubbles: true })); });
}
async function tick(ms = 5) { await act(async () => { await new Promise<void>((r) => setTimeout(r, ms)); }); }

function setCode(host: HTMLElement, value: string) {
  const input = byId(host, "share-code") as HTMLInputElement;
  expect(input).toBeTruthy();
  act(() => {
    Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, "value")!.set!.call(input, value);
    input.dispatchEvent(new Event("input", { bubbles: true }));
  });
}

/** Mount, wait for the editor api, open the modal, enter a code. */
async function readyToStart(host: HTMLElement) {
  const cta = () => byId(host, "make-collaborative") as HTMLButtonElement | null;
  for (let i = 0; i < 200 && cta()?.disabled !== false; i++) await tick();
  expect(cta()?.disabled, "the editor never became ready").toBe(false);
  click(cta());
  setCode(host, "redwood");
}

describe("LocalEditor go-live", () => {
  it("paints the progress overlay BEFORE the serialise/onGoLive work starts", async () => {
    // Records what was on screen at the moment the pipeline was entered —
    // the regression this pins is state-then-sync-work in one tick, where
    // the overlay is mounted and unmounted without ever being painted.
    let overlayUpAtCall: boolean | null = null;
    let phaseTextAtCall: string | null = null;
    const onGoLive = vi.fn((_b: Uint8Array, _c?: string, _p?: (p: GoLivePhase) => void) => {
      overlayUpAtCall = !!document.querySelector('[data-testid="golive-progress"]');
      phaseTextAtCall = document.querySelector('[data-testid="golive-phase"]')?.textContent ?? null;
      return new Promise<void>(() => {}); // never settles — progress stays up
    });
    const host = render(createElement(LocalEditor, {
      httpBase: "http://blankhost", onGoLive, store: new InMemoryBundleStore(), autosaveMs: 100000,
    }));
    await readyToStart(host);
    click(byId(host, "start-collab"));
    for (let i = 0; i < 100 && onGoLive.mock.calls.length === 0; i++) await tick();
    expect(onGoLive).toHaveBeenCalled();
    expect(overlayUpAtCall, "the overlay must be committed before the heavy work begins").toBe(true);
    expect(phaseTextAtCall).toContain("Preparing");
  });

  it("moves the overlay text as onGoLive reports phases", async () => {
    let report: ((p: GoLivePhase) => void) | undefined;
    const onGoLive = vi.fn((_b: Uint8Array, _c?: string, p?: (p: GoLivePhase) => void) => {
      report = p;
      return new Promise<void>(() => {});
    });
    const host = render(createElement(LocalEditor, {
      httpBase: "http://blankhost", onGoLive, store: new InMemoryBundleStore(), autosaveMs: 100000,
    }));
    await readyToStart(host);
    click(byId(host, "start-collab"));
    for (let i = 0; i < 100 && !report; i++) await tick();
    act(() => report!("encrypt"));
    expect(byId(host, "golive-phase")!.textContent).toContain("Encrypting");
    act(() => report!("upload"));
    expect(byId(host, "golive-phase")!.textContent).toContain("Uploading");
  });

  it("puts a refusal's reason on screen and hands the editor back", async () => {
    const onGoLive = vi.fn(async () => {
      throw new GoLiveError("This document is too large to share live (the limit is about 5 MB of document).", 413);
    });
    const host = render(createElement(LocalEditor, {
      httpBase: "http://blankhost", onGoLive, store: new InMemoryBundleStore(), autosaveMs: 100000,
    }));
    await readyToStart(host);
    click(byId(host, "start-collab"));
    for (let i = 0; i < 100 && !byId(host, "golive-error"); i++) await tick();
    expect(byId(host, "golive-error"), "the refusal's reason must surface").toBeTruthy();
    expect(byId(host, "golive-error")!.textContent).toContain("too large");
    expect(byId(host, "golive-progress"), "the progress overlay must come down").toBeNull();
    expect((byId(host, "start-collab") as HTMLButtonElement).disabled, "retry must be possible").toBe(false);
  });
});
