/**
 * #160: the run-format toggles must alternate on EVERY click.
 *
 * They read the cached `fmt` to decide what to send, and `fmt` was refreshed
 * from a read taken immediately after `applyFormat`. The engine's model does
 * not update synchronously — measured in Chromium on
 * wild2-med-phase23-protocol, `getSelectionFormat()` still reported the old
 * value at 0ms, at a microtask, at two animation frames and at 50ms, and only
 * told the truth by 100ms. So the bar cached the PREVIOUS answer and the next
 * click asked for the state the text was already in:
 *
 *     bold spans: 49 -> 249 -> 249 -> 49 -> 49
 *
 * Clicks 2 and 4 did nothing, so turning bold off took three clicks. ⌘B was
 * always fine, because the keyboard path reads the format live.
 *
 * WHY FOUR CLICKS. A two-click test passes against this bug — the first click
 * works and the second is where it goes wrong only if you check that the state
 * came BACK. That is why it shipped in every build we have made.
 *
 * The stub below makes `applyFormat` land one macrotask later, which is the
 * minimal honest encoding of "the engine is not synchronous". The toolbar must
 * not depend on reading back its own write.
 */
import { act, createElement } from "react";
import { createRoot, type Root } from "react-dom/client";
import { afterEach, describe, expect, it } from "vitest";
import { DocxToolbar } from "../src/toolbar.js";
import type { DocxViewApi } from "../src/index.js";

type Live = {
  bold: boolean; italic: boolean; underline: boolean; strike: boolean;
  verticalAlign?: "superscript" | "subscript";
};

interface Recorder {
  api: DocxViewApi;
  /** The patch each click asked for, in order. */
  asked: Record<string, unknown>[];
  /** What the document actually holds, after the engine catches up. */
  state: () => Live;
}

/**
 * An engine that answers with what it held BEFORE the pending write lands —
 * the behaviour measured in the browser, reduced to its essentials.
 */
function recordingApi(): Recorder {
  const visible: Live = { bold: false, italic: false, underline: false, strike: false };
  const asked: Record<string, unknown>[] = [];
  const impl: Record<string, unknown> = {
    getSelectionFormat: () => ({ ...visible }),
    applyFormat: (patch: Record<string, unknown>) => {
      asked.push(patch);
      // Lands later, exactly like the real one: a read taken now still sees
      // the old value.
      setTimeout(() => {
        for (const [key, value] of Object.entries(patch)) {
          if (key === "verticalAlign") visible.verticalAlign = (value as Live["verticalAlign"]) ?? undefined;
          else (visible as Record<string, unknown>)[key] = value;
        }
      }, 0);
    },
  };
  const api = new Proxy({} as Record<string, unknown>, {
    get: (_t, key: string) => (key in impl ? impl[key] : () => undefined),
  }) as unknown as DocxViewApi;
  return { api, asked, state: () => ({ ...visible }) };
}

let root: Root | null = null;
let host: HTMLElement | null = null;

afterEach(async () => {
  if (root) await act(async () => root!.unmount());
  host?.remove();
  root = null;
  host = null;
});

async function mountBar(api: DocxViewApi): Promise<HTMLElement> {
  host = document.createElement("div");
  document.body.appendChild(host);
  root = createRoot(host);
  await act(async () => { root!.render(createElement(DocxToolbar, { api })); });
  return host;
}

function control(bar: HTMLElement, tip: string): HTMLElement {
  const found = [...bar.querySelectorAll<HTMLElement>("button")].find(
    // The bar moves `title` to `data-tip` on first hover, so both must match.
    (el) => (el.getAttribute("title") ?? el.getAttribute("data-tip") ?? "").startsWith(tip),
  );
  if (!found) throw new Error(`no "${tip}" control on the bar`);
  return found;
}

/** Click, then let the engine's pending write land, like a real user pausing. */
async function clickAndSettle(el: HTMLElement) {
  await act(async () => {
    el.dispatchEvent(new MouseEvent("mousedown", { bubbles: true, cancelable: true }));
    el.dispatchEvent(new MouseEvent("click", { bubbles: true, cancelable: true }));
  });
  await act(async () => { await new Promise((r) => setTimeout(r, 5)); });
}

describe("#160 · a format toggle alternates on every click", () => {
  /**
   * Every toggle that decides its next value from the selection's current
   * format. The audit for #160 found exactly these six; the other controls
   * that read `fmt` (font, size, colour, highlight, character style) send an
   * absolute value the user chose, so no phase can build up in them.
   */
  const TOGGLES = [
    ["Bold", (live: Live) => live.bold, [true, false, true, false]],
    ["Italic", (live: Live) => live.italic, [true, false, true, false]],
    ["Underline", (live: Live) => live.underline, [true, false, true, false]],
    ["Strikethrough", (live: Live) => live.strike, [true, false, true, false]],
    ["Superscript", (live: Live) => live.verticalAlign === "superscript", [true, false, true, false]],
    ["Subscript", (live: Live) => live.verticalAlign === "subscript", [true, false, true, false]],
  ] as const;

  for (const [tip, read, want] of TOGGLES) {
    it(`${tip} goes on, off, on, off over four clicks`, async () => {
      const engine = recordingApi();
      const bar = await mountBar(engine.api);
      const button = control(bar, tip);

      const seen: boolean[] = [];
      for (let i = 0; i < 4; i++) {
        await clickAndSettle(button);
        seen.push(read(engine.state()));
      }

      // Was [on, on, off, off]: clicks 2 and 4 asked for the state the text
      // was already in.
      expect(seen, `${tip} did not alternate`).toEqual([...want]);
      expect(engine.asked.length, `${tip} did not send four patches`).toBe(4);
    });
  }

  it("keeps the button's own highlight in step with the click", async () => {
    const engine = recordingApi();
    const bar = await mountBar(engine.api);
    const button = control(bar, "Bold");
    const lit = () => button.style.background !== "transparent" && button.style.background !== "";

    expect(lit(), "the bar started with Bold lit").toBe(false);
    await clickAndSettle(button);
    // The engine cannot answer yet at click time, so the bar has to show what
    // was asked for; otherwise the user gets no feedback on the click that
    // just worked.
    expect(lit(), "Bold applied but the button did not light").toBe(true);
    await clickAndSettle(button);
    expect(lit(), "Bold removed but the button stayed lit").toBe(false);
  });
});
