/**
 * #157: a panel opened with the MOUSE has to take the keyboard with it.
 *
 * The gallery triggers carry `onMouseDown={e => e.preventDefault()}` so the
 * editor keeps its caret and selection — deliberate, and load-bearing for
 * applying a highlight to selected text. The side effect was that clicking a
 * trigger left `document.activeElement` on the editor's hidden textarea, so
 * Tab moved FORWARD from the document, away from the bar, and the open panel
 * could not be reached by keyboard at all.
 *
 * Measured in Chromium on contract.docx, opening every panel a mouse can
 * open: 35 of 51 left activeElement on TEXTAREA. After: 0 of 51.
 *
 * The PANEL takes focus, not its first control. Focusing a control runs that
 * control's own focus handling, and the table-size grid reads focus as a
 * choice — it relabelled itself "1 x 1" the instant the panel opened, which
 * broke two of #152's tests. Tab from the container reaches the first control
 * anyway, which is all that was actually missing.
 */
import { act, createElement } from "react";
import { createRoot, type Root } from "react-dom/client";
import { afterEach, describe, expect, it } from "vitest";
import { DocxToolbar } from "../src/toolbar.js";
import type { DocxViewApi } from "../src/index.js";

let root: Root | null = null;
let host: HTMLElement | null = null;

afterEach(async () => {
  if (root) await act(async () => root!.unmount());
  host?.remove();
  root = null;
  host = null;
});

function stubApi(): DocxViewApi {
  const impl: Record<string, unknown> = {
    getSelectionFormat: () => ({ bold: false, italic: false, underline: false, strike: false }),
    listStyles: () => [],
  };
  return new Proxy({} as Record<string, unknown>, {
    get: (_t, key: string) => (key in impl ? impl[key] : () => undefined),
  }) as unknown as DocxViewApi;
}

async function mountBar(): Promise<HTMLElement> {
  host = document.createElement("div");
  document.body.appendChild(host);
  root = createRoot(host);
  await act(async () => { root!.render(createElement(DocxToolbar, { api: stubApi() })); });
  const bar = host.querySelector<HTMLElement>("[data-dxw-toolbar-mode]");
  if (!bar) throw new Error("the toolbar never rendered");
  return bar;
}

async function selectTab(bar: HTMLElement, name: string) {
  const tab = [...bar.querySelectorAll<HTMLElement>("[data-dxw-toolbar-tabs] button")]
    .find((el) => (el.textContent ?? "").trim().toLowerCase() === name);
  if (!tab) throw new Error(`no "${name}" tab`);
  await act(async () => { tab.dispatchEvent(new MouseEvent("click", { bubbles: true, cancelable: true })); });
}

function control(bar: HTMLElement, tip: string): HTMLElement {
  const found = [...bar.querySelectorAll<HTMLElement>("button")].find(
    (el) => (el.getAttribute("title") ?? el.getAttribute("data-tip") ?? "") === tip,
  );
  if (!found) throw new Error(`no "${tip}" control on this tab`);
  return found;
}

/** Open with a mouse press, exactly as the defect describes. */
async function openByMouse(bar: HTMLElement, tip: string): Promise<HTMLElement> {
  const trigger = control(bar, tip);
  const before = new Set(document.querySelectorAll("*"));
  await act(async () => {
    trigger.dispatchEvent(new MouseEvent("mousedown", { bubbles: true, cancelable: true }));
    trigger.dispatchEvent(new MouseEvent("click", { bubbles: true, cancelable: true }));
  });
  const fresh = [...document.querySelectorAll<HTMLElement>("div, span")]
    .filter((el) => !before.has(el) && el.style.position === "fixed" && el.style.zIndex !== "");
  const panel = fresh.find((el) => !fresh.some((other) => other !== el && other.contains(el)));
  if (!panel) throw new Error(`"${tip}" opened no panel`);
  return panel;
}

describe("#157 · a mouse-opened panel takes the keyboard with it", () => {
  for (const [tab, tip] of [
    ["home", "Highlight color"],
    ["home", "Text Effects and Typography"],
    ["insert", "Table"],
    ["insert", "Insert cover page"],
    ["insert", "Insert advanced symbol"],
  ] as const) {
    it(`${tip} is reachable from the keyboard once clicked`, async () => {
      const bar = await mountBar();
      await selectTab(bar, tab);
      const panel = await openByMouse(bar, tip);

      expect(
        panel.contains(document.activeElement),
        `${tip}: focus stayed outside the panel (${document.activeElement?.tagName}), so Tab leads away from it`,
      ).toBe(true);
      // Focusable on purpose, but never a stop on the way through the bar.
      expect(panel.getAttribute("tabindex"), `${tip}: the panel is in the tab order`).toBe("-1");
      // The first Tab from here has somewhere to go inside the panel.
      const reachable = panel.querySelectorAll(
        "button:not(:disabled), input:not(:disabled), select:not(:disabled), textarea:not(:disabled)",
      );
      expect(reachable.length, `${tip}: nothing inside the panel to Tab to`).toBeGreaterThan(0);
    });
  }

  it("leaves a form popover's own choice of field alone", async () => {
    const bar = await mountBar();
    await selectTab(bar, "insert");
    const panel = await openByMouse(bar, "Insert bookmark");
    // The form focuses its field in its own effect, which runs after the
    // shared one — the panel must not have kept focus for itself.
    expect(document.activeElement?.tagName, "the bookmark form did not get its field focused").toBe("INPUT");
    expect(panel.contains(document.activeElement)).toBe(true);
  });
});
