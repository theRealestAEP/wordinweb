/**
 * #146: shared machinery for the popover smoke suite.
 *
 * The toolbar has one popover or dialog per feature and, before this, four of
 * them had a test. The rest could be shipped broken — a panel that crashes on
 * open, or opens empty, or traps the user because nothing dismisses it — and
 * every suite in the package would still be green. This module supplies the
 * parts that let ONE table drive all of them:
 *
 *  - `mountToolbar` — a real `DocxView` behind a real `DocxToolbar`, because
 *    most of these panels read the engine when they open and a hand-written
 *    api stub is a second implementation to keep in step.
 *  - `openSurface` / `panelOf` — find the panel a control opened WITHOUT the
 *    panel having to declare itself. Nothing in the file marks "I am a
 *    popover", so a table listing per-surface selectors would go stale the
 *    moment someone adds a panel. Instead: photograph the document, click,
 *    and the new subtree IS the panel.
 *  - `watchConsole` — a popover that throws inside a React event handler does
 *    not fail a test by itself; React logs and carries on. This turns that
 *    log into a failure.
 */
import { act, createElement } from "react";
import { createRoot, type Root } from "react-dom/client";
import { zipSync, strToU8 } from "fflate";
import { DocxView, type DocxViewApi } from "../src/index.js";
import { DocxToolbar } from "../src/toolbar.js";

function docx(body: string): Uint8Array {
  return zipSync({
    "[Content_Types].xml": strToU8(
      `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`,
    ),
    "_rels/.rels": strToU8(
      `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`,
    ),
    "word/document.xml": strToU8(
      `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
        `<w:body>${body}</w:body></w:document>`,
    ),
  });
}

/** One paragraph is enough: these tests are about the bar, not the document. */
export const FIXTURE = docx(
  `<w:p><w:r><w:t xml:space="preserve">Body text for the smoke suite</w:t></w:r></w:p>`,
);

/** A 2×2 table, for the surfaces that only exist with the caret inside one. */
export const TABLE_FIXTURE = (() => {
  const cell = (text: string) =>
    `<w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr><w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p></w:tc>`;
  return docx(
    `<w:p><w:r><w:t xml:space="preserve">Before</w:t></w:r></w:p>` +
      `<w:tbl><w:tblPr><w:tblW w:w="4800" w:type="dxa"/></w:tblPr>` +
      `<w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid>` +
      `<w:tr>${cell("r0c0")}${cell("r0c1")}</w:tr>` +
      `<w:tr>${cell("r1c0")}${cell("r1c1")}</w:tr></w:tbl>` +
      `<w:p><w:r><w:t xml:space="preserve">After</w:t></w:r></w:p>`,
  );
})();

export async function tick(ms = 5) {
  await act(async () => {
    await new Promise<void>((resolve) => setTimeout(resolve, ms));
  });
}

export async function click(element: HTMLElement) {
  await act(async () => {
    element.dispatchEvent(new MouseEvent("mousedown", { bubbles: true, cancelable: true }));
    element.dispatchEvent(new MouseEvent("mouseup", { bubbles: true, cancelable: true }));
    element.dispatchEvent(new MouseEvent("click", { bubbles: true, cancelable: true }));
  });
}

export async function pressKey(target: EventTarget, key: string) {
  await act(async () => {
    target.dispatchEvent(new KeyboardEvent("keydown", { key, bubbles: true, cancelable: true }));
    target.dispatchEvent(new KeyboardEvent("keyup", { key, bubbles: true, cancelable: true }));
  });
}

export interface MountedToolbar {
  bar: HTMLElement;
  /** The rendered document, needed to drive the caret. */
  page: HTMLElement;
  api: () => DocxViewApi;
  unmount: () => Promise<void>;
}

export async function mountToolbar(source: Uint8Array = FIXTURE): Promise<MountedToolbar> {
  const page = document.createElement("div");
  document.body.appendChild(page);
  const pageRoot: Root = createRoot(page);
  const bar = document.createElement("div");
  document.body.appendChild(bar);
  const barRoot: Root = createRoot(bar);
  const seen: { api: DocxViewApi | null } = { api: null };

  await act(async () => {
    pageRoot.render(
      createElement(DocxView, {
        source,
        editable: true,
        onReady: (api: DocxViewApi) => {
          seen.api = api;
        },
      }),
    );
  });
  for (let i = 0; i < 40 && !page.querySelector(".dxw-page"); i++) await tick();
  if (!seen.api) throw new Error("DocxView never reported an api");

  await act(async () => {
    barRoot.render(createElement(DocxToolbar, { api: seen.api }));
  });

  return {
    bar,
    page,
    api: () => seen.api!,
    unmount: async () => {
      await act(async () => {
        barRoot.unmount();
        pageRoot.unmount();
      });
      bar.remove();
      page.remove();
    },
  };
}

/**
 * Put the caret inside the first cell of `TABLE_FIXTURE`'s table, which is
 * what raises the Table Format tab.
 *
 * `find` leaves a selection rather than a caret, so this collapses it the way
 * a user would — the same route `table-properties-toolbar` proved.
 */
export async function caretIntoTable(t: MountedToolbar) {
  t.api().find("r0c0");
  await tick();
  await act(async () => {
    const target =
      (t.page.contains(document.activeElement) ? (document.activeElement as HTMLElement) : t.page.querySelector("textarea")) ?? t.page;
    target.dispatchEvent(new KeyboardEvent("keydown", { key: "ArrowRight", bubbles: true, cancelable: true }));
    await new Promise((resolve) => setTimeout(resolve, 2));
  });
  if (t.api().getTableCellFill() === undefined) throw new Error("caret never landed inside the table");
  // The bar recomputes its context from this event.
  await act(async () => {
    document.dispatchEvent(new Event("dxw-selection"));
  });
  await tick();
}

/**
 * Find a control on the bar by the tooltip text the user sees.
 *
 * `title` is matched, and `data-tip` with it: the bar moves `title` to
 * `data-tip` on first hover so the OS tooltip does not double up, and a test
 * that hovered before it clicked would otherwise stop finding the control.
 */
export function controlByTip(bar: HTMLElement, tip: string): HTMLElement {
  const found = [...bar.querySelectorAll<HTMLElement>("button, [role='button']")].find(
    (el) => (el.getAttribute("title") ?? el.getAttribute("data-tip")) === tip,
  );
  if (!found) {
    const available = [...bar.querySelectorAll<HTMLElement>("button")]
      .map((el) => el.getAttribute("title") ?? el.getAttribute("data-tip"))
      .filter(Boolean);
    throw new Error(`no control tipped "${tip}" on this tab. Present: ${JSON.stringify(available)}`);
  }
  return found;
}

export async function selectTab(bar: HTMLElement, tab: string) {
  const button = bar.querySelector<HTMLButtonElement>(`button[data-tab="${tab}"]`);
  if (!button) throw new Error(`no tab button for "${tab}"`);
  await click(button);
}

/** Every element currently in the document, including inside portals. */
function elementSnapshot(): Set<Element> {
  return new Set(document.querySelectorAll("*"));
}

/**
 * The subtree that appeared since `before`, which is the panel the click
 * opened. A "new root" is a new element whose parent was already there;
 * picking the one with the most descendants skips a re-rendered icon inside
 * the trigger and lands on the panel itself.
 */
function newSubtreeRoot(before: Set<Element>): HTMLElement | null {
  const roots: HTMLElement[] = [];
  for (const element of document.querySelectorAll<HTMLElement>("*")) {
    if (before.has(element)) continue;
    const parent = element.parentElement;
    if (parent && !before.has(parent)) continue;
    roots.push(element);
  }
  if (!roots.length) return null;
  return roots.reduce((best, candidate) =>
    candidate.querySelectorAll("*").length > best.querySelectorAll("*").length ? candidate : best,
  );
}

export interface OpenedPanel {
  /** The panel's root element. */
  panel: HTMLElement;
  /** The control that opened it. */
  trigger: HTMLElement;
}

/**
 * Click a control and return whatever it put on screen. Throws with the
 * control's name when nothing appeared, so "it did not open" reads as such in
 * the failure rather than as a null dereference three lines later.
 */
export async function openSurface(bar: HTMLElement, tip: string): Promise<OpenedPanel> {
  const trigger = controlByTip(bar, tip);
  const before = elementSnapshot();
  await click(trigger);
  await tick(0);
  const panel = newSubtreeRoot(before);
  if (!panel) throw new Error(`clicking "${tip}" put nothing on screen`);
  return { panel, trigger };
}

/** True while the panel is still in the document. */
export function isOpen(panel: HTMLElement): boolean {
  return document.contains(panel);
}

/**
 * Place the caret in the body text.
 *
 * A freshly mounted view has no caret, and the engine's insert calls all go
 * through `getCaretTarget()` — so without this they return early and do
 * nothing at all. `find` leaves a selection, which ArrowRight collapses.
 */
export async function caretIntoBody(t: MountedToolbar, needle: string) {
  t.api().find(needle);
  await tick();
  await act(async () => {
    const target =
      (t.page.contains(document.activeElement) ? (document.activeElement as HTMLElement) : t.page.querySelector("textarea")) ?? t.page;
    target.dispatchEvent(new KeyboardEvent("keydown", { key: "ArrowRight", bubbles: true, cancelable: true }));
    await new Promise((resolve) => setTimeout(resolve, 2));
  });
}

/**
 * Put a shape in the document, which raises the Format tab.
 *
 * Inserting selects the new drawing, so no separate click is needed — the
 * engine reports it as the selected object straight away.
 */
export async function selectAShape(t: MountedToolbar) {
  await caretIntoBody(t, "Body text");
  await act(async () => {
    if (!t.api().insertShape("rect", "Shape text")) throw new Error("insertShape refused");
  });
  await tick();
  if (!t.api().getSelectedObjectContext()) throw new Error("the new shape is not the selected object");
  // The bar recomputes its context from this event.
  await act(async () => {
    document.dispatchEvent(new Event("dxw-selection"));
  });
  await tick();
}

/**
 * Controls inside `panel` that a keyboard can reach.
 *
 * The `aria-hidden` native `<select>`/`<input>` that the styled controls keep
 * as an event bridge is excluded: it is deliberately out of the tab order, so
 * counting it would let a panel of unreachable `<div onClick>` swatches pass
 * as keyboard-operable.
 */
export function focusableControls(panel: HTMLElement): HTMLElement[] {
  return interactiveControls(panel).filter((element) => !element.hasAttribute("disabled"));
}

/**
 * The same controls, disabled ones included.
 *
 * `focusableControls` drops disabled controls, because a user cannot operate
 * them — right for "is this panel usable at all". Wrong for comparing two
 * opens: a control that comes back switched off would read as a control that
 * vanished, which points at the wrong defect.
 */
export function interactiveControls(panel: HTMLElement): HTMLElement[] {
  const selector = [
    "button",
    "input",
    "select",
    "textarea",
    "a[href]",
    "[tabindex]:not([tabindex='-1'])",
    "[contenteditable='true']",
  ].join(",");
  return [...panel.querySelectorAll<HTMLElement>(selector)].filter(
    (element) => !element.closest("[aria-hidden='true']"),
  );
}

/**
 * The name a screen reader would announce for `element`, or "" when it has
 * none.
 *
 * This is the accessible-name computation cut down to the sources this bar
 * actually uses, in the order the spec resolves them. `placeholder` is last
 * and is a genuine fallback for a text input, but only that — a button whose
 * only content is an icon resolves to "" here, which is exactly what a screen
 * reader announces for it.
 */
export function accessibleName(element: HTMLElement): string {
  const aria = element.getAttribute("aria-label");
  if (aria?.trim()) return aria.trim();

  const labelledBy = element.getAttribute("aria-labelledby");
  if (labelledBy) {
    const text = labelledBy
      .split(/\s+/)
      .map((id) => element.ownerDocument.getElementById(id)?.textContent ?? "")
      .join(" ")
      .trim();
    if (text) return text;
  }

  const id = element.getAttribute("id");
  // Walked rather than selected: an id may hold characters a CSS selector
  // would have to escape, and jsdom has no CSS.escape to do it with.
  const forLabel = id
    ? [...element.ownerDocument.querySelectorAll("label")].find((label) => label.htmlFor === id)
    : null;
  const wrapping = element.closest("label");
  const labelText = (forLabel?.textContent ?? wrapping?.textContent ?? "").trim();
  if (labelText) return labelText;

  // Content only names a control that can take its name from content —
  // a text input is named by its label, never by the text sitting next to it.
  const namedByContent = !(element instanceof HTMLInputElement)
    && !(element instanceof HTMLTextAreaElement)
    && !(element instanceof HTMLSelectElement);
  const content = (element.textContent ?? "").trim();
  if (namedByContent && content) return content;

  const title = element.getAttribute("title");
  if (title?.trim()) return title.trim();

  const alt = element.querySelector("img[alt]")?.getAttribute("alt");
  if (namedByContent && alt?.trim()) return alt.trim();

  const placeholder = element.getAttribute("placeholder");
  if (placeholder?.trim()) return placeholder.trim();

  return "";
}

/** Reachable controls inside `panel` that announce as nothing. */
export function unnamedControls(panel: HTMLElement): HTMLElement[] {
  return focusableControls(panel).filter((element) => !accessibleName(element));
}

/**
 * What a panel offers, in a form two opens can be compared by.
 *
 * Deliberately NOT the panel's HTML. Inline `left`/`top` differ between opens
 * for a panel that measures its anchor, and a diff of two 4KB style strings
 * says nothing a reader can act on. What matters for "works once, then not
 * again" is what the user can still do: which controls are there, what they
 * announce as, and whether they are switched off.
 */
export function panelSignature(panel: HTMLElement): string[] {
  return interactiveControls(panel).map((element) => {
    const tag = element.tagName.toLowerCase();
    const type = element.getAttribute("type") ?? "";
    const disabled = element.hasAttribute("disabled") ? " disabled" : "";
    return `${tag}${type ? `[${type}]` : ""} "${accessibleName(element)}"${disabled}`;
  });
}

/** An element outside every panel and outside the bar, to click on. */
export function outsideTarget(): HTMLElement {
  const existing = document.getElementById("dxw-smoke-outside");
  if (existing) return existing as HTMLElement;
  const element = document.createElement("div");
  element.id = "dxw-smoke-outside";
  document.body.appendChild(element);
  return element;
}

/**
 * A modal surface's outermost element is a backdrop covering the window, so
 * "outside the dialog" means the backdrop and nothing else — a click landing
 * anywhere else is impossible for a real user and would be a fake failure.
 */
function backdropOf(panel: HTMLElement): HTMLElement | null {
  return panel.querySelector("[aria-modal='true']") ? panel : null;
}

/** Click where a user would to dismiss `panel` without using it. */
export async function clickOutside(panel?: HTMLElement) {
  const target = (panel && backdropOf(panel)) ?? outsideTarget();
  await act(async () => {
    target.dispatchEvent(new MouseEvent("mousedown", { bubbles: true, cancelable: true }));
    target.dispatchEvent(new MouseEvent("mouseup", { bubbles: true, cancelable: true }));
    target.dispatchEvent(new MouseEvent("click", { bubbles: true, cancelable: true }));
  });
  await tick(0);
}

/**
 * Shut a panel by whatever means works, and report whether anything did.
 *
 * Used to get back to a known state between checks, so that a surface which
 * ignores Escape fails the Escape check ONLY — without this, every later
 * check on that surface fails too and one defect reads as four.
 */
export async function closeSurface({ panel, trigger }: OpenedPanel): Promise<boolean> {
  if (!isOpen(panel)) return true;
  await pressKey(panel, "Escape");
  if (!isOpen(panel)) return true;
  await clickOutside(panel);
  if (!isOpen(panel)) return true;
  // Last resort: the control that opened it. Every one of these toggles.
  await click(trigger);
  await tick(0);
  return !isOpen(panel);
}

export interface ConsoleWatch {
  /** Everything console.error and console.warn were given, joined per call. */
  messages: string[];
  /** Forget what has been recorded, so setup noise is not read as a defect. */
  reset: () => void;
  stop: () => string[];
}

/**
 * Record console.error/console.warn and unhandled rejections.
 *
 * React swallows a throw inside an event handler after logging it, so without
 * this a popover whose open handler explodes still renders a panel and passes
 * every other check in this file.
 */
export function watchConsole(): ConsoleWatch {
  const messages: string[] = [];
  const realError = console.error;
  const realWarn = console.warn;
  const record = (kind: string) => (...args: unknown[]) => {
    messages.push(`${kind}: ${args.map((a) => (a instanceof Error ? `${a.message}` : String(a))).join(" ")}`);
  };
  console.error = record("console.error");
  console.warn = record("console.warn");
  const onRejection = (event: PromiseRejectionEvent) => {
    messages.push(`unhandled rejection: ${String(event.reason)}`);
  };
  const onError = (event: ErrorEvent) => {
    messages.push(`window error: ${event.message}`);
  };
  window.addEventListener("unhandledrejection", onRejection as EventListener);
  window.addEventListener("error", onError as EventListener);
  return {
    messages,
    reset: () => {
      messages.length = 0;
    },
    stop: () => {
      console.error = realError;
      console.warn = realWarn;
      window.removeEventListener("unhandledrejection", onRejection as EventListener);
      window.removeEventListener("error", onError as EventListener);
      return messages;
    },
  };
}
