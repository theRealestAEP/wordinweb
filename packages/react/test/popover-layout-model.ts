/**
 * The browser layout model the toolbar's placement tests measure against.
 *
 * Extracted from `popover-clamp` (#148) so that the clamp suite and the
 * 72-surface smoke suite (#146) share ONE model rather than two that drift.
 * The code is that suite's, unchanged; only its home moved.
 *
 * WHY A MODEL AT ALL. jsdom performs no layout: every `getBoundingClientRect()`
 * is zero and `offsetWidth` is zero, so a test that asked "is the panel's right
 * edge inside innerWidth?" would answer 0 <= 900 and pass on a panel that is
 * plainly off screen in a browser. The boxes below are placed the way a browser
 * places them: a `fixed` box lands at its own `left`/`top`, an `absolute` box
 * lands relative to its nearest positioned ancestor, and an `overflow` box is
 * no taller than its `max-height`. That is the whole model, and it is enough to
 * tell a clamped panel from an unclamped one.
 *
 * The rects are checked against a real browser in the parity demo; see the
 * report on #148.
 */
import { act, createElement } from "react";
import { createRoot, type Root } from "react-dom/client";
import { zipSync, strToU8 } from "fflate";
import { expect } from "vitest";
import { DocxView, type DocxViewApi } from "../src/index.js";
import { DocxToolbar } from "../src/toolbar.js";

/** Narrow enough that a control on the right of the bar is near the edge. */
export const VIEWPORT = { width: 900, height: 560 };
/** The bar, along the top, as tall as one ribbon line plus the tab strip. */
export const BAR = { left: 0, top: 0, width: VIEWPORT.width, height: 64 };
/** Every panel is treated as this tall before its own max-height applies. */
export const PANEL_HEIGHT = 260;
/** The margin the toolbar keeps between a panel and the window edge. */
export const MARGIN = 8;

export interface Box { left: number; top: number; width: number; height: number }

/** Rects the test has decided on: the bar, and the control being opened. */
export const placed = new WeakMap<Element, Box>();

/**
 * A CSS length in pixels, or null when there is no declaration at all.
 *
 * It resolves the three shapes toolbar.tsx writes: a plain `Npx`, a
 * `var(--name, Npx)` reduced to its fallback, and `min(a, calc(100vw - Npx))`
 * reduced against this viewport. The colour menu and the layout menu declare
 * their widths as themable tokens, so a parser that only read `Npx` would
 * treat a 236px panel as having none.
 *
 * It THROWS on a declaration it cannot read rather than returning null. A
 * parser with a `?? guess` behind it is how a measurement quietly becomes an
 * estimate and an assertion quietly stops being able to fail.
 */
export function cssPx(value: string): number | null {
  const text = (value ?? "").trim();
  if (text === "") return null;
  const plain = /^(-?\d+(?:\.\d+)?)px$/.exec(text);
  if (plain) return Number(plain[1]);
  // var(--name, 236px) → 236
  const token = /^var\(\s*--[\w-]+\s*,\s*(-?\d+(?:\.\d+)?)px\s*\)$/.exec(text);
  if (token) return Number(token[1]);
  // min(<a>, calc(100vw - Npx)) → min(a, viewport - N), either order
  const clamp = /^min\((.+)\)$/.exec(text);
  if (clamp) {
    const parts = splitTopLevel(clamp[1]).map((part) => {
      const vw = /^calc\(\s*100vw\s*-\s*(\d+(?:\.\d+)?)px\s*\)$/.exec(part.trim());
      return vw ? VIEWPORT.width - Number(vw[1]) : cssPx(part.trim());
    });
    if (parts.every((n): n is number => n !== null)) return Math.min(...parts);
  }
  throw new Error(`the layout model cannot read the CSS length ${JSON.stringify(text)}`);
}

/** Split "a, calc(b - c)" on its top-level commas only. */
function splitTopLevel(text: string): string[] {
  const parts: string[] = [];
  let depth = 0;
  let start = 0;
  for (let i = 0; i < text.length; i++) {
    if (text[i] === "(") depth++;
    else if (text[i] === ")") depth--;
    else if (text[i] === "," && depth === 0) {
      parts.push(text.slice(start, i));
      start = i + 1;
    }
  }
  parts.push(text.slice(start));
  return parts;
}

function positionedAncestor(el: HTMLElement): HTMLElement | null {
  for (let node = el.parentElement; node; node = node.parentElement) {
    if (placed.has(node)) return node;
    if (node.style.position === "relative" || node.style.position === "absolute" || node.style.position === "fixed") return node;
  }
  return null;
}

/**
 * A box with no width of its own is as wide as its text. jsdom measures
 * nothing, so this stands in for text measurement the same way the canvas
 * stub in setup.ts does — the tooltip is the box that needs it.
 *
 * ONE LIMIT WORTH KNOWING. jsdom keeps a CSS value it cannot validate only
 * when it contains `var()`; `min(560px, calc(100vw - 32px))` with no token is
 * dropped outright, so `style.width` reads "" and this fallback fires with no
 * way to tell it apart from a box that never declared a width. The SmartArt
 * modal is the one panel in the file written that way, and it is centred by
 * its overlay rather than anchored to a control, so a clamp assertion on it
 * would be meaningless anyway — do not add one and assume it holds.
 */
function intrinsicWidth(el: HTMLElement): number {
  return Math.min(VIEWPORT.width - 2 * MARGIN, (el.textContent ?? "").length * 7 + 16);
}

/** Where this box lands on the screen. */
export function boxOf(el: HTMLElement): Box {
  const decided = placed.get(el);
  if (decided) return decided;
  const style = el.style;
  const positioned = style.position === "fixed" || style.position === "absolute";
  if (!positioned) {
    const host = positionedAncestor(el);
    return host ? { ...boxOf(host), width: 0, height: 0 } : { left: 0, top: 0, width: 0, height: 0 };
  }
  const width = cssPx(style.width) ?? intrinsicWidth(el);
  const height = Math.min(PANEL_HEIGHT, cssPx(style.maxHeight) ?? Infinity);
  // The one transform this file uses: a box centred on its anchor.
  const shift = style.transform === "translateX(-50%)" ? -width / 2 : 0;
  if (style.position === "fixed") {
    return { left: (cssPx(style.left) ?? 0) + shift, top: cssPx(style.top) ?? 0, width, height };
  }
  const host = positionedAncestor(el);
  const base = host ? boxOf(host) : { left: 0, top: 0, width: VIEWPORT.width, height: VIEWPORT.height };
  const left = style.left !== ""
    ? base.left + (cssPx(style.left) ?? 0)
    : style.right !== ""
      ? base.left + base.width - (cssPx(style.right) ?? 0) - width
      : base.left;
  return { left, top: base.top + (cssPx(style.top) ?? 0), width, height };
}

/**
 * Patch geometry onto `HTMLElement.prototype` and return the undo.
 *
 * It is global while installed, so anything else the test drives is measured
 * by this model too. Install it around the OPEN rather than around a setup
 * step that runs the real engine — caret placement and object selection read
 * element geometry, and they should get jsdom's zeros, not this.
 */
export function installLayout(): () => void {
  const win = window as unknown as { innerWidth: number; innerHeight: number };
  const priorWidth = win.innerWidth;
  const priorHeight = win.innerHeight;
  win.innerWidth = VIEWPORT.width;
  win.innerHeight = VIEWPORT.height;
  const priorRect = HTMLElement.prototype.getBoundingClientRect;
  const own = (name: "offsetWidth" | "offsetHeight") =>
    Object.getOwnPropertyDescriptor(HTMLElement.prototype, name);
  const priorOffsets = { offsetWidth: own("offsetWidth"), offsetHeight: own("offsetHeight") };
  HTMLElement.prototype.getBoundingClientRect = function () {
    const box = boxOf(this as HTMLElement);
    return {
      x: box.left, y: box.top, left: box.left, top: box.top,
      width: box.width, height: box.height,
      right: box.left + box.width, bottom: box.top + box.height,
      toJSON() { return this; },
    } as DOMRect;
  };
  Object.defineProperty(HTMLElement.prototype, "offsetWidth", {
    configurable: true, get(this: HTMLElement) { return boxOf(this).width; },
  });
  Object.defineProperty(HTMLElement.prototype, "offsetHeight", {
    configurable: true, get(this: HTMLElement) { return boxOf(this).height; },
  });
  return () => {
    HTMLElement.prototype.getBoundingClientRect = priorRect;
    for (const [name, descriptor] of Object.entries(priorOffsets)) {
      if (descriptor) Object.defineProperty(HTMLElement.prototype, name, descriptor);
      else delete (HTMLElement.prototype as unknown as Record<string, unknown>)[name];
    }
    win.innerWidth = priorWidth;
    win.innerHeight = priorHeight;
  };
}

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
        `<w:body><w:p><w:r><w:t xml:space="preserve">Body</w:t></w:r></w:p></w:body></w:document>`,
    ),
  });
}

export async function tick(ms = 5) {
  await act(async () => { await new Promise<void>((resolve) => setTimeout(resolve, ms)); });
}

export interface Mounted {
  bar: HTMLElement;
  unmount: () => Promise<void>;
}

/** A view and a bar, with the toolbar's own rect seeded into the model. */
export async function mount(): Promise<Mounted> {
  const page = document.createElement("div");
  document.body.appendChild(page);
  const pageRoot: Root = createRoot(page);
  const bar = document.createElement("div");
  document.body.appendChild(bar);
  const barRoot: Root = createRoot(bar);
  const seen: { api: DocxViewApi | null } = { api: null };
  await act(async () => {
    pageRoot.render(createElement(DocxView, {
      source: docx(""), editable: true, onReady: (api: DocxViewApi) => { seen.api = api; },
    }));
  });
  for (let i = 0; i < 40 && !page.querySelector(".dxw-page"); i++) await tick();
  if (!seen.api) throw new Error("DocxView never reported an api");
  await act(async () => { barRoot.render(createElement(DocxToolbar, { api: seen.api })); });
  const toolbar = bar.querySelector<HTMLElement>("[data-dxw-toolbar-mode]");
  if (!toolbar) throw new Error("the toolbar never rendered");
  placed.set(toolbar, BAR);
  return {
    bar,
    unmount: async () => {
      await act(async () => { barRoot.unmount(); pageRoot.unmount(); });
      bar.remove();
      page.remove();
    },
  };
}

export async function selectTab(bar: HTMLElement, name: string) {
  const tab = [...bar.querySelectorAll<HTMLElement>("[data-dxw-toolbar-tabs] button")]
    .find((el) => (el.textContent ?? "").trim().toLowerCase() === name);
  if (!tab) throw new Error(`no "${name}" tab on the bar`);
  await act(async () => { tab.dispatchEvent(new MouseEvent("click", { bubbles: true, cancelable: true })); });
}

export function controlByTip(bar: HTMLElement, tip: string): HTMLElement {
  const found = [...bar.querySelectorAll<HTMLElement>("button")]
    .find((el) => (el.getAttribute("title") ?? el.getAttribute("data-tip")) === tip);
  if (!found) throw new Error(`no control tipped "${tip}" on this tab`);
  return found;
}

/**
 * Put a control where the defect shows: hard against the right edge of the
 * bar. Its wrapping span is the containing block an unclamped `absolute`
 * panel is placed against, so both get the rect.
 */
export function placeControl(control: HTMLElement, box: Box) {
  placed.set(control, box);
  const wrapper = control.parentElement;
  if (wrapper) placed.set(wrapper, box);
}

export const AT_RIGHT_EDGE: Box = { left: VIEWPORT.width - 36, top: 30, width: 32, height: 26 };
export const AT_LEFT_EDGE: Box = { left: 4, top: 30, width: 32, height: 26 };

/** The panel a click opened: the subtree that was not there before. */
function newPanel(before: Set<Element>, bar: HTMLElement): HTMLElement {
  const candidates = [...document.querySelectorAll<HTMLElement>("div, span")]
    .filter((el) => !before.has(el))
    .filter((el) => el.style.position === "fixed" || el.style.position === "absolute")
    .filter((el) => el.style.zIndex !== "" && bar.contains(el) === bar.contains(el));
  if (candidates.length === 0) throw new Error("clicking the control opened no panel");
  // The outermost one: a panel can hold positioned children of its own.
  return candidates.find((el) => !candidates.some((other) => other !== el && other.contains(el)))!;
}

export async function open(bar: HTMLElement, tip: string, where: Box = AT_RIGHT_EDGE): Promise<HTMLElement> {
  const control = controlByTip(bar, tip);
  placeControl(control, where);
  const before = new Set(document.querySelectorAll("*"));
  await act(async () => {
    control.dispatchEvent(new MouseEvent("mousedown", { bubbles: true, cancelable: true }));
    control.dispatchEvent(new MouseEvent("click", { bubbles: true, cancelable: true }));
  });
  await tick();
  return newPanel(before, bar);
}

/** What the user sees: the box the panel actually occupies. */
export function rectOf(panel: HTMLElement) {
  const box = boxOf(panel);
  return { ...box, right: box.left + box.width, bottom: box.top + box.height };
}

export function expectInsideViewport(panel: HTMLElement, label: string) {
  const rect = rectOf(panel);
  expect(rect.width, `${label}: panel has no width, so this assertion proves nothing`).toBeGreaterThan(0);
  expect(rect.height, `${label}: panel has no height, so this assertion proves nothing`).toBeGreaterThan(0);
  expect(rect.right, `${label}: right edge ${rect.right} is past the window's ${VIEWPORT.width}`).toBeLessThanOrEqual(VIEWPORT.width - MARGIN);
  expect(rect.left, `${label}: left edge ${rect.left} is off the left of the window`).toBeGreaterThanOrEqual(0);
  expect(rect.bottom, `${label}: bottom edge ${rect.bottom} is past the window's ${VIEWPORT.height}`).toBeLessThanOrEqual(VIEWPORT.height);
  expect(rect.top, `${label}: top edge ${rect.top} is above the window`).toBeGreaterThanOrEqual(0);
}
