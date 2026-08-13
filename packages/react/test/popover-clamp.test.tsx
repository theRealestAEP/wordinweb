/**
 * #148: every toolbar popover, gallery and tooltip stays inside the window.
 *
 * The reported instance was the text-effects gallery: opened from a control
 * near the right edge of a narrow window it ran off the screen, the last
 * swatch cut in half and the rest unreachable. Its tooltip spilled with it.
 * The instance is not the bug — most panels in the bar were placed by plain
 * CSS (`position: absolute; top: 28; left: 0`) and every one of them did the
 * same thing.
 *
 * WHY THIS FILE CARRIES A LAYOUT MODEL. jsdom performs no layout: every
 * `getBoundingClientRect()` is zero, `offsetWidth` is zero, and a test that
 * asked "is the panel's right edge inside innerWidth?" would answer 0 <= 900
 * and pass on a panel that is plainly off screen in a browser. So the boxes
 * below are placed the way a browser places them: a `fixed` box lands at its
 * own `left`/`top`, an `absolute` box lands relative to its nearest positioned
 * ancestor, and an `overflow` box is no taller than its `max-height`. That is
 * the whole model, and it is enough to tell a clamped panel from an unclamped
 * one — the unclamped versions of these panels fail every assertion here.
 *
 * The rects are checked against a real browser in the parity demo; see the
 * report on #148. This suite is what keeps it fixed.
 */
import { act, createElement } from "react";
import { createRoot, type Root } from "react-dom/client";
import { zipSync, strToU8 } from "fflate";
import { afterEach, describe, expect, it } from "vitest";
import { DocxView, type DocxViewApi } from "../src/index.js";
import { DocxToolbar } from "../src/toolbar.js";

/** Narrow enough that a control on the right of the bar is near the edge. */
const VIEWPORT = { width: 900, height: 560 };
/** The bar, along the top, as tall as one ribbon line plus the tab strip. */
const BAR = { left: 0, top: 0, width: VIEWPORT.width, height: 64 };
/** Every panel is treated as this tall before its own max-height applies. */
const PANEL_HEIGHT = 260;
/** The margin the toolbar keeps between a panel and the window edge. */
const MARGIN = 8;

interface Box { left: number; top: number; width: number; height: number }

/** Rects the test has decided on: the bar, and the control being opened. */
const placed = new WeakMap<Element, Box>();

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
function cssPx(value: string): number | null {
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
function boxOf(el: HTMLElement): Box {
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

function installLayout(): () => void {
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

async function tick(ms = 5) {
  await act(async () => { await new Promise<void>((resolve) => setTimeout(resolve, ms)); });
}

interface Mounted {
  bar: HTMLElement;
  unmount: () => Promise<void>;
}

async function mount(): Promise<Mounted> {
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

async function selectTab(bar: HTMLElement, name: string) {
  const tab = [...bar.querySelectorAll<HTMLElement>("[data-dxw-toolbar-tabs] button")]
    .find((el) => (el.textContent ?? "").trim().toLowerCase() === name);
  if (!tab) throw new Error(`no "${name}" tab on the bar`);
  await act(async () => { tab.dispatchEvent(new MouseEvent("click", { bubbles: true, cancelable: true })); });
}

function controlByTip(bar: HTMLElement, tip: string): HTMLElement {
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
function placeControl(control: HTMLElement, box: Box) {
  placed.set(control, box);
  const wrapper = control.parentElement;
  if (wrapper) placed.set(wrapper, box);
}

const AT_RIGHT_EDGE: Box = { left: VIEWPORT.width - 36, top: 30, width: 32, height: 26 };
const AT_LEFT_EDGE: Box = { left: 4, top: 30, width: 32, height: 26 };

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

async function open(bar: HTMLElement, tip: string, where: Box = AT_RIGHT_EDGE): Promise<HTMLElement> {
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
function rectOf(panel: HTMLElement) {
  const box = boxOf(panel);
  return { ...box, right: box.left + box.width, bottom: box.top + box.height };
}

function expectInsideViewport(panel: HTMLElement, label: string) {
  const rect = rectOf(panel);
  expect(rect.width, `${label}: panel has no width, so this assertion proves nothing`).toBeGreaterThan(0);
  expect(rect.height, `${label}: panel has no height, so this assertion proves nothing`).toBeGreaterThan(0);
  expect(rect.right, `${label}: right edge ${rect.right} is past the window's ${VIEWPORT.width}`).toBeLessThanOrEqual(VIEWPORT.width - MARGIN);
  expect(rect.left, `${label}: left edge ${rect.left} is off the left of the window`).toBeGreaterThanOrEqual(0);
  expect(rect.bottom, `${label}: bottom edge ${rect.bottom} is past the window's ${VIEWPORT.height}`).toBeLessThanOrEqual(VIEWPORT.height);
  expect(rect.top, `${label}: top edge ${rect.top} is above the window`).toBeGreaterThanOrEqual(0);
}

let restoreLayout: (() => void) | null = null;
let mounted: Mounted | null = null;

afterEach(async () => {
  await mounted?.unmount();
  mounted = null;
  restoreLayout?.();
  restoreLayout = null;
});

/**
 * One entry per popover family in the bar. The point is the class: a fix that
 * only moved the text-effects gallery leaves the same defect in the other
 * thirty panels, which is exactly how #148 was reported twice.
 */
const SURFACES: { tab: string; tip: string; note: string }[] = [
  { tab: "home", tip: "Text Effects and Typography", note: "the gallery in the bug report" },
  { tab: "home", tip: "Highlight color", note: "a panel with no declared width" },
  { tab: "insert", tip: "Insert link", note: "left-aligned form" },
  { tab: "insert", tip: "Insert shape", note: "right-aligned gallery" },
  { tab: "insert", tip: "Insert text box", note: "right-aligned form" },
  { tab: "insert", tip: "Watermark", note: "right-aligned form" },
  { tab: "insert", tip: "Insert cover page", note: "a panel that hung from top: 30" },
  { tab: "insert", tip: "Insert advanced symbol", note: "was clamped across but not down" },
  { tab: "insert", tip: "Insert equation", note: "was clamped across but not down" },
  { tab: "insert", tip: "Insert bookmark", note: "left-aligned form" },
  { tab: "insert", tip: "Insert cross-reference", note: "left-aligned form" },
];

describe("#148 · toolbar panels stay inside the window", () => {
  for (const surface of SURFACES) {
    it(`${surface.tip} — opened from the right edge (${surface.note})`, async () => {
      restoreLayout = installLayout();
      mounted = await mount();
      await selectTab(mounted.bar, surface.tab);
      const panel = await open(mounted.bar, surface.tip);
      expectInsideViewport(panel, surface.tip);
    });
  }

  it("a panel opened from the LEFT edge does not hang off the left either", async () => {
    restoreLayout = installLayout();
    mounted = await mount();
    await selectTab(mounted.bar, "insert");
    // Right-aligned panels are the ones at risk here: their right edge is
    // pinned to a control that has almost no room to its left.
    const panel = await open(mounted.bar, "Insert shape", AT_LEFT_EDGE);
    expectInsideViewport(panel, "Insert shape at the left edge");
  });

  it("the text-effects gallery keeps every swatch reachable", async () => {
    restoreLayout = installLayout();
    mounted = await mount();
    await selectTab(mounted.bar, "home");
    const panel = await open(mounted.bar, "Text Effects and Typography");
    const rect = rectOf(panel);
    const swatches = panel.querySelectorAll("button");
    expect(swatches.length, "the gallery rendered no swatches").toBeGreaterThanOrEqual(6);
    // The grid fills the panel, so the last swatch is inside the window
    // exactly when the panel is. Half a swatch showing was the report.
    expect(rect.right).toBeLessThanOrEqual(VIEWPORT.width - MARGIN);
    expect(rect.left).toBeGreaterThanOrEqual(MARGIN);
  });

  it("a panel opened from a bar near the bottom flips above it", async () => {
    restoreLayout = installLayout();
    mounted = await mount();
    const toolbar = mounted.bar.querySelector<HTMLElement>("[data-dxw-toolbar-mode]")!;
    // A bar docked low: there is no room under it, and plenty over it.
    const low = { left: 0, top: VIEWPORT.height - 64, width: VIEWPORT.width, height: 64 };
    placed.set(toolbar, low);
    await selectTab(mounted.bar, "home");
    const panel = await open(mounted.bar, "Text Effects and Typography", {
      left: VIEWPORT.width - 36, top: low.top + 4, width: 32, height: 26,
    });
    const rect = rectOf(panel);
    expectInsideViewport(panel, "Text Effects opened from a low bar");
    expect(rect.bottom, "the panel should sit above the control, not under it").toBeLessThanOrEqual(low.top);
  });

  it("a tall panel in a short window is capped so it can scroll", async () => {
    restoreLayout = installLayout();
    mounted = await mount();
    (window as unknown as { innerHeight: number }).innerHeight = 200;
    await selectTab(mounted.bar, "insert");
    const panel = await open(mounted.bar, "Insert cross-reference");
    const cap = cssPx(panel.style.maxHeight);
    expect(cap, "no max-height, so a tall panel would run off the bottom").not.toBeNull();
    expect(boxOf(panel).top + Math.min(PANEL_HEIGHT, cap!)).toBeLessThanOrEqual(200);
    expect(["auto", "scroll"]).toContain(panel.style.overflowY || panel.style.overflow);
  });

  /**
   * The parser above is only useful if it returns the number a browser
   * returns. These two panels declare their widths as theme tokens, and a
   * parser that only read `Npx` would have silently called the colour menu
   * ~51px wide (its swatches have no text), placing its right edge 185px
   * inside the truth and passing a panel that was off screen.
   *
   * Chromium, demo at 900x600: the colour menu measures 236 and the layout
   * menu 304, from these same declarations. The model has to agree.
   */
  it("resolves themable widths to the same number a browser measures", async () => {
    restoreLayout = installLayout();
    mounted = await mount();
    await selectTab(mounted.bar, "home");
    const colorMenu = await open(mounted.bar, "Text color");
    expect(colorMenu.style.width).toContain("var(--dxw-color-menu-width");
    expect(boxOf(colorMenu).width, "the colour menu should model as 236, as Chromium measures it").toBe(236);
    expectInsideViewport(colorMenu, "Text color");

    await selectTab(mounted.bar, "layout");
    const layoutMenu = await open(mounted.bar, "Margins");
    expect(layoutMenu.style.width).toContain("var(--dxw-layout-menu-width");
    expect(boxOf(layoutMenu).width, "the layout menu should model as 304, as Chromium measures it").toBe(304);
    expectInsideViewport(layoutMenu, "Margins");
  });

  it("refuses to guess a width it cannot read", async () => {
    restoreLayout = installLayout();
    const box = document.createElement("div");
    box.style.position = "fixed";
    box.style.setProperty("width", "min(var(--x, 40px), 12em)");
    // Silence would mean the next unreadable declaration becomes an estimate
    // and the assertion built on it can no longer fail.
    expect(() => boxOf(box)).toThrow(/cannot read the CSS length/);
  });

  /**
   * #151. Escape did nothing in 24 of these panels, so a keyboard user who
   * opened one had no way back out. Closing one and leaving focus on the
   * body is the same bug in a smaller form, so the trigger has to get it.
   */
  for (const [tab, tip] of [["home", "Text Effects and Typography"], ["insert", "Insert link"]] as const) {
    it(`Escape closes ${tip} and hands focus back to its trigger`, async () => {
      restoreLayout = installLayout();
      mounted = await mount();
      await selectTab(mounted.bar, tab);
      const trigger = controlByTip(mounted.bar, tip);
      const panel = await open(mounted.bar, tip);
      // Restoring focus needs the panel's own ref, so this also catches a
      // panel that reads its placement but never attached one — five did,
      // which a browser found and a rect assertion never would.
      const inner = panel.querySelector<HTMLElement>("button, input, textarea")!;
      inner.focus();
      expect(document.activeElement, "the test never got focus into the panel").toBe(inner);

      await act(async () => {
        document.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape", bubbles: true }));
      });
      await tick(20);
      expect(panel.isConnected, "Escape left the panel open").toBe(false);
      expect(document.activeElement, "Escape closed the panel but stranded focus").toBe(trigger);
    });
  }

  /**
   * Popovers nest: the shape gallery holds a colour menu. Both listen on the
   * document, so without an order one Escape takes the gallery away as well
   * as the swatch grid the user meant to dismiss.
   */
  it("Escape closes only the innermost of two stacked panels", async () => {
    restoreLayout = installLayout();
    mounted = await mount();
    await selectTab(mounted.bar, "insert");
    const gallery = await open(mounted.bar, "Insert shape");
    const colorTrigger = gallery.querySelector<HTMLElement>("[data-dxw-color-trigger]");
    expect(colorTrigger, "the shape gallery no longer holds a colour menu").toBeTruthy();
    await act(async () => {
      colorTrigger!.dispatchEvent(new MouseEvent("mousedown", { bubbles: true, cancelable: true }));
      colorTrigger!.dispatchEvent(new MouseEvent("click", { bubbles: true, cancelable: true }));
    });
    await tick();
    const swatches = document.querySelector("[data-dxw-color-menu]");
    expect(swatches, "the colour menu never opened").toBeTruthy();

    await act(async () => {
      document.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape", bubbles: true }));
    });
    await tick(20);
    expect(document.querySelector("[data-dxw-color-menu]"), "Escape left the colour menu open").toBeNull();
    expect(gallery.isConnected, "Escape took the shape gallery away with the colour menu").toBe(true);

    // A second Escape then closes the gallery, which is the way out.
    await act(async () => {
      document.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape", bubbles: true }));
    });
    await tick(20);
    expect(gallery.isConnected, "a second Escape did not close the gallery").toBe(false);
  });

  it("the tooltip under a control at the right edge stays on screen", async () => {
    restoreLayout = installLayout();
    mounted = await mount();
    await selectTab(mounted.bar, "home");
    const control = controlByTip(mounted.bar, "Text Effects and Typography");
    placeControl(control, AT_RIGHT_EDGE);
    await act(async () => {
      control.dispatchEvent(new MouseEvent("mouseover", { bubbles: true, cancelable: true }));
    });
    await tick(700);
    // Found by the text the user reads, not by a marker: a marker only the
    // fixed version carries would turn a geometry test into a spelling test.
    const tooltip = [...document.querySelectorAll<HTMLElement>("div")]
      .find((el) => el.style.position === "fixed" && el.textContent === "Text Effects and Typography");
    expect(tooltip, "the tooltip never appeared").toBeTruthy();
    const rect = rectOf(tooltip!);
    expect(rect.width, "the tooltip has no width, so this assertion proves nothing").toBeGreaterThan(0);
    expect(rect.right, `the tooltip's right edge ${rect.right} is past the window`).toBeLessThanOrEqual(VIEWPORT.width);
    expect(rect.left).toBeGreaterThanOrEqual(0);
  });
});
