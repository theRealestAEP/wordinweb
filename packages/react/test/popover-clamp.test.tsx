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
 * The browser layout model this measures against now lives in
 * `popover-layout-model`, so that this suite and the 72-surface smoke suite
 * (#146) share ONE model instead of two that drift. Read that file for why a
 * model is needed at all — jsdom lays nothing out, so an unmodelled "is the
 * right edge inside innerWidth?" asks 0 <= 900 and passes on a panel that is
 * plainly off screen.
 *
 * What stays here is this ticket's own evidence: the surfaces that were
 * reported, the flip-above-a-low-bar case, the height cap, the themable
 * widths pinned to what Chromium measures, and #151's focus handover.
 */
import { act } from "react";
import { afterEach, describe, expect, it } from "vitest";
import {
  AT_LEFT_EDGE,
  AT_RIGHT_EDGE,
  boxOf,
  controlByTip,
  cssPx,
  expectInsideViewport,
  installLayout,
  MARGIN,
  mount,
  open,
  PANEL_HEIGHT,
  placeControl,
  placed,
  rectOf,
  selectTab,
  tick,
  VIEWPORT,
  type Mounted,
} from "./popover-layout-model.js";

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
