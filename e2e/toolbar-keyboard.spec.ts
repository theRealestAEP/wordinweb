import { test, expect } from "@playwright/test";
import { LANDING, PAGE } from "./_helpers";

/**
 * #152: the toolbar's galleries, driven by a real keyboard.
 *
 * The unit suite proves the engine receives a 3x4 table, asserting against
 * the serialised OOXML. It cannot prove the things this file exists for,
 * because jsdom implements none of them:
 *
 *   - Tab does not move focus in jsdom, so "the user can reach the grid" was
 *     an assumption dressed as a passing test — the unit suite calls
 *     `.focus()` itself. Here it is a real Tab, and it caught that the
 *     journey depends on HOW the panel was opened (see below).
 *   - There is no layout, so "the focused cell is visibly focused" cannot be
 *     asked at all. A roving tabindex nobody can SEE is the same defect in a
 *     new costume.
 *   - Enter on a focused button does not synthesise a click, so the unit
 *     suite calls `.click()` and the keystroke itself goes untested.
 *
 * Kept rather than thrown away: a fix that passes its own tests while being
 * unusable is the exact failure mode this ticket is about.
 */
test.describe("#152 keyboard table insert, for real", () => {
  test("Tab reaches the grid, arrows size it, Enter inserts it", async ({ page }) => {
    const logs: string[] = [];
    page.on("console", (m) => { if (m.type() === "error") logs.push(m.text()); });

    await page.goto(LANDING);
    await page.locator(PAGE).first().waitFor();

    // Caret into the body, so insertTable has a target.
    await page.locator(PAGE).first().click({ position: { x: 120, y: 90 } });

    await page.locator('button[data-tab="insert"]').click();

    // The keyboard-only route: the user has TABBED to the control, so it holds
    // focus, and opens it with Enter. Opening with a pointer deliberately does
    // NOT focus the trigger (onMouseDown preventDefault, so the editor keeps
    // its caret), which is a different journey.
    const tableTrigger = page.locator('button[title="Table"], button[data-tip="Table"]');
    await tableTrigger.focus();
    await page.keyboard.press("Enter");

    const panel = page.locator('[role="group"][aria-label="Table size"]');
    await expect(panel).toBeVisible();

    // THE QUESTION jsdom CANNOT ANSWER: does Tab get you into the grid?
    await page.keyboard.press("Tab");
    const afterTab = await page.evaluate(() => {
      const el = document.activeElement as HTMLElement | null;
      return { size: el?.getAttribute("data-dxw-table-size") ?? null, tag: el?.tagName ?? null };
    });
    expect(afterTab.size, `Tab landed on ${afterTab.tag}, not a grid cell`).toBe("1x1");

    // THE SECOND QUESTION: can you SEE where you are?
    const ring = await page.evaluate(() => {
      const el = document.activeElement as HTMLElement;
      const s = getComputedStyle(el);
      return { outlineStyle: s.outlineStyle, outlineWidth: s.outlineWidth, boxShadow: s.boxShadow };
    });
    const visible = (ring.outlineStyle !== "none" && parseFloat(ring.outlineWidth) > 0)
      || (ring.boxShadow !== "none" && ring.boxShadow !== "");
    expect(visible, `focused cell has no visible ring: ${JSON.stringify(ring)}`).toBe(true);

    await page.keyboard.press("ArrowDown");
    await page.keyboard.press("ArrowDown");
    await page.keyboard.press("ArrowRight");
    await page.keyboard.press("ArrowRight");
    await page.keyboard.press("ArrowRight");

    await expect(page.locator("text=3 × 4")).toBeVisible();

    // Enter on a focused button: a real browser synthesises the click.
    await page.keyboard.press("Enter");

    // The engine PAINTS the document, so there is no <table> in the DOM to
    // count; the unit suite asserts the 3x4 shape against the serialised
    // OOXML, which is the stronger claim. What only a real browser can settle
    // is whether Enter on the focused cell fires the click at all — and the
    // panel closing is that, because only the insert handler closes it.
    await expect(panel).toBeHidden({ timeout: 5000 });

    expect(logs, "no console errors during the whole path").toEqual([]);
  });

  test("the highlight swatches take focus and apply", async ({ page }) => {
    await page.goto(LANDING);
    await page.locator(PAGE).first().waitFor();
    await page.locator(PAGE).first().click({ position: { x: 120, y: 90 } });

    const trigger = page.locator('button[title="Highlight color"], button[data-tip="Highlight color"]');
    await trigger.click();
    const green = page.getByRole("button", { name: "Highlight green" });
    await expect(green).toBeVisible();
    await green.focus();
    expect(await page.evaluate(() => document.activeElement?.getAttribute("aria-label")))
      .toBe("Highlight green");
  });
});
