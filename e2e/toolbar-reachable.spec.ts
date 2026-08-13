import { test, expect, type Page } from "@playwright/test";
import { LANDING, PAGE } from "./_helpers";

/**
 * #158: no window width may put a toolbar control out of reach.
 *
 * The defect was measured in a real browser and the unit suite models control
 * widths with a stub, so it is checked back in a real browser: jsdom's
 * measurements are plausible rather than true, and "which controls fit at
 * 900px" is exactly the question a plausible width can answer wrongly while
 * looking right. 900 and 700 are the widths the sweep reported — at 900,
 * Layout hid Hyphenation and Review hid Compare Documents, each with no
 * chevron and so no way back.
 */

const TABS = ["home", "insert", "draw", "layout", "review"] as const;
const WIDTHS = [1100, 900, 700] as const;

/** Every control on the bar right now, by the tooltip the user reads. */
async function shownControls(page: Page): Promise<string[]> {
  return page.evaluate(() => {
    const bar = document.querySelector("[data-dxw-toolbar-ribbon]");
    if (!bar) return [];
    return [...bar.querySelectorAll("button")]
      .filter((el) => (el as HTMLElement).dataset.dxwToolbarExpand === undefined)
      .filter((el) => {
        // A real browser answers this properly: a folded control has no box.
        const rect = el.getBoundingClientRect();
        return rect.width > 0 && rect.height > 0;
      })
      .map((el) => el.getAttribute("title") ?? el.getAttribute("data-tip") ?? el.getAttribute("aria-label"))
      .filter((name): name is string => !!name);
  });
}

/** What the user can get to: on the bar, or behind the chevron. */
async function reachable(page: Page, tab: string, width: number): Promise<string[]> {
  await page.setViewportSize({ width, height: 800 });
  await page.locator(`button[data-tab="${tab}"]`).click();
  await page.waitForTimeout(150);

  const collapsed = await shownControls(page);
  const chevron = page.locator("[data-dxw-toolbar-expand]");
  if (await chevron.count() === 0) return [...new Set(collapsed)].sort();

  await chevron.click();
  await page.waitForTimeout(150);
  const expanded = await shownControls(page);
  // Put it back, so the next measurement starts from the collapsed bar.
  if (await chevron.count() > 0) await chevron.click();
  await page.waitForTimeout(150);
  return [...new Set([...collapsed, ...expanded])].sort();
}

test.describe("#158 every toolbar control stays reachable at every width", () => {
  for (const tab of TABS) {
    test(`${tab} loses nothing between 1400px and 700px`, async ({ page }) => {
      await page.setViewportSize({ width: 1400, height: 900 });
      await page.goto(LANDING);
      await page.locator(PAGE).first().waitFor();

      const reference = await reachable(page, tab, 1400);
      expect(reference.length, `${tab} has controls to lose`).toBeGreaterThan(4);

      for (const width of WIDTHS) {
        const got = await reachable(page, tab, width);
        const missing = reference.filter((name) => !got.includes(name));
        expect(
          missing,
          `${tab} at ${width}px cannot reach ${missing.length} control(s)`,
        ).toEqual([]);
      }
    });
  }

  test("a bar that hides something always shows the chevron", async ({ page }) => {
    // The invariant directly, at the width the defect was reported at.
    await page.setViewportSize({ width: 1400, height: 900 });
    await page.goto(LANDING);
    await page.locator(PAGE).first().waitFor();

    for (const tab of TABS) {
      const full = await reachable(page, tab, 1400);
      for (const width of [900, 700]) {
        await page.setViewportSize({ width, height: 800 });
        await page.locator(`button[data-tab="${tab}"]`).click();
        await page.waitForTimeout(150);
        const onBar = await shownControls(page);
        const missing = full.filter((name) => !onBar.includes(name));
        if (missing.length === 0) continue;
        await expect(
          page.locator("[data-dxw-toolbar-expand]"),
          `${tab} at ${width}px hides ${missing.join(", ")} and offers no chevron`,
        ).toHaveCount(1);
      }
    }
  });
});
