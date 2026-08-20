import { test, expect, type Page } from "@playwright/test";
import { modelAndPictureDocx } from "./_docx-fixtures";
import { LANDING, PAGE, goLive } from "./_helpers";

/**
 * USER-REPORTED (#153): a 3D model's "Move" button drags it to the WRONG
 * place, then stops moving it at all. Measured before the fix, three +60/0
 * drags gave +138.8/+389.3, +138.8/+389.3, then 0/0 — the object never lands
 * where the pointer put it, and it walks vertically from a drag with no
 * vertical component.
 *
 * The claim under test is not "it moved". It is that the object's landing
 * spot equals the drag, three drags running. A floating PICTURE rides along
 * as the control: it takes the same move path and already lands exactly, so
 * a fix that quietly breaks it fails here rather than in a user's document.
 *
 * Real mouse input on purpose: dispatchEvent hits an element whether or not
 * the browser could ever route a click there, and the Move button sits on a
 * selection overlay stacked over a live <model-viewer>.
 */

/**
 * Open the fixture through the demo's File ▸ Open .docx input.
 *
 * `live` picks the ROUTE, and both have to be exercised. The demo's local
 * editor sets no `onIntent`, so it takes the local branch; the desktop app
 * pins its local editing onto the collab route (engine 9166144), so its
 * `onIntent` is defined even offline and every edit takes the intent branch.
 * The two write positions through different code, and a fix proved on one
 * says nothing about the other.
 */
async function openFixture(page: Page, live = true): Promise<void> {
  await page.goto(LANDING);
  await expect(page.getByTestId("local-editor")).toBeVisible();
  await page.locator('input[type="file"][accept*=".docx"]').setInputFiles({
    name: "model-and-picture.docx",
    mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    buffer: modelAndPictureDocx(),
  });
  await expect(page.locator(PAGE).first()).toBeVisible();
  await expect(page.locator("[data-dxw-model3d]")).toHaveCount(1);
  await expect(page.locator(`${PAGE} img`)).toHaveCount(1);
  if (!live) return;
  await goLive(page);
  await expect(page.locator("[data-dxw-model3d]")).toHaveCount(1);
}

/** Select an object, then put it at an exact page position through the wrap
 * bar's Position command — the same write a drag makes, with no gesture. */
async function positionAt(page: Page, selector: string, x: number, y: number): Promise<void> {
  await inView(page, selector);
  const box = (await page.locator(selector).boundingBox())!;
  await page.mouse.click(box.x + box.width / 2, box.y + 20);
  await expect(page.locator("[data-dxw-object-selection]")).toHaveCount(1);
  // The object's OWN wrap bar, not the ribbon's same-named button.
  await page.locator('button[title="Exact page position (px)"]').click();
  const dialog = page.locator("[data-dxw-number-pair-dialog]");
  await expect(dialog).toHaveCount(1);
  await dialog.getByLabel("X (pixels)").fill(String(x));
  await dialog.getByLabel("Y (pixels)").fill(String(y));
  await dialog.getByRole("button", { name: "Apply" }).click();
  await expect(dialog).toHaveCount(0);
  await page.waitForTimeout(350);
}

/** Press-move-release with a real mouse, exactly `dx`/`dy` client px. */
async function dragBy(page: Page, from: { x: number; y: number }, dx: number, dy: number): Promise<void> {
  await page.mouse.move(from.x, from.y);
  await page.mouse.down();
  await page.mouse.move(from.x + dx / 2, from.y + dy / 2);
  await page.mouse.move(from.x + dx, from.y + dy);
  await page.mouse.up();
  await page.waitForTimeout(350);
}

const centre = (b: { x: number; y: number; width: number; height: number }) => ({
  x: b.x + b.width / 2,
  y: b.y + b.height / 2,
});

/**
 * The object's position in DOCUMENT px (its inline left/top on the page
 * surface). Client rects would answer a different question here: a drag can
 * carry the object below the fold, and every later measurement would then be
 * reporting the scroll as movement.
 */
async function at(page: Page, selector: string): Promise<{ left: number; top: number }> {
  return page.evaluate((sel) => {
    const el = document.querySelector<HTMLElement>(sel)!;
    return { left: parseFloat(el.style.left) || 0, top: parseFloat(el.style.top) || 0 };
  }, selector);
}

/** Scroll the object into view and hand back a grab point that is really on
 * screen — a press outside the viewport hits nothing, which reads exactly
 * like "the object stopped responding". */
async function inView(page: Page, selector: string): Promise<void> {
  await page.evaluate((sel) => document.querySelector<HTMLElement>(sel)!.scrollIntoView({ block: "center" }), selector);
  await page.waitForTimeout(250);
}

async function onScreen(page: Page, box: { x: number; y: number }): Promise<boolean> {
  const vp = page.viewportSize()!;
  return box.x > 0 && box.y > 0 && box.x < vp.width && box.y < vp.height;
}

/**
 * The same defect with no gesture at all, and with an exact expected answer
 * rather than a delta. "Exact page position" reads the object's own CSS
 * position — page-surface px — and writes the number straight back through
 * setFloatingPagePosition, which resolves it against the ANCHOR's origin
 * instead. Ask for (0,0) and the object lands ON its anchor origin, which
 * names the defect: landed = requested + anchorOrigin.
 *
 * Both routes, because they are different code. Both were wrong.
 */
for (const live of [false, true]) {
  const route = live ? "collab route (what the desktop app runs)" : "local route";
  test.describe(`#153 exact page position, ${route}`, () => {
    test("the model lands on the page position it was given", async ({ page }) => {
      await openFixture(page, live);
      await positionAt(page, "[data-dxw-model3d]", 0, 0);
      const got = await at(page, "[data-dxw-model3d]");
      expect(
        Math.abs(got.left) < 1 && Math.abs(got.top) < 1,
        `Position (0,0) must land the model at (0,0), got (${got.left.toFixed(3)}, ${got.top.toFixed(3)})`,
      ).toBe(true);
    });

    test("and on a second, non-zero position, without compounding", async ({ page }) => {
      await openFixture(page, live);
      await positionAt(page, "[data-dxw-model3d]", 100, 200);
      const first = await at(page, "[data-dxw-model3d]");
      await positionAt(page, "[data-dxw-model3d]", 140, 240);
      const second = await at(page, "[data-dxw-model3d]");
      expect(
        Math.abs(first.left - 100) < 1 && Math.abs(first.top - 200) < 1 &&
        Math.abs(second.left - 140) < 1 && Math.abs(second.top - 240) < 1,
        `Position must be absolute: (100,200) gave (${first.left.toFixed(3)}, ${first.top.toFixed(3)}), ` +
        `(140,240) gave (${second.left.toFixed(3)}, ${second.top.toFixed(3)})`,
      ).toBe(true);
    });

    test("the floating picture control keeps landing where it is put", async ({ page }) => {
      await openFixture(page, live);
      await positionAt(page, `${PAGE} img`, 120, 60);
      const got = await at(page, `${PAGE} img`);
      expect(
        Math.abs(got.left - 120) < 1 && Math.abs(got.top - 60) < 1,
        `Position (120,60) must land the picture at (120,60), got (${got.left.toFixed(3)}, ${got.top.toFixed(3)})`,
      ).toBe(true);
    });
  });
}

test.describe("#153 a 3D model's Move button", () => {
  test("three +60px drags leave the model exactly +180px across and 0 down", async ({ page }) => {
    await openFixture(page);
    const start = await at(page, "[data-dxw-model3d]");

    for (let i = 1; i <= 3; i++) {
      await inView(page, "[data-dxw-model3d]");
      // Select the object: its Move button only exists on the selection
      // overlay. Press above the centre rotate puck, which owns the middle.
      const box = (await page.locator("[data-dxw-model3d]").boundingBox())!;
      await page.mouse.click(box.x + box.width / 2, box.y + 20);
      const move = page.locator("[data-dxw-object-move]");
      await expect(move, `the Move button should be on the overlay before drag ${i}`).toHaveCount(1);
      const grab = centre((await move.boundingBox())!);
      expect(await onScreen(page, grab), `the Move button must be on screen for drag ${i}`).toBe(true);
      await dragBy(page, grab, 60, 0);
    }

    const end = await at(page, "[data-dxw-model3d]");
    const dx = end.left - start.left;
    const dy = end.top - start.top;
    expect(
      Math.abs(dx - 180) < 1 && Math.abs(dy) < 1,
      `three +60/0 drags must land the model at +180/0, got +${dx.toFixed(1)}/+${dy.toFixed(1)}`,
    ).toBe(true);
  });

  test("the floating picture control still lands exactly where it is dragged", async ({ page }) => {
    await openFixture(page);
    const start = await at(page, `${PAGE} img`);

    for (let i = 1; i <= 3; i++) {
      await inView(page, `${PAGE} img`);
      const box = (await page.locator(`${PAGE} img`).first().boundingBox())!;
      await page.mouse.click(centre(box).x, centre(box).y);
      await expect(page.locator("[data-dxw-object-selection]")).toHaveCount(1);
      const grab = centre((await page.locator(`${PAGE} img`).first().boundingBox())!);
      expect(await onScreen(page, grab), `the picture must be on screen for drag ${i}`).toBe(true);
      await dragBy(page, grab, 60, 0);
    }

    const end = await at(page, `${PAGE} img`);
    const dx = end.left - start.left;
    const dy = end.top - start.top;
    expect(
      Math.abs(dx - 180) < 1 && Math.abs(dy) < 1,
      `three +60/0 drags must land the picture at +180/0, got +${dx.toFixed(1)}/+${dy.toFixed(1)}`,
    ).toBe(true);
  });
});
