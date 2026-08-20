import { test, expect, type Page } from "@playwright/test";
import { modelAndPictureDocx } from "./_docx-fixtures";
import { LANDING, PAGE, goLive, saveB64 } from "./_helpers";

/**
 * USER-REPORTED (#161): clicking a floating picture to select it did nothing.
 *
 * The visible half was that no selection appeared. The half nobody had seen
 * is that the same click EDITED THE DOCUMENT: a picture is not a glyph, so
 * the press found no caret, Word's click-and-type read it as a click in empty
 * space and split a paragraph — and in a room that rides the wire, so every
 * peer got the inserted paragraph too. The re-render it caused is what
 * replaced the picture's node mid-gesture and left the selection overlay
 * attached to a dead subtree.
 *
 * WHAT MAKES IT REPRODUCE is the picture sitting BELOW the last line of body
 * text on its page — click-and-type only acts there. The sidebar table is
 * just what pushes it past that line; the ordering is not itself the cause.
 *
 * Driven in COLLAB mode with real mouse input. The desktop app pins its local
 * editing onto the collab route, and click-and-type takes a different branch
 * on each, so a local-only test would be testing code the app never runs.
 * Real input because the overlay this asserts on is only ever built from a
 * genuine hit test.
 */

async function openLive(page: Page): Promise<void> {
  await page.goto(LANDING);
  await expect(page.getByTestId("local-editor")).toBeVisible();
  await page.locator('input[type="file"][accept*=".docx"]').setInputFiles({
    name: "object-select.docx",
    mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    buffer: modelAndPictureDocx({ pictureAfterTable: true }),
  });
  await expect(page.locator(PAGE).first()).toBeVisible();
  await expect(page.locator(`${PAGE} img`)).toHaveCount(1);
  await goLive(page);
}

/** Scroll the picture into view and click its centre with a real mouse. */
async function clickPicture(page: Page): Promise<void> {
  await page.evaluate(() => document.querySelector<HTMLElement>(".dxw-page img")!.scrollIntoView({ block: "center" }));
  await page.waitForTimeout(250);
  const box = (await page.locator(`${PAGE} img`).first().boundingBox())!;
  await page.mouse.click(box.x + box.width / 2, box.y + box.height / 2);
  await page.waitForTimeout(400);
}

test.describe("#161 selecting a floating object below the last line of text", () => {
  test("the selection overlay is really in the document, not on a detached node", async ({ page }) => {
    await openLive(page);
    await clickPicture(page);

    // `document.contains`, not a count. The overlay was BUILT every time —
    // selectImage ran and appended it — but to the parent of a node that a
    // re-render had already replaced, so it never entered the page. A test
    // asserting the call happened, or even that an element exists, passes
    // against the bug.
    const overlay = await page.evaluate(() => {
      const el = document.querySelector("[data-dxw-object-selection]");
      return { exists: !!el, inDocument: !!el && document.contains(el) };
    });
    expect(overlay, "clicking a floating picture must put its selection overlay in the page").toEqual({
      exists: true,
      inDocument: true,
    });
  });

  test("and selecting it does not edit the document", async ({ page }) => {
    await openLive(page);
    const before = await saveB64(page);
    await clickPicture(page);
    const after = await saveB64(page);

    expect(before, "the fixture must have saved before the click").not.toBeNull();
    expect(after, "selecting an object must not change a single byte of the document").toBe(before);
  });

  test("the selection chrome it builds is hit-testable", async ({ page }) => {
    await openLive(page);
    await clickPicture(page);

    // The resize handles and the wrap bar are children of that same overlay,
    // so they inherit whichever element it was built from. Hit-testing them
    // is what proves the whole chrome is live rather than merely present:
    // an overlay in a detached subtree still answers querySelector.
    const chrome = await page.evaluate(() => {
      const handle = document.querySelector<HTMLElement>("[data-dxw-img-handle]");
      if (!handle) return { handles: 0, handleHit: false };
      const r = handle.getBoundingClientRect();
      const hit = document.elementFromPoint(r.left + r.width / 2, r.top + r.height / 2);
      return {
        handles: document.querySelectorAll("[data-dxw-img-handle]").length,
        handleHit: !!hit && (hit === handle || handle.contains(hit)),
      };
    });
    expect(chrome.handles, "a selected picture offers its resize handles").toBeGreaterThan(0);
    expect(chrome.handleHit, "and the browser can actually hit one").toBe(true);
  });

  test("clicking away still deselects", async ({ page }) => {
    await openLive(page);
    await clickPicture(page);
    await expect(page.locator("[data-dxw-object-selection]")).toHaveCount(1);

    // Suppressing the caret handler for the object press must not leak into
    // the NEXT press, or the object could never be deselected.
    const away = await page.evaluate(() => {
      for (const el of Array.from(document.querySelectorAll<HTMLElement>(".dxw-page *"))) {
        if (el.children.length || !(el.textContent ?? "").trim()) continue;
        if (el.closest("[data-dxw-model3d]") || el.closest("[data-dxw-image-format]")) continue;
        const r = el.getBoundingClientRect();
        if (r.width < 5 || r.top < 10 || r.bottom > window.innerHeight - 10) continue;
        return { x: r.left + 4, y: r.top + r.height / 2 };
      }
      return null;
    });
    expect(away, "the fixture must offer a clickable line of body text").not.toBeNull();
    await page.mouse.click(away!.x, away!.y);
    await page.waitForTimeout(400);
    await expect(page.locator("[data-dxw-object-selection]")).toHaveCount(0);
  });
});
