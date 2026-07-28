import { test, expect } from "@playwright/test";
import { BOARD_CODE } from "./_helpers";

/**
 * THE BARE-URL FLOW — the one the user actually types. Every other spec
 * passes `?server=`, which is why the same-origin regression (deploy arc)
 * shipped green: with no override the app derived its API from the page
 * origin, GET /blank hit Vite, and the local editor showed "Failed to
 * render document: invalid zip data". Found live by the user, not the
 * board. The dev stack now injects VITE_COLLAB_SERVER (scripts/dev.mjs)
 * so a bare URL resolves the API correctly in dev; production builds omit
 * the env and stay same-origin. This spec is the ONLY coverage of that
 * resolution chain's dev half — do not add `?server=` to it.
 */
test("bare URL (no ?server=) boots the local editor and can go live", async ({ page }) => {
  await page.goto("/");
  await expect(page.getByTestId("local-editor")).toBeVisible();
  // The blank template must have loaded — the regression's exact symptom
  // was this error text where the page should be.
  await expect(page.locator("text=/Failed to render/")).toHaveCount(0);
  await expect(page.locator(".dxw-page")).toBeVisible();
  // And the full go-live path resolves HTTP + WS without the override.
  await page.getByTestId("make-collaborative").click();
  await expect(page.getByTestId("collab-modal")).toBeVisible();
  await page.getByTestId("share-code").fill(BOARD_CODE); // a code is required to go live
  await page.getByTestId("start-collab").click();
  await expect(page).toHaveURL(/[?&]doc=/);
  await expect(page.getByTestId("download")).toBeVisible();
});
