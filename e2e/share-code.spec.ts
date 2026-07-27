import { test, expect, type Page } from "@playwright/test";
import { LANDING, PAGE, waitHook } from "./_helpers";

/**
 * Share-code UX (doc 13 §7). The code is a second factor beyond the magic
 * link: it mixes into key derivation AND gates the hub, so a leaked link alone
 * can neither enter nor decrypt. Two UX rules this exercises:
 *
 *  1. The CREATOR, who just typed the code, must NOT be asked to re-enter it —
 *     their own code is remembered on their device and seeds the join.
 *  2. Anyone can always START A NEW document — a forgotten code locks you out
 *     of the OLD session, but never out of making a fresh one.
 */

const CODE = "246810";

/** Make the current blank local doc collaborative WITH a share code. */
async function goLiveWithCode(page: Page, code: string): Promise<string> {
  await page.goto(LANDING);
  await expect(page.getByTestId("local-editor")).toBeVisible();
  await page.getByTestId("share-code").fill(code);
  await page.getByTestId("make-collaborative").click();
  await expect(page).toHaveURL(/[?&]doc=/);
  await expect(page).toHaveURL(/#k=/);
  return page.url();
}

const codeInput = (page: Page) => page.getByPlaceholder("6-digit share code");

test.describe("zero-custody demo — share code UX", () => {
  test("the creator is NOT re-prompted for the code they just set", async ({ page }) => {
    await goLiveWithCode(page, CODE);
    // Lands straight in the live editor — no code prompt for the creator.
    await expect(page.getByTestId("download")).toBeVisible();
    await expect(page.locator(PAGE)).toBeVisible();
    await expect(codeInput(page), "creator must not see the share-code prompt").toHaveCount(0);
    await waitHook(page);
    const refused = await page.evaluate(
      () => (window as unknown as { __ww: { _session: { refused: string | null } } }).__ww._session.refused,
    );
    expect(refused, "creator's connection must not be code-refused").toBeNull();
  });

  test("a code-gated doc still refuses a joiner who lacks the code (the gate holds)", async ({ browser }) => {
    const owner = await browser.newContext();
    const joiner = await browser.newContext();
    try {
      const a = await owner.newPage();
      const url = await goLiveWithCode(a, CODE);
      await expect(a.getByTestId("download")).toBeVisible();

      // A different device/profile has no stored code → it is prompted.
      const b = await joiner.newPage();
      await b.goto(url);
      await expect(codeInput(b), "joiner without the code must be prompted").toBeVisible();
      // While refused, the editor isn't mounted (download etc. are App chrome
      // and stay visible — the meaningful signal is no rendered document).
      await expect(b.locator(PAGE)).toHaveCount(0);

      // The correct code lets the joiner in.
      await codeInput(b).fill(CODE);
      await b.getByRole("button", { name: "Join" }).click();
      await expect(b.locator(PAGE)).toBeVisible();
      await expect(codeInput(b), "prompt clears once joined").toHaveCount(0);
    } finally {
      await owner.close();
      await joiner.close();
    }
  });

  test("forgot the code → the refusal screen offers a 'new document' escape", async ({ browser }) => {
    const owner = await browser.newContext();
    const joiner = await browser.newContext();
    try {
      const a = await owner.newPage();
      const url = await goLiveWithCode(a, CODE);
      await expect(a.getByTestId("download")).toBeVisible();

      const b = await joiner.newPage();
      await b.goto(url);
      await expect(codeInput(b)).toBeVisible();
      // The escape hatch: leave the locked doc and start a fresh one.
      await b.getByTestId("refused-new-document").click();
      await expect(b.getByTestId("local-editor")).toBeVisible();
      await expect(b).not.toHaveURL(/[?&]doc=/);
    } finally {
      await owner.close();
      await joiner.close();
    }
  });

  test("the header 'New document' button returns to the local editor", async ({ page }) => {
    await goLiveWithCode(page, CODE);
    await expect(page.getByTestId("download")).toBeVisible();
    await page.getByTestId("new-document").click();
    await expect(page.getByTestId("local-editor")).toBeVisible();
    await expect(page).not.toHaveURL(/[?&]doc=/);
  });
});
