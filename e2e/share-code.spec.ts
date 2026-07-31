import { test, expect, type Page } from "@playwright/test";
import { LANDING, PAGE, waitHook } from "./_helpers";

/**
 * Share-code UX (doc 13 §7). The code is a second factor beyond the magic
 * link: it mixes into key derivation AND gates the hub, so a leaked link alone
 * can neither enter nor decrypt. It is now MANDATORY — a session cannot be
 * started without one — so what this spec covers is the three UX rules that
 * makes load-bearing:
 *
 *  0. There is no way to go live without a code (the first test): the enforcement
 *     is a disabled button, so it is worth pinning that the button is really
 *     disabled and that whitespace does not satisfy it.
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
  // The code field moved INTO the collab dialog, where the explanation of
  // what a code does sits next to it.
  await page.getByTestId("make-collaborative").click();
  await expect(page.getByTestId("collab-modal")).toBeVisible();
  await page.getByTestId("share-code").fill(code);
  await page.getByTestId("start-collab").click();
  await expect(page).toHaveURL(/[?&]doc=/);
  await expect(page).toHaveURL(/#k=/);
  return page.url();
}

// The JOIN prompt's field (the refusal screen), not the modal's — they are
// separate controls with separate testids. Selected by testid rather than by
// placeholder text: the placeholder is user-facing copy, and the last time it
// was the selector, correcting the copy would have silently broken the specs
// that depend on this prompt appearing.
const codeInput = (page: Page) => page.getByTestId("join-share-code");

test.describe("zero-custody demo — share code UX", () => {
  test("a code is REQUIRED — Start Collab stays disabled until one is typed", async ({ page }) => {
    await page.goto(LANDING);
    await expect(page.getByTestId("local-editor")).toBeVisible();
    await page.getByTestId("make-collaborative").click();
    await expect(page.getByTestId("collab-modal")).toBeVisible();
    // The disabled button IS the enforcement — there is no validation behind
    // it to fall back on, so if it ever enables while empty, a session can be
    // created that a leaked link opens on its own.
    await expect(page.getByTestId("start-collab")).toBeDisabled();
    // Whitespace is not a code (the field is trimmed before it is used, so a
    // spaces-only value would otherwise go live with an EMPTY one).
    await page.getByTestId("share-code").fill("   ");
    await expect(page.getByTestId("start-collab")).toBeDisabled();
    await page.getByTestId("share-code").fill(CODE);
    await expect(page.getByTestId("start-collab")).toBeEnabled();
  });

  test("the creator is NOT re-prompted for the code they just set", async ({ page }) => {
    await goLiveWithCode(page, CODE);
    // Lands straight in the live editor — no code prompt for the creator.
    await expect(page.getByTestId("toolbar")).toBeVisible();
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
      await expect(a.getByTestId("toolbar")).toBeVisible();

      // A different device/profile has no stored code → it is prompted.
      const b = await joiner.newPage();
      await b.goto(url);
      await expect(codeInput(b), "joiner without the code must be prompted").toBeVisible();
      await expect(
        b.getByTestId("readonly-banner"),
        "the share-code prompt must not show a separate write-status banner",
      ).toHaveCount(0);
      // While refused, the editor isn't mounted (download etc. are App chrome
      // and stay visible — the meaningful signal is no rendered document).
      await expect(b.locator(PAGE)).toHaveCount(0);

      // The correct code lets the joiner in.
      await codeInput(b).fill(CODE);
      await b.getByTestId("join-submit").click();
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
      await expect(a.getByTestId("toolbar")).toBeVisible();

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
    await expect(page.getByTestId("toolbar")).toBeVisible();
    await page.getByTestId("new-document").click();
    await expect(page.getByTestId("local-editor")).toBeVisible();
    await expect(page).not.toHaveURL(/[?&]doc=/);
  });
});
