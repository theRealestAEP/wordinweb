import { test, expect } from "@playwright/test";
import { createCollabDoc, joinCollab, converge, saveB64, PAGE } from "./_helpers";

/**
 * Image upload across real browsers (plan doc 16 §5, the client half).
 *
 * User report: "uploading images doesn't work at all." In a room it truly
 * didn't — the toolbar's image button was gated off entirely, and the api's
 * insert path had no collab branch to gate ON, so un-gating alone would have
 * produced a local-only mutation and a forked room.
 *
 * What this proves end to end: the placer's bytes go over HTTP addressed by
 * their own sha256, only the ADDRESS rides the sequenced intent, and a second
 * browser — which never saw the file — reconstructs identical pixels by
 * fetching that address and re-deriving the hash itself.
 */

/** A real 1×1 PNG: the browser decodes this for real (createImageBitmap). */
const PNG_B64 =
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==";
const PNG = Buffer.from(PNG_B64, "base64");

/** The toolbar's hidden image picker (the icon picker takes .svg only). */
const IMAGE_INPUT = 'input[type="file"][accept*="image/png"]';

/** Media parts present in a client's live package. */
async function mediaParts(page: import("@playwright/test").Page): Promise<string[]> {
  return page.evaluate(() =>
    (window as any).__ww._session.doc.pkg.names().filter((n: string) => n.startsWith("word/media/")),
  );
}

test.describe("zero-custody demo — image upload over the media relay", () => {
  test("A uploads a PNG; B renders it and both packages hold the same bytes", async ({ browser }) => {
    const ctxA = await browser.newContext();
    const ctxB = await browser.newContext();
    const a = await ctxA.newPage();
    const b = await ctxB.newPage();
    try {
      const url = await createCollabDoc(a);
      await joinCollab(b, url);
      await expect.poll(() => b.getByTestId("roster-chip").count()).toBe(2);

      // The button is REAL in a room now (the whole group was gated off
      // before, so there was nothing to click at all). It lives on the
      // ribbon's Insert tab.
      await a.locator('button[data-tab="insert"]').click();
      await expect(a.locator('button[title="Insert image"]')).toBeVisible();

      // Place a caret, then pick a file exactly as the toolbar does.
      const box = (await a.locator(PAGE).first().boundingBox())!;
      await a.mouse.click(box.x + 30, box.y + 25);
      const before = await saveB64(a);
      await a.locator(IMAGE_INPUT).setInputFiles({ name: "dot.png", mimeType: "image/png", buffer: PNG });

      // PLACER: the image is in A's package and painted on A's page.
      await expect.poll(() => mediaParts(a), { timeout: 8000 }).toHaveLength(1);
      await expect.poll(() => saveB64(a), { timeout: 5000 }).not.toBe(before);
      await expect(a.locator(`${PAGE} img`).first()).toBeVisible();

      // PEER: B never saw the file. It reconstructs the pixels by fetching
      // the committed address — and the packages become byte-identical,
      // which for a docx means the media part matched too.
      await expect.poll(() => mediaParts(b), { timeout: 10000 }).toHaveLength(1);
      await expect(b.locator(`${PAGE} img`).first()).toBeVisible();
      await converge([a, b], "image upload");

      // No skeleton is left behind once the bytes have landed.
      await expect(b.locator(".dxw-media-skeleton")).toHaveCount(0);
    } finally {
      await ctxA.close();
      await ctxB.close();
    }
  });

  test("a LATE joiner receives the image it never saw uploaded", async ({ browser }) => {
    // Closed from both ends this arc: the joiner's snapshot carries the
    // image's REGISTRATION but not its declared sha (that address lives only
    // in intents already folded into the snapshot), so a joiner used to
    // reserve the box and never fill it. The welcome now carries the address
    // map — plaintext rooms in `welcome.media`, encrypted rooms (this demo)
    // inside the SEALED checkpoint body, so a blind server still learns
    // nothing about the document's part structure.
    const ctxA = await browser.newContext();
    const ctxC = await browser.newContext();
    const a = await ctxA.newPage();
    const c = await ctxC.newPage();
    try {
      const url = await createCollabDoc(a);
      const box = (await a.locator(PAGE).first().boundingBox())!;
      await a.mouse.click(box.x + 30, box.y + 25);
      await a.locator(IMAGE_INPUT).setInputFiles({ name: "dot.png", mimeType: "image/png", buffer: PNG });
      await expect.poll(() => mediaParts(a), { timeout: 8000 }).toHaveLength(1);

      // Carol joins only now.
      await joinCollab(c, url);
      await expect.poll(() => c.getByTestId("roster-chip").count()).toBe(2);
      await expect.poll(() => mediaParts(c), { timeout: 10000 }).toHaveLength(1);
      await expect(c.locator(`${PAGE} img`).first()).toBeVisible();
      await expect(c.locator(".dxw-media-skeleton")).toHaveCount(0);
      await converge([a, c], "late joiner image");
    } finally {
      await ctxA.close();
      await ctxC.close();
    }
  });
});
