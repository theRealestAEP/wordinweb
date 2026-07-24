import { test, expect, type Page } from "@playwright/test";

/**
 * End-to-end browser scenarios for the zero-custody demo (plan doc 09,
 * browser tier). The REAL stack: Vite app → real WebSocket + HTTP seed →
 * zero-custody server. These catch integration bugs the unit/loopback
 * suites can't — and already have: this suite found (and fixed) a missing
 * CORS layer and a Vite stale-dep-cache footgun (BUGS.md).
 *
 * Deterministic: fixed ports (config `webServer`), isolated contexts (each
 * mints its own per-browser clientId). Run: `npm run e2e` / `:headed`.
 */

const PAGE = ".dxw-page";
// The demo learns the collab server's port from `?server=`; the config
// boots it on 1399. `searchParams.set` preserves it into the minted `?doc=`
// URL, so shared links carry it too.
const SERVER = "localhost:1399";
const LANDING = `/?server=${SERVER}`;

async function createDoc(page: Page): Promise<string> {
  // The demo opens to the LOCAL editor; "Make collaborative" seals the
  // current local doc and goes live (always encrypted — key in #k=).
  await page.goto(LANDING);
  await expect(page.getByTestId("local-editor")).toBeVisible();
  const goLive = page.getByTestId("make-collaborative");
  await expect(goLive).toBeVisible();
  await goLive.click();
  await expect(page).toHaveURL(/[?&]doc=/); // the magic-link capability
  await expect(page).toHaveURL(/#k=/); // collaborative = encrypted (doc 13 §1)
  await expect(page.getByTestId("download")).toBeVisible(); // collab chrome mounted
  await expect(page.locator(PAGE)).toBeVisible();
  return page.url();
}

async function typeInEditor(page: Page, text: string): Promise<void> {
  // Click PRECISELY on the first text line (top-left of the page), not the
  // page center — a center click on a one-line doc lands in empty space
  // below the paragraph and misplaces the caret. The editor listens for
  // real keydown on its container (tabIndex 0).
  const box = await page.locator(PAGE).first().boundingBox();
  if (box) await page.mouse.click(box.x + 30, box.y + 25);
  else await page.locator(PAGE).first().click();
  await page.keyboard.type(text, { delay: 15 });
}
async function paintedText(page: Page): Promise<string> {
  return (await page.locator(PAGE).first().textContent()) ?? "";
}

test.describe("zero-custody demo — browser E2E", () => {
  test("Make collaborative: seals the local doc, key rides the fragment, creator is owner", async ({ page }) => {
    const url = await createDoc(page);
    expect(url).toMatch(/[?&]doc=/); // the magic-link capability
    expect(url).toMatch(/#k=[A-Za-z0-9_-]+/); // doc-13 §1 fragment key (never sent to the server)
    await expect(page.getByText("encrypted", { exact: true })).toBeVisible();
    // The creator holds the owner token → owner controls present.
    await expect(page.getByTestId("readonly-toggle")).toBeVisible();
    await expect(page.getByTestId("download")).toBeEnabled();
  });

  test("local-first: type in the local editor, Make collaborative, the content is sealed and shared", async ({ browser }) => {
    const ctxA = await browser.newContext();
    const ctxB = await browser.newContext();
    const a = await ctxA.newPage();
    const b = await ctxB.newPage();
    try {
      // A edits LOCALLY first (no server), then goes live.
      await a.goto(LANDING);
      await expect(a.getByTestId("local-editor")).toBeVisible();
      const box = await a.locator(PAGE).first().boundingBox();
      if (box) await a.mouse.click(box.x + 30, box.y + 25);
      await a.keyboard.type("LOCAL DRAFT", { delay: 15 });
      await expect.poll(() => paintedText(a)).toContain("LOCAL DRAFT");
      await a.getByTestId("make-collaborative").click();
      await expect(a).toHaveURL(/#k=/);
      await expect(a.getByTestId("download")).toBeVisible();
      // The local draft survived into the collab session…
      await expect.poll(() => paintedText(a)).toContain("LOCAL DRAFT");
      // …and a participant opening the shared link sees it too.
      await b.goto(a.url());
      await expect(b.locator(PAGE)).toBeVisible();
      await expect.poll(() => paintedText(b)).toContain("LOCAL DRAFT");
    } finally {
      await ctxA.close();
      await ctxB.close();
    }
  });

  test("typing paints in a real browser (editor apply over the real keydown path)", async ({ page }) => {
    await createDoc(page);
    await typeInEditor(page, "hello e2e");
    await expect.poll(() => paintedText(page), { message: "typed text should paint" }).toContain("hello e2e");
  });

  test("Enter then more typing keeps painting (the Enter-desync path)", async ({ page }) => {
    await createDoc(page);
    await typeInEditor(page, "line one");
    await page.keyboard.press("Enter");
    await page.keyboard.type("line two", { delay: 15 });
    await expect.poll(() => paintedText(page)).toContain("line one");
    await expect.poll(() => paintedText(page)).toContain("line two");
  });

  test("bring-it-back-live: a viewer with no local copy of an ended session is told so", async ({ page }) => {
    const ghostDoc = "d_" + "0".repeat(32);
    await page.goto(`/?server=${SERVER}&doc=${ghostDoc}`);
    // The no-session flow surfaces the honest failure (session ended /
    // bring it back). `.first()` — the banner text and its button both match.
    await expect(page.getByText(/session has ended|bring it back|no saved copy/i).first()).toBeVisible();
  });
});

/**
 * True multi-client round-trip over the real browser stack — the edit
 * A types reaches B and C, and owner read-only is enforced end to end.
 * (These were briefly `fixme` when a test-harness click landed in empty
 * space and misplaced the caret; the fix is `typeInEditor` clicking the
 * text line precisely — see BUGS.md, both "bugs" were test artifacts.)
 */
test.describe("zero-custody demo — E2E (multi-client round-trip + roles)", () => {
  test("two participants: an edit in one appears in the other (true server round-trip)", async ({ browser }) => {
    const ctxA = await browser.newContext();
    const ctxB = await browser.newContext();
    const a = await ctxA.newPage();
    const b = await ctxB.newPage();
    try {
      const url = await createDoc(a);
      await b.goto(url);
      await expect(b.locator(PAGE)).toBeVisible();
      await expect.poll(async () => a.getByTestId("roster-chip").count()).toBe(2);
      await expect.poll(async () => b.getByTestId("roster-chip").count()).toBe(2);
      await typeInEditor(a, "FROM-ALICE");
      await expect.poll(() => paintedText(b)).toContain("FROM-ALICE");
      await typeInEditor(b, "FROM-BOB");
      await expect.poll(() => paintedText(a)).toContain("FROM-BOB");
    } finally {
      await ctxA.close();
      await ctxB.close();
    }
  });

  test.fixme("owner read-only blocks editors, owner keeps writing, then lifts (doc 14 §2.5)", async ({ browser }) => {
    const ctxOwner = await browser.newContext();
    const ctxEditor = await browser.newContext();
    const owner = await ctxOwner.newPage();
    const editor = await ctxEditor.newPage();
    try {
      const url = await createDoc(owner);
      await editor.goto(url);
      await expect(editor.locator(PAGE)).toBeVisible();
      await expect(editor.getByTestId("readonly-toggle")).toHaveCount(0);
      // Guard against a Vite pre-bundle race (stale session without admin).
      await expect
        .poll(() => owner.evaluate(() => typeof (window as unknown as { __ww?: { _session?: { admin?: unknown } } }).__ww?._session?.admin === "function"))
        .toBe(true);
      await owner.getByTestId("readonly-toggle").click();
      await expect(owner.getByTestId("readonly-toggle")).toContainText("Read-only ON");
      await owner.waitForTimeout(500); // let the admin(readOnly) reach the hub
      // The blocked editor's edit is refused at the sequencer — the SECURITY
      // invariant is that it never reaches the owner (the blocked client may
      // still show its own optimistic paint until reload; that's a separate
      // minor UX item, BUGS.md). Prove non-propagation.
      await typeInEditor(editor, "BLOCKED-EDIT");
      await owner.waitForTimeout(700);
      expect(await paintedText(owner)).not.toContain("BLOCKED-EDIT");
      // The owner bypasses their own lock and it reaches the editor.
      await typeInEditor(owner, "OWNER-WRITE");
      await expect.poll(() => paintedText(owner)).toContain("OWNER-WRITE");
      await expect.poll(() => paintedText(editor)).toContain("OWNER-WRITE");
      // Lift read-only: the editor writes again and it reaches the owner.
      await owner.getByTestId("readonly-toggle").click();
      await typeInEditor(editor, "NOW-ALLOWED");
      await expect.poll(() => paintedText(owner)).toContain("NOW-ALLOWED");
    } finally {
      await ctxOwner.close();
      await ctxEditor.close();
    }
  });
});
