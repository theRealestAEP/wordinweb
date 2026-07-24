import { test, expect, type Page } from "@playwright/test";

/**
 * End-to-end browser scenarios for the zero-custody demo (plan doc 09,
 * browser tier). The REAL stack: Vite app → real WebSocket + HTTP seed →
 * zero-custody server. These catch integration bugs the unit/loopback
 * suites can't — and already have: this suite found (and fixed) a missing
 * CORS layer and a Vite stale-dep-cache footgun, and it CAUGHT a real
 * browser-only collaboration bug (see the `fixme` block at the bottom).
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

async function createDoc(page: Page, mode: "encrypted" | "plaintext"): Promise<string> {
  await page.goto(LANDING);
  const label = mode === "encrypted" ? "New encrypted document" : "New plaintext document";
  const createBtn = page.getByRole("button", { name: label });
  await expect(createBtn).toBeVisible();
  await createBtn.click();
  await expect(page).toHaveURL(/[?&]doc=/); // the magic-link capability is in the URL
  await expect(page.locator(PAGE)).toBeVisible(); // editor ready (past ConnectingNotice)
  return page.url();
}

async function typeInEditor(page: Page, text: string): Promise<void> {
  await page.locator(PAGE).first().click();
  await page.keyboard.type(text, { delay: 15 });
}
async function paintedText(page: Page): Promise<string> {
  return (await page.locator(PAGE).first().textContent()) ?? "";
}

test.describe("zero-custody demo — browser E2E", () => {
  test("create a plaintext doc: URL carries the capability, editor mounts, creator is owner", async ({ page }) => {
    const url = await createDoc(page, "plaintext");
    expect(url).toMatch(/[?&]doc=/);
    expect(url).not.toContain("#k="); // plaintext: no key in the fragment
    // The creator holds the owner token → the read-only control is present
    // (the whole seed→owner-token→admin-UI wiring, exercised in the browser).
    await expect(page.getByTestId("readonly-toggle")).toBeVisible();
    await expect(page.getByTestId("download")).toBeEnabled();
  });

  test("create an encrypted doc: the key rides the URL fragment, encrypted badge shows", async ({ page }) => {
    const url = await createDoc(page, "encrypted");
    expect(url).toMatch(/#k=[A-Za-z0-9_-]+$/); // doc-13 §1 fragment key (never sent to the server)
    await expect(page.getByText("encrypted", { exact: true })).toBeVisible();
  });

  test("typing paints locally in a real browser (editor optimistic apply, the dead-typing path)", async ({ page }) => {
    // NOTE: this verifies the editor's LOCAL optimistic apply over the real
    // keydown path — not the server round-trip (which the two-participant
    // fixme below exercises and which currently fails; see its comment).
    await createDoc(page, "encrypted");
    await typeInEditor(page, "hello e2e");
    await expect.poll(() => paintedText(page), { message: "typed text should paint" }).toContain("hello e2e");
  });

  test("Enter then more typing keeps painting (the Enter-desync path, local)", async ({ page }) => {
    await createDoc(page, "encrypted");
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
 * KNOWN BUG — captured, deterministic, NOT yet fixed (see BUGS.md). In a
 * real browser, typing into a freshly-created (blank) collab doc does not
 * EMIT an intent, so nothing round-trips and remote participants never see
 * the edit. Root cause (narrowed via this harness): the editor's caret
 * lands on run/paragraph XmlElements that are reachable via
 * `doc.findParentOf` but are NOT in `doc.editableRoots()`, so
 * `StableIds.encodeCaret` returns null → `onIntent` never fires → no
 * `submit` frame (verified by a browser-native WebSocket spy). The
 * unit/loopback suites miss it because they call `connection.submit`
 * directly, bypassing the editor→onIntent→encode path, and jsdom's stubbed
 * geometry places the caret differently than a real browser. Un-fixme when
 * the editor fix lands.
 */
test.describe("zero-custody demo — E2E (blocked on the known emit bug)", () => {
  test.fixme("two participants: an edit in one appears in the other (true server round-trip)", async ({ browser }) => {
    const ctxA = await browser.newContext();
    const ctxB = await browser.newContext();
    const a = await ctxA.newPage();
    const b = await ctxB.newPage();
    try {
      const url = await createDoc(a, "encrypted");
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
      const url = await createDoc(owner, "encrypted");
      await editor.goto(url);
      await expect(editor.locator(PAGE)).toBeVisible();
      await expect(editor.getByTestId("readonly-toggle")).toHaveCount(0);
      await owner.getByTestId("readonly-toggle").click();
      await expect(owner.getByTestId("readonly-toggle")).toContainText("Read-only ON");
      await typeInEditor(editor, "BLOCKED-EDIT");
      await owner.waitForTimeout(500);
      expect(await paintedText(editor)).not.toContain("BLOCKED-EDIT");
      await typeInEditor(owner, "OWNER-WRITE");
      await expect.poll(() => paintedText(owner)).toContain("OWNER-WRITE");
      await owner.getByTestId("readonly-toggle").click();
      await typeInEditor(editor, "NOW-ALLOWED");
      await expect.poll(() => paintedText(editor)).toContain("NOW-ALLOWED");
    } finally {
      await ctxOwner.close();
      await ctxEditor.close();
    }
  });
});
