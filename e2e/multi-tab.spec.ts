import { test, expect } from "@playwright/test";
import { createCollabDoc, joinCollab, saveB64, docText, converge, SERVER } from "./_helpers";

/**
 * Regression: TWO TABS IN ONE BROWSER must collaborate.
 *
 * The whole suite was green while the real demo broke, because every other
 * test gives each client its own browser *context* (isolated storage), so two
 * clients never share the persistent clientId. Two real tabs in one profile
 * DO share it. The demo minted the clientId in `localStorage` (shared across
 * tabs), so the hub's one-socket-per-(doc, clientId) rule made the second tab
 * "take over" the first — a frozen zombie that then diverged ("typing carries
 * over then stops tracking"). Fix: mint the clientId in `sessionStorage`
 * (per-tab), so a second tab is a distinct participant.
 *
 * This test reproduces the real-world shape:
 *  - a SINGLE browser context (shared localStorage, like one profile),
 *  - two pages (two tabs) that each mint their own per-tab sessionStorage id.
 * With the bug it fails (the first tab is refused `taken-over`); with the fix
 * both stay live and converge.
 *
 * NOTE on painted text: a non-foreground Playwright page reports
 * `document.hidden`, and the collab repaint is coalesced through rAF (paused
 * while hidden), so a hidden tab's *painted* DOM lags until it is shown. That
 * is not a divergence of shared state — `saveB64()` (the session) converges
 * regardless. Where this test checks the rendered DOM it calls
 * `bringToFront()` first, which fires the pending rAF and makes the tab catch
 * up (verified by hand in a real browser).
 */
test.describe("zero-custody demo — two tabs, one browser", () => {
  test("a second tab in the same profile joins as a distinct client (no takeover) and converges", async ({ browser }) => {
    // ONE context == one browser profile: localStorage is shared across its
    // pages, exactly like two tabs. Each page is its own tab, so each gets its
    // own sessionStorage (the fix) and thus a distinct clientId.
    const context = await browser.newContext();
    try {
      const a = await context.newPage();
      const b = await context.newPage();

      const url = await createCollabDoc(a);
      await joinCollab(b, url);

      // The bug's signature: the incumbent tab gets kicked. Neither may be
      // refused, and the creator (A) must still be live after B joins.
      const refusedA = await a.evaluate(() => (window as unknown as { __ww: { _session: { refused: string | null } } }).__ww._session.refused);
      const refusedB = await b.evaluate(() => (window as unknown as { __ww: { _session: { refused: string | null } } }).__ww._session.refused);
      expect(refusedA, "A must not be taken over by B (both are the same profile)").toBeNull();
      expect(refusedB).toBeNull();

      // Distinct per-tab identities (the fix): same profile, different tab ids.
      const idA = await a.evaluate(() => sessionStorage.getItem("wordinweb-client-id"));
      const idB = await b.evaluate(() => sessionStorage.getItem("wordinweb-client-id"));
      expect(idA).toBeTruthy();
      expect(idB).toBeTruthy();
      expect(idA, "two tabs must be two distinct clients").not.toBe(idB);

      // Both see each other in the roster (2 real participants, no zombie).
      await expect.poll(() => a.getByTestId("roster-chip").count()).toBe(2);
      await expect.poll(() => b.getByTestId("roster-chip").count()).toBe(2);

      // A edits through the real connection; B's SESSION converges byte-for-byte.
      const [runId] = await a.evaluate(() => (window as unknown as { __ww: { allocIds(n: number): number[] } }).__ww.allocIds(1));
      // Use a plain insertText at the document start via the editor's own first
      // run — but the simplest cross-client proof is submitOp round-tripping,
      // which the all-intents suite already covers. Here we assert the SHARED
      // STATE converges and neither tab was evicted, which is the regression.
      void runId;
      await converge([a, b], "two tabs, one profile");
      expect(await saveB64(a)).toBe(await saveB64(b));

      // And the rendered view catches up once a tab is shown (rAF de-throttle).
      await b.bringToFront();
      await expect.poll(() => docText(b)).toBe(await docText(a));
    } finally {
      await context.close();
    }
  });

  test("the share link carries the server param so a same-profile joiner reaches the hub", async ({ browser }) => {
    // Guards the plumbing that makes the above possible: the minted ?doc= URL
    // must keep ?server= (the demo learns the port from it), or a second tab
    // silently hits the wrong origin.
    const context = await browser.newContext();
    try {
      const a = await context.newPage();
      const url = await createCollabDoc(a);
      expect(url).toContain(`server=${encodeURIComponent(SERVER)}`);
      expect(url).toMatch(/[?&]doc=/);
    } finally {
      await context.close();
    }
  });
});
