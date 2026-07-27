import { test, expect } from "@playwright/test";
import { createCollabDoc, joinCollab, converge, saveB64, PAGE } from "./_helpers";

/**
 * Live-repro pin for the second user-reported bug (2026-07-24): "headers and
 * footers can't be opened or edited in collab mode."
 *
 * Three things were wrong at once. The header/footer ribbon group was gated
 * off in a room; creating the part was a local-only structural mutation with
 * no wire form; and header paragraphs — though they carry stable ids — were
 * invisible to the apply's run index, so every keystroke in a header was a
 * clean reject everywhere except the typist's own screen.
 *
 * The gesture here is the Word-native one (double-click the top margin band),
 * which is also the path that could fork a room even with the ribbon gated.
 */
test.describe("zero-custody demo — header/footer editing", () => {
  test("double-clicking the margin band opens the header; typing replicates", async ({ browser }) => {
    const ctxA = await browser.newContext();
    const ctxB = await browser.newContext();
    const a = await ctxA.newPage();
    const b = await ctxB.newPage();
    try {
      const url = await createCollabDoc(a);
      await joinCollab(b, url);
      await expect.poll(() => b.getByTestId("roster-chip").count()).toBe(2);
      const before = await saveB64(a);

      // Double-click well inside the top margin band (above the body's first
      // line): Word enters header editing, creating the part if there is none
      // — and the demo's blank document has none.
      const box = (await a.locator(PAGE).first().boundingBox())!;
      await a.mouse.dblclick(box.x + 120, box.y + 25);

      // The header chrome really engaged (guards against "the click did
      // nothing", which is how this looked to the user).
      await expect(a.locator(".dxw-hf-mode")).toHaveCount(1);
      await a.keyboard.type("Acme Report", { delay: 15 });

      // The part was created and the text landed in it — on BOTH clients,
      // byte-identically. (document.xml alone would show neither.)
      await expect.poll(() => saveB64(a), { timeout: 5000 }).not.toBe(before);
      await expect.poll(() => headerText(a), { timeout: 5000 }).toContain("Acme Report");
      await expect.poll(() => headerText(b), { timeout: 8000 }).toContain("Acme Report");
      await converge([a, b], "header typing");
    } finally {
      await ctxA.close();
      await ctxB.close();
    }
  });
});

/** Text held by the live replica's header part(s). */
async function headerText(page: import("@playwright/test").Page): Promise<string> {
  return page.evaluate(() => {
    const doc = (window as any).__ww._session.doc;
    const text = (e: any): string =>
      (e.name.endsWith(":t") ? e.text : "") + e.children.map(text).join("");
    return doc
      .editableRoots()
      .filter((r: any) => r.name.endsWith("hdr") || r.name.endsWith("ftr"))
      .map(text)
      .join("|");
  });
}
