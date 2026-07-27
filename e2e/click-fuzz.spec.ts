import { test, expect, type Browser } from "@playwright/test";
import { createCollabDoc, joinCollab, converge, docText, PAGE } from "./_helpers";

/**
 * RANDOM-CLICK typing fuzz (added after live desyncs the scripted e2e kept
 * missing): real users click ARBITRARY page points — mid-text, empty lines,
 * far below the last paragraph, margins — then type. Every such gesture must
 * either land as replicated edits or do nothing; a click that mutates only
 * the local doc forks the room permanently. The precise-click helpers
 * (clickTextStart) structurally excluded this whole class — this suite is
 * the guard for it. First run caught the click-below fork on round 1: the
 * whitespace click-and-type path spliced 23 filler paragraphs locally with
 * no emission (now gated in collab).
 *
 * Each round: a random point in the page, a typed marker, sometimes Enter —
 * then BYTE-IDENTICAL convergence. Seeded PRNG; the seed is logged so a
 * failure replays exactly.
 */
async function fuzzRun(browser: Browser, seed: number, enterChance: number): Promise<void> {
  let s = seed >>> 0;
  const rand = () => ((s = (s * 1664525 + 1013904223) >>> 0) / 0x100000000);
  console.log(`click-fuzz seed: ${seed}`);

  const ctxA = await browser.newContext();
  const ctxB = await browser.newContext();
  const a = await ctxA.newPage();
  const b = await ctxB.newPage();
  try {
    const url = await createCollabDoc(a);
    await joinCollab(b, url);
    await expect.poll(() => b.getByTestId("roster-chip").count()).toBe(2);

    const boxA = (await a.locator(PAGE).first().boundingBox())!;
    const boxB = (await b.locator(PAGE).first().boundingBox())!;

    const ROUNDS = 14;
    for (let i = 0; i < ROUNDS; i++) {
      const useA = i % 2 === 0;
      const page = useA ? a : b;
      const box = useA ? boxA : boxB;
      // Anywhere on the page sheet: from the top margin to well below any
      // content (the sheet is 1056px tall; stay in the visible viewport).
      const x = box.x + 15 + rand() * (box.width - 40);
      const y = box.y + 15 + rand() * Math.min(box.height - 30, 620);
      await page.mouse.click(x, y);
      if (rand() < enterChance) await page.keyboard.press("Enter");
      // Cmd/Ctrl+Enter (page break) — a real-user gesture that was a silent
      // local-only mutation until wired; keep it in the gesture mix.
      if (rand() < 0.15) await page.keyboard.press(process.platform === "darwin" ? "Meta+Enter" : "Control+Enter");
      await page.keyboard.type(`r${i}${useA ? "a" : "b"} `, { delay: 10 });
      await converge([a, b], `round ${i} (${useA ? "A" : "B"} clicked ${Math.round(x - box.x)},${Math.round(y - box.y)})`);
    }

    // The doc must not still be empty: at least the rounds that clicked on
    // or near content landed — a fully dead run means clicks are broken.
    expect((await docText(a)).length).toBeGreaterThan(0);
  } finally {
    await ctxA.close();
    await ctxB.close();
  }
}

test.describe("zero-custody demo — random-click typing fuzz", () => {
  test("random clicks + typing from both clients never fork the doc (seed 1)", async ({ browser }) => {
    test.setTimeout(120_000);
    await fuzzRun(browser, 0xC0FFEE ^ 20260724, 0.25);
  });

  test("random clicks + typing never fork the doc (seed 2, Enter-heavy)", async ({ browser }) => {
    test.setTimeout(120_000);
    await fuzzRun(browser, 0xBADF00D, 0.5);
  });
});
