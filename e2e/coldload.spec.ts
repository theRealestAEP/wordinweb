import { test, expect, type Page } from "@playwright/test";
import { zipSync, strToU8 } from "fflate";
import { LANDING } from "./_helpers";

/**
 * COLD PAGE LOAD, NO SESSION — "initial page load is slow even with no
 * session at all", as a repeatable phase benchmark. Two scenarios:
 *
 *  1. blank    — a fresh browser: fetch /blank, parse, lay out one page.
 *  2. bigdoc   — the autosave slot holds a ~110-page document (the NIH
 *                shape): restore from IndexedDB, parse, lay out the lot.
 *
 * Phases per load (STRESS-METRIC coldload):
 *   networkMs   navigation start → main document responseEnd
 *   domReadyMs  navigation start → DOMContentLoaded (script eval included)
 *   editorMs    navigation start → the local editor shell on screen
 *   firstPageMs navigation start → the first painted page
 *   layoutMs / renderMs — the mount paint's own breakdown (__dxwPerf.mount)
 *
 * Context for the numbers: the same instrument run against the published
 * 0.1.22 (scripts/bench-local-typing.mjs, WW_PKG=…) showed IDENTICAL
 * parse+layout cost and identical module-eval cost, so a slow cold load is
 * app-shell or storage work, not the engine.
 */

const DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const PARAS = 4000;
const PAGE = ".dxw-page";

function bigDocx(paras: number): Buffer {
  let seed = 23;
  const rnd = () => ((seed = (Math.imul(seed, 1103515245) + 12345) & 0x7fffffff) / 0x7fffffff);
  const word = () => {
    let w = "";
    const len = 3 + Math.floor(rnd() * 8);
    for (let i = 0; i < len; i++) w += String.fromCharCode(97 + Math.floor(rnd() * 26));
    return w;
  };
  const para = (i: number) => {
    let text = `Paragraph ${i}:`;
    for (let j = 0; j < 18; j++) text += ` ${word()}`;
    return `<w:p><w:r><w:t xml:space="preserve">${text}. </w:t></w:r></w:p>`;
  };
  let body = "";
  for (let i = 0; i < paras; i++) body += para(i);
  const documentXml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}</w:body></w:document>`;
  return Buffer.from(zipSync({
    "[Content_Types].xml": strToU8(
      `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
        `<Default Extension="xml" ContentType="application/xml"/>` +
        `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
        `<Override PartName="/word/document.xml" ContentType="${DOCX_MIME}.main+xml"/></Types>`,
    ),
    "_rels/.rels": strToU8(
      `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`,
    ),
    "word/document.xml": strToU8(documentXml),
  }));
}

const arm = (page: Page) =>
  page.addInitScript(() => {
    (window as unknown as { __dxwPerf: unknown }).__dxwPerf = { samples: [], jobs: {} };
  });

async function measureLoad(page: Page, scenario: string): Promise<void> {
  const t0 = Date.now();
  await page.goto(LANDING);
  if (scenario === "bigdoc-restore") {
    // THE UX CONTRACT: restoring a big document must show the loading
    // overlay, not a bare frozen page — this is the frame on screen for the
    // seconds the synchronous parse + pagination takes.
    await expect(page.getByTestId("doc-loading")).toBeVisible({ timeout: 10_000 });
  }
  await expect(page.getByTestId("local-editor")).toBeVisible({ timeout: 60_000 });
  const editorMs = Date.now() - t0;
  await expect(page.locator(PAGE).first()).toBeVisible({ timeout: 180_000 });
  const firstPageMs = Date.now() - t0;
  const nav = await page.evaluate(() => {
    const n = performance.getEntriesByType("navigation")[0] as PerformanceNavigationTiming;
    return { responseEnd: n.responseEnd, domReady: n.domContentLoadedEventEnd };
  });
  const mount = await page.evaluate(
    () => ({ ...((window as unknown as { __dxwPerf?: { mount?: Record<string, number> } }).__dxwPerf?.mount ?? {}) }),
  );
  console.log(
    `STRESS-METRIC coldload scenario=${scenario} paragraphs=${scenario === "blank" ? 1 : PARAS} env=browser ` +
      `networkMs=${nav.responseEnd.toFixed(0)} domReadyMs=${nav.domReady.toFixed(0)} editorMs=${editorMs} ` +
      `firstPageMs=${firstPageMs} layoutMs=${(mount.layout ?? 0).toFixed(0)} renderMs=${(mount.render ?? 0).toFixed(0)} ` +
      `totalPages=${mount.totalPages ?? 0}`,
  );
}

test("cold load phases: blank landing and a big autosaved document", async ({ page }) => {
  test.setTimeout(420_000);
  await arm(page);

  // Scenario 1: fresh browser, blank template.
  await measureLoad(page, "blank");

  // Seed the autosave slot with the big document, then flush it.
  await page.locator('input[type="file"][accept*="docx"]').setInputFiles({
    name: "big.docx", mimeType: DOCX_MIME, buffer: bigDocx(PARAS),
  });
  await expect
    .poll(() => page.locator(PAGE).count(), { message: "big doc never paginated", timeout: 180_000 })
    .toBeGreaterThan(50);
  // Mark dirty (autosave only writes after input), then flush via the
  // visibilitychange path the editor already listens on.
  await page.locator(PAGE).first().click({ position: { x: 30, y: 25 } });
  await page.keyboard.type("Z");
  await page.evaluate(() => {
    Object.defineProperty(document, "visibilityState", { value: "hidden", configurable: true });
    document.dispatchEvent(new Event("visibilitychange"));
  });
  // The write is async; give IndexedDB a moment to land it.
  await page.waitForTimeout(1500);

  // Scenario 2: cold load restoring the big document from this browser.
  await measureLoad(page, "bigdoc-restore");

  // The restore must actually have happened (not the blank fallback).
  await expect
    .poll(() => page.locator(PAGE).count(), { timeout: 180_000 })
    .toBeGreaterThan(50);
});
