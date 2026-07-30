import { test, expect } from "@playwright/test";
import { zipSync, strToU8 } from "fflate";
import { LANDING, BOARD_CODE } from "./_helpers";

/**
 * GO-LIVE ON A BIG DOCUMENT — the "Make collaborative on the NIH document
 * chugs" report, in a real browser.
 *
 * Drives the demo's actual pipeline (open big doc → Make collaborative →
 * Start Collab) and captures the STRESS-METRIC lines the app now emits
 * (golive-serialize from the local editor, golive-seal from the seal/PUT
 * flow), re-printing them to stdout so perf-report.mjs can ingest them from
 * this suite's log. Also pins the UX contract: the progress overlay is on
 * screen during the pipeline — a click that freezes the tab with no feedback
 * is the failure mode this exists to prevent.
 *
 * ~9000 paragraphs (~245 pages) rather than the full 500: pagination on
 * LOAD dominates the test budget, and the go-live pipeline itself scales
 * with bytes, which scripts/bench-golive.mjs measures at full size.
 */

const DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const PARAS = 9000;
const PAGE = ".dxw-page";

function bigDocx(paras: number): Buffer {
  // Seeded word soup (matches bench-golive.mjs): repetitive prose zips ~40:1
  // and would understate the serialise/seal byte volume by an order of magnitude.
  let seed = 42;
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

test("going collaborative on a ~250-page document shows progress and reports phase timings", async ({ page }) => {
  test.setTimeout(300_000);

  // Re-print the app's own metric lines so this suite's log is ingestible.
  const metricLines: string[] = [];
  page.on("console", (msg) => {
    const text = msg.text();
    if (text.startsWith("STRESS-METRIC golive-")) {
      metricLines.push(text);
      console.log(text + ` paragraphs=${PARAS} env=browser`);
    }
  });

  await page.goto(LANDING);
  await expect(page.getByTestId("local-editor")).toBeVisible();
  await page.locator('input[type="file"][accept*="docx"]').setInputFiles({
    name: "big.docx", mimeType: DOCX_MIME, buffer: bigDocx(PARAS),
  });
  await expect(page.locator(PAGE).first()).toBeVisible({ timeout: 120_000 });
  await expect
    .poll(() => page.locator(PAGE).count(), { message: "big doc never finished paginating", timeout: 180_000 })
    .toBeGreaterThan(100);

  await page.getByTestId("make-collaborative").click();
  await expect(page.getByTestId("collab-modal")).toBeVisible();
  await page.getByTestId("share-code").fill(BOARD_CODE);
  const clickedAt = Date.now();
  await page.getByTestId("start-collab").click();

  // THE UX CONTRACT: real progress on screen during the pipeline — never a
  // frozen tab and a silent gap. The overlay outlives the pipeline's first
  // phases by design (it stays up until the collab screen replaces this one),
  // so it is stably observable here.
  await expect(page.getByTestId("golive-progress")).toBeVisible();

  // The pipeline completes into a real session.
  await expect(page).toHaveURL(/[?&]doc=/, { timeout: 60_000 });
  await expect(page).toHaveURL(/#k=/);
  await expect(page.getByTestId("toolbar")).toBeVisible();
  const wallMs = Date.now() - clickedAt;
  console.log(`STRESS-METRIC golive-wall paragraphs=${PARAS} env=browser clickToSessionMs=${wallMs}`);

  // Both halves of the pipeline reported numbers.
  expect(metricLines.some((l) => l.includes("golive-serialize")), "the serialise phase must be measured").toBe(true);
  expect(metricLines.some((l) => l.includes("golive-seal")), "the seal/upload phases must be measured").toBe(true);
});
