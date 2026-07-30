import { test, expect, type Page } from "@playwright/test";
import { zipSync, strToU8 } from "fflate";
import { LANDING, PAGE, goLive, waitHook } from "./_helpers";

/**
 * BIG-DOCUMENT TYPING LATENCY — the "500-page document is uneditable"
 * regression, reproduced the way the user hit it: open a large .docx in the
 * LOCAL editor, click somewhere, type; then make it collaborative and type
 * again.
 *
 * Two regressions are pinned here, both specific to documents past DocxView's
 * background-layout threshold (>50 pages), which is why the 200-paragraph
 * stress suite never caught them:
 *
 *  LOCAL — the first keystroke after every caret move forced a full
 *  doc.refresh(); its modelVersion bump routed rerender into an async
 *  whole-document relayout behind an INERT container (data-dxw-layout-busy),
 *  eating input for seconds after every click-then-type.
 *
 *  COLLAB — DocxView's renderSignal rode CollabSession.version, which ticks
 *  on every onChange including the editor's own submit + echo, so EVERY
 *  keystroke queued that same inert whole-document relayout.
 *
 * The sharp post-fix assertion is the same for both: plain typing (click,
 * then characters) must never put the surface into the layout-busy/inert
 * state. Latency numbers are emitted as STRESS-METRIC lines; the hard bounds
 * on them are deliberately loose (CI machines vary).
 */

const DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const PARAS = 2200; // ~60+ real pages — safely past BACKGROUND_LAYOUT_PAGE_THRESHOLD (50)

function bigDocx(paras: number): Buffer {
  const para = (i: number) =>
    `<w:p><w:r><w:t xml:space="preserve">Paragraph ${i}: the quick brown fox jumps over the lazy dog while the committee deliberates at length. </w:t></w:r></w:p>`;
  let body = "";
  for (let i = 0; i < paras; i++) body += para(i);
  const documentXml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}</w:body></w:document>`;
  const zipped = zipSync({
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
  });
  return Buffer.from(zipped);
}

function metric(scenario: string, fields: Record<string, string | number>): void {
  const body = Object.entries(fields)
    .map(([k, v]) => `${k}=${typeof v === "number" ? (Number.isInteger(v) ? v : v.toFixed(1)) : v}`)
    .join(" ");
  console.log(`STRESS-METRIC ${scenario} ${body}`);
}

/** Count occurrences of the sentinel char in the rendered document. ("Z"
 * does not occur in the generated body text, so each typed Z is countable.) */
async function sentinelCount(page: Page): Promise<number> {
  return page.evaluate(() => (document.querySelector(".dxw-pages")?.textContent?.match(/Z/g) ?? []).length);
}

/** True while the editing surface is in the background-layout (inert) state. */
async function layoutBusy(page: Page): Promise<boolean> {
  return page.evaluate(() => !!document.querySelector("[data-dxw-layout-busy]"));
}

/**
 * Click a text span (a caret JUMP — the regression trigger), then type one
 * sentinel char and measure the time until it is painted. Reports whether the
 * surface ever entered the layout-busy state while we typed.
 */
async function clickThenType(page: Page, round: number): Promise<{ ms: number; busy: boolean }> {
  // A different line each round so every round moves the caret (the
  // regression trigger is the caret JUMP before the keystroke).
  const box = await page.locator(PAGE).first().boundingBox();
  expect(box, "first page must be on screen").toBeTruthy();
  await page.mouse.click(box!.x + 30, box!.y + 25 + ((round * 37) % 300));
  const before = await sentinelCount(page);
  const t0 = Date.now();
  await page.keyboard.type("Z");
  let busy = false;
  await expect
    .poll(
      async () => {
        busy = busy || (await layoutBusy(page));
        return sentinelCount(page);
      },
      { message: `round ${round}: typed char never painted`, timeout: 20_000 },
    )
    .toBeGreaterThan(before);
  return { ms: Date.now() - t0, busy };
}

async function typeRounds(page: Page, scenario: string, rounds: number): Promise<void> {
  const times: number[] = [];
  let busyRounds = 0;
  for (let i = 0; i < rounds; i++) {
    const { ms, busy } = await clickThenType(page, i);
    times.push(ms);
    if (busy) busyRounds++;
  }
  const sorted = [...times].sort((a, b) => a - b);
  metric(scenario, {
    paragraphs: PARAS,
    rounds,
    p50: sorted[Math.floor(sorted.length / 2)],
    max: sorted[sorted.length - 1],
    busyRounds,
  });
  // THE REGRESSION PIN: plain typing must never enter the inert
  // background-layout state — that state is what ate keystrokes for seconds
  // per click-then-type on big documents. (Both pre-fix behaviors trip it on
  // every round here: local via the caret-move refresh, collab via the
  // version-driven renderSignal.)
  expect(busyRounds, `${scenario}: typing must not trigger the inert whole-document relayout`).toBe(0);
  // Loose sanity bound only (CI varies): the pre-fix failure mode was
  // multi-second stalls; a healthy keystroke paints well under a second.
  expect(sorted[sorted.length - 1], `${scenario}: worst keystroke`).toBeLessThan(5_000);
}

test.describe("big document typing (>50 pages)", () => {
  test("local editor and collab editor stay interactive on a 60+ page document", async ({ page }) => {
    test.setTimeout(300_000);

    // ---- LOCAL: open the big docx in the landing editor ----
    await page.goto(LANDING);
    await expect(page.getByTestId("local-editor")).toBeVisible();
    await page.locator('input[type="file"][accept*="docx"]').setInputFiles({
      name: "big.docx",
      mimeType: DOCX_MIME,
      buffer: bigDocx(PARAS),
    });
    await expect(page.locator(PAGE).first()).toBeVisible({ timeout: 60_000 });
    await expect
      .poll(() => page.locator(PAGE).count(), { message: "big doc never finished paginating", timeout: 120_000 })
      .toBeGreaterThan(50);
    // Let the initial background layout fully settle before measuring.
    await expect.poll(() => layoutBusy(page), { timeout: 60_000 }).toBe(false);
    const pages = await page.locator(PAGE).count();
    metric("bigdoc-pages", { paragraphs: PARAS, pages });

    await typeRounds(page, "bigdoc-local-clicktype", 5);

    // ---- COLLAB: the same document, made collaborative ----
    const url = await goLive(page);
    expect(url).toContain("#k=");
    await waitHook(page);
    await expect.poll(() => layoutBusy(page), { timeout: 120_000 }).toBe(false);

    await typeRounds(page, "bigdoc-collab-clicktype", 5);
  });
});
