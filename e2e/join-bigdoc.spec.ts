import { test, expect, type Page } from "@playwright/test";
import { zipSync, strToU8 } from "fflate";
import { LANDING, BOARD_CODE, goLive, PAGE, waitHook } from "./_helpers";

/**
 * JOINING A SESSION ON A BIG DOCUMENT — the "click-to-session took ~7s, of
 * which ~6.8s was the collab editor mounting" report, as a REPEATABLE phase
 * benchmark. Prints STRESS-METRIC lines for perf-report.mjs:
 *
 *   join-welcome  (from the encrypted connection, armed via __dxwPerf):
 *                 deriveMs / openMs / mirrorParseMs / replicaParseMs / tailMs
 *   join-mount    (measured here): codeToFirstPageMs — share-code submit to
 *                 the first painted page — plus the mount paint's layout and
 *                 render ms from __dxwPerf.last.
 *
 * Two of the phases are the ones the fixes target:
 *  - CollabEditor used to run a full session.doc.save() (a whole-document
 *    serialize) between the welcome and the first paint, purely to fill a
 *    `source` prop the live view never parses. Removed; the delta shows up
 *    in codeToFirstPageMs.
 *  - The checkpoint is parsed TWICE (mirror + replica), visible as the two
 *    *ParseMs phases — reported so the cost stays measurable, since halving
 *    it means handing one parsed document across, which is a correctness
 *    trade the numbers must justify first.
 */

const DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const PARAS = 4000;

function bigDocx(paras: number): Buffer {
  let seed = 11;
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

test("joining a big-document session reports its mount phases", async ({ page, browser }) => {
  test.setTimeout(420_000);
  await arm(page);
  const reprint = (p: Page, tag: string) =>
    p.on("console", (m) => {
      const text = m.text();
      if (text.startsWith("STRESS-METRIC")) console.log(`${text} role=${tag} paragraphs=${PARAS} env=browser`);
    });
  reprint(page, "owner");

  await page.goto(LANDING);
  await expect(page.getByTestId("local-editor")).toBeVisible();
  await page.locator('input[type="file"][accept*="docx"]').setInputFiles({
    name: "big.docx", mimeType: DOCX_MIME, buffer: bigDocx(PARAS),
  });
  await expect(page.locator(PAGE).first()).toBeVisible({ timeout: 180_000 });
  await expect
    .poll(() => page.locator(PAGE).count(), { message: "big doc never paginated", timeout: 180_000 })
    .toBeGreaterThan(50);
  const url = await goLive(page);

  // Joiner: fresh context, so the share-code prompt is part of the flow.
  const ctx = await browser.newContext();
  const joiner = await ctx.newPage();
  await arm(joiner);
  const welcomeLines: string[] = [];
  joiner.on("console", (m) => {
    const text = m.text();
    if (text.startsWith("STRESS-METRIC")) {
      if (text.startsWith("STRESS-METRIC join-welcome")) welcomeLines.push(text);
      console.log(`${text} role=joiner paragraphs=${PARAS} env=browser`);
    } else if (m.type() === "error" || m.type() === "warning") {
      console.log(`[joiner ${m.type()}] ${text.slice(0, 300)}`);
    }
  });

  await joiner.goto(url);
  const prompt = joiner.getByTestId("join-share-code");
  await expect(prompt).toBeVisible({ timeout: 15_000 });
  await prompt.fill(BOARD_CODE);
  const t0 = Date.now();
  await joiner.getByTestId("join-submit").click();
  await expect(joiner.locator(PAGE).first()).toBeVisible({ timeout: 120_000 });
  const codeToFirstPageMs = Date.now() - t0;
  await waitHook(joiner);

  // The MOUNT paint's breakdown (recorded once by DocxView when armed;
  // `last` is overwritten by later incremental paints).
  const mount = await joiner.evaluate(
    () => ({ ...((window as unknown as { __dxwPerf?: { mount?: Record<string, number> } }).__dxwPerf?.mount ?? {}) }),
  );
  console.log(
    `STRESS-METRIC join-mount paragraphs=${PARAS} env=browser codeToFirstPageMs=${codeToFirstPageMs} ` +
      `layoutMs=${(mount.layout ?? 0).toFixed(0)} renderMs=${(mount.render ?? 0).toFixed(0)} totalPages=${mount.totalPages ?? 0}`,
  );

  // The phase line must exist — this is what keeps the join measurable
  // rather than anecdotal (and it only prints when the parse phases ran).
  expect(welcomeLines.length, "the encrypted welcome must report its phase timings").toBeGreaterThanOrEqual(1);

  await ctx.close();
});
