import { test, expect, type Page } from "@playwright/test";
import { deflateSync } from "node:zlib";
import { BOARD_CODE, LANDING, PAGE, enterCodeIfPrompted, waitHook } from "./_helpers";

/**
 * USER-REPORTED (live demo): insert a photo in the LOCAL editor, go
 * collaborative, then MOVE the image — and it disappears. Plus a second,
 * unconfirmed report that the peer never sees the image at all.
 *
 * This is the flow with no coverage anywhere else: every other media spec
 * uploads while ALREADY collaborative, so the bytes ride the relay. Here the
 * image is inserted before go-live, which means its bytes travel inside the
 * sealed genesis checkpoint and the relay is never involved at insert time.
 *
 * Real geometry is the point of doing this in a browser: the drag path reads
 * getBoundingClientRect and the laid-out item box, and a jsdom reproduction of
 * the same steps converges happily because every rect there is zero.
 */

/** A real, decodable PNG of the given size — the browser must actually decode
 * it (createImageBitmap), so a fake byte string will not do. */
function makePng(width: number, height: number): Buffer {
  const crcTable: number[] = [];
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    crcTable[n] = c >>> 0;
  }
  const crc32 = (buf: Buffer): number => {
    let c = 0xffffffff;
    for (const byte of buf) c = crcTable[(c ^ byte) & 0xff] ^ (c >>> 8);
    return (c ^ 0xffffffff) >>> 0;
  };
  const chunk = (type: string, data: Buffer): Buffer => {
    const len = Buffer.alloc(4);
    len.writeUInt32BE(data.length);
    const body = Buffer.concat([Buffer.from(type, "ascii"), data]);
    const crc = Buffer.alloc(4);
    crc.writeUInt32BE(crc32(body));
    return Buffer.concat([len, body, crc]);
  };
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(width, 0);
  ihdr.writeUInt32BE(height, 4);
  ihdr[8] = 8; // bit depth
  ihdr[9] = 2; // truecolor RGB
  const raw = Buffer.alloc(height * (1 + width * 3));
  for (let y = 0; y < height; y++) {
    const row = y * (1 + width * 3);
    raw[row] = 0; // filter: none
    for (let x = 0; x < width; x++) {
      const p = row + 1 + x * 3;
      raw[p] = (x * 255) / width;        // a gradient, so the image is visibly real
      raw[p + 1] = (y * 255) / height;
      raw[p + 2] = 128;
    }
  }
  return Buffer.concat([
    Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    chunk("IHDR", ihdr),
    chunk("IDAT", deflateSync(raw)),
    chunk("IEND", Buffer.alloc(0)),
  ]);
}

/** The Pictures file input, which the local editor's toolbar owns. */
async function insertPhotoLocally(page: Page): Promise<void> {
  const input = page.locator('input[type="file"][accept*="image/png"]');
  await input.setInputFiles({ name: "photo.png", mimeType: "image/png", buffer: makePng(800, 450) });
}

const imgCount = (page: Page) => page.locator(`${PAGE} img`).count();
const skeletonCount = (page: Page) => page.locator("[data-dxw-media-state]").count();

test.describe("user-reported: a photo inserted before go-live", () => {
  test("the peer receives a checkpoint-borne image, and MOVING it does not lose it", async ({ browser }) => {
    const ctxA = await browser.newContext();
    const ctxB = await browser.newContext();
    const a = await ctxA.newPage();
    const b = await ctxB.newPage();
    try {
      // ---- the user's flow, exactly: type, insert a photo, THEN go live.
      await a.goto(LANDING);
      await expect(a.getByTestId("local-editor")).toBeVisible();
      const box = await a.locator(PAGE).first().boundingBox();
      if (box) await a.mouse.click(box.x + 30, box.y + 25);
      await a.keyboard.type("test", { delay: 15 });
      await insertPhotoLocally(a);
      await expect.poll(() => imgCount(a), { message: "the photo should paint locally" }).toBe(1);

      await a.getByTestId("make-collaborative").click();
      await expect(a.getByTestId("collab-modal")).toBeVisible();
      await a.getByTestId("share-code").fill(BOARD_CODE);
      await a.getByTestId("start-collab").click();
      await expect(a).toHaveURL(/[?&]doc=/);
      await expect(a.getByTestId("download")).toBeVisible();
      await waitHook(a);
      // The photo survived the seal.
      await expect.poll(() => imgCount(a), { message: "the photo should survive go-live" }).toBe(1);

      // ---- Q1: the peer. Bytes rode the CHECKPOINT; the relay never had them.
      await b.goto(a.url());
      await enterCodeIfPrompted(b);
      await expect(b.locator(PAGE).first()).toBeVisible();
      await waitHook(b);
      await expect
        .poll(() => imgCount(b), { message: "peer must render an image that arrived via the checkpoint" })
        .toBe(1);
      expect(await skeletonCount(b), "peer must not be stuck on a skeleton").toBe(0);

      // ---- Q2: select the image and MOVE it. The move rides the wire as
      // setImageWrap + setFloatingPagePosition (inline drawings become
      // floating on drag, like Word), so "the peer still has an image" proves
      // nothing on its own — it has its own copy either way. Convergence is
      // the claim that matters.
      const pageBox = await a.locator(PAGE).first().boundingBox();
      const before = (await a.locator(`${PAGE} img`).first().boundingBox())!;
      await a.mouse.click(before.x + before.width / 2, before.y + before.height / 2);
      await expect(a.locator("[data-dxw-object-selection]")).toHaveCount(1);
      // Grab the CENTRE, well away from the resize handles on every edge.
      await a.mouse.move(before.x + before.width / 2, before.y + before.height / 2);
      await a.mouse.down();
      await a.mouse.move(before.x + before.width / 2 + 60, before.y + before.height / 2 + 40, { steps: 8 });
      await a.mouse.up();
      await a.waitForTimeout(700);

      expect(await imgCount(a), "the image must not vanish when moved").toBe(1);
      expect(await skeletonCount(a), "and must not degrade into a skeleton").toBe(0);
      const after = await a.locator(`${PAGE} img`).first().boundingBox();
      // "Vanished" can also mean "rendered off the sheet", which a count misses.
      expect(
        after!.x + after!.width > pageBox!.x && after!.x < pageBox!.x + pageBox!.width &&
        after!.y + after!.height > pageBox!.y && after!.y < pageBox!.y + pageBox!.height,
        `image landed off the page sheet: img=${JSON.stringify(after)} page=${JSON.stringify(pageBox)}`,
      ).toBe(true);

      // The move was SEQUENCED, not a local-only mutation that a later
      // reconciliation would silently discard.
      const kinds = await a.evaluate(() =>
        ((window as unknown as { __ww: { _session: { activity: { kind: string }[] } } }).__ww
          ._session.activity ?? []).map((x) => x.kind));
      expect(kinds, "the move must ride the wire").toContain("setFloatingPagePosition");

      // The peer keeps the image AND agrees byte-for-byte after the move.
      await expect.poll(() => imgCount(b), { message: "the peer must still have the image" }).toBe(1);
      await expect
        .poll(async () => {
          const [x, y] = await Promise.all([
            a.evaluate(() => (window as unknown as { __ww: { saveB64(): string | null } }).__ww.saveB64()),
            b.evaluate(() => (window as unknown as { __ww: { saveB64(): string | null } }).__ww.saveB64()),
          ]);
          return x !== null && x === y;
        }, { message: "placer and peer must converge after the move", timeout: 8000 })
        .toBe(true);
    } finally {
      await ctxA.close();
      await ctxB.close();
    }
  });

  test("a mis-grabbed resize handle cannot collapse the photo to a sliver", async ({ page }) => {
    // THE ACTUAL USER BUG. The handles sit on the image's edges — where you
    // aim to drag a picture — so a mis-grab resized instead of moving, and
    // one gesture produced 624x8 (measured here before the fix). Undo could
    // not recover it: resizeDrawing has no inverse. The floor is what keeps
    // the object grabbable, and therefore recoverable by hand.
    await page.goto(LANDING);
    await expect(page.getByTestId("local-editor")).toBeVisible();
    const start = await page.locator(PAGE).first().boundingBox();
    await page.mouse.click(start!.x + 30, start!.y + 25);
    await page.keyboard.type("test", { delay: 10 });
    await insertPhotoLocally(page);
    await expect.poll(() => imgCount(page)).toBe(1);
    await page.getByTestId("make-collaborative").click();
    await expect(page.getByTestId("collab-modal")).toBeVisible();
    await page.getByTestId("share-code").fill(BOARD_CODE);
    await page.getByTestId("start-collab").click();
    await expect(page.getByTestId("download")).toBeVisible();
    await waitHook(page);

    const before = (await page.locator(`${PAGE} img`).first().boundingBox())!;
    await page.mouse.click(before.x + before.width / 2, before.y + before.height / 2);
    await expect(page.locator("[data-dxw-object-selection]")).toHaveCount(1);
    // Grab the bottom edge and yank it up past the top of the image.
    await page.mouse.move(before.x + before.width / 2, before.y + before.height);
    await page.mouse.down();
    await page.mouse.move(before.x + before.width / 2, before.y - 400, { steps: 12 });
    await page.mouse.up();
    await page.waitForTimeout(700);

    const after = (await page.locator(`${PAGE} img`).first().boundingBox())!;
    // Three handle-widths (MIN_DRAG_SIZE_PX); a hair under to absorb rounding
    // and any page zoom in the harness.
    expect(after.height, `collapsed to a sliver: ${JSON.stringify(after)}`).toBeGreaterThan(20);
    expect(await imgCount(page)).toBe(1);
  });
});
