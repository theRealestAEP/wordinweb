import { test, expect } from "@playwright/test";
import { createCollabDoc, joinCollab, converge, docText, typeInEditor, waitHook, PAGE } from "./_helpers.js";

/**
 * USER-REPORTED (2026-07-24, live demo): "make a doc collaborative and then
 * another incognito window starts typing first — it desyncs". The joiner had
 * EXTRA local-only paragraphs the owner never received.
 *
 * Two variants: the joiner typing after the session is fully ready
 * (deterministic ordering pin), and the racy real-user shape — typing the
 * moment the page paints, potentially before the session hook is live.
 */
test.describe("joiner types first after make-collaborative", () => {
  test("joiner types FIRST (after ready), then owner — byte-identical", async ({ browser }) => {
    const ctxA = await browser.newContext();
    const owner = await ctxA.newPage();
    const url = await createCollabDoc(owner);

    const ctxB = await browser.newContext(); // fresh context = incognito
    const joiner = await ctxB.newPage();
    await joinCollab(joiner, url);

    // The JOINER types before the owner has ever typed.
    await typeInEditor(joiner, "kkqwerqwe");
    await joiner.keyboard.press("Enter");
    await joiner.keyboard.type("kj", { delay: 15 });

    await converge([owner, joiner], "joiner-first initial burst");
    expect(await docText(owner)).toContain("kkqwerqwe");
    expect(await docText(owner)).toContain("kj");

    // Then the owner types too; still byte-identical.
    await typeInEditor(owner, "jjjjqwrqwe");
    await converge([owner, joiner], "owner follow-up");
    expect(await docText(joiner)).toContain("jjjjqwrqwe");

    await ctxA.close();
    await ctxB.close();
  });

  test("LOCAL CONTENT before go-live, then joiner types first (user's exact flow)", async ({ browser }) => {
    const ctxA = await browser.newContext();
    const owner = await ctxA.newPage();
    // Type multi-paragraph content in the LOCAL editor BEFORE going live —
    // the sealed bytes then contain paragraphs created by local splits.
    await owner.goto("/?server=localhost:1399");
    await expect(owner.getByTestId("local-editor")).toBeVisible();
    await typeInEditor(owner, "jjjjqwrqwe");
    await owner.keyboard.press("Enter");
    await owner.keyboard.type("kj", { delay: 15 });
    await owner.keyboard.press("Enter");
    await owner.keyboard.type("qweqwqrwe", { delay: 15 });
    const { goLive } = await import("./_helpers.js");
    const url = await goLive(owner);

    const ctxB = await browser.newContext();
    const joiner = await ctxB.newPage();
    await joinCollab(joiner, url);

    // Joiner clicks into the content and types FIRST — including an Enter
    // (the user's divergence showed an extra blank paragraph + a line the
    // owner never received).
    await typeInEditor(joiner, "kk");
    await joiner.keyboard.press("Enter");
    await joiner.keyboard.type("kkqwerqwe", { delay: 15 });

    await converge([owner, joiner], "content-seed joiner-first");
    expect(await docText(owner)).toContain("kkqwerqwe");

    await typeInEditor(owner, "zz");
    await converge([owner, joiner], "owner reply");
    expect(await docText(joiner)).toContain("zz");

    await ctxA.close();
    await ctxB.close();
  });

  test("joiner's FIRST action is click-below-content-and-type (whitespace extender)", async ({ browser }) => {
    const ctxA = await browser.newContext();
    const owner = await ctxA.newPage();
    await owner.goto("/?server=localhost:1399");
    await expect(owner.getByTestId("local-editor")).toBeVisible();
    await typeInEditor(owner, "jjjjqwrqwe");
    await owner.keyboard.press("Enter");
    await owner.keyboard.type("kj", { delay: 15 });
    await owner.keyboard.press("Enter");
    await owner.keyboard.type("qweqwqrwe", { delay: 15 });
    const { goLive } = await import("./_helpers.js");
    const url = await goLive(owner);

    const ctxB = await browser.newContext();
    const joiner = await ctxB.newPage();
    await joinCollab(joiner, url);

    // The user's divergence shape: an EMPTY paragraph then a typed line,
    // present only on the joiner. That is the click-and-type whitespace
    // extender — click well BELOW the last text line, then type.
    const box = await joiner.locator(PAGE).first().boundingBox();
    if (!box) throw new Error("no page box");
    await joiner.mouse.click(box.x + box.width / 2, box.y + 220); // open whitespace below 3 short lines
    await joiner.keyboard.type("kkqwerqwe", { delay: 15 });

    await converge([owner, joiner], "joiner click-below-and-type");
    expect(await docText(owner)).toContain("kkqwerqwe");

    // Room still healthy both directions.
    await typeInEditor(owner, "zz");
    await converge([owner, joiner], "owner after extender");
    expect(await docText(joiner)).toContain("zz");

    await ctxA.close();
    await ctxB.close();
  });

  test("joiner types IMMEDIATELY at first paint (before ready-wait) — no silent local fork", async ({ browser }) => {
    const ctxA = await browser.newContext();
    const owner = await ctxA.newPage();
    const url = await createCollabDoc(owner);

    const ctxB = await browser.newContext();
    const joiner = await ctxB.newPage();
    // Deliberately NOT joinCollab(): race the session — type as soon as the
    // page paints, without waiting for the __ww hook / ready.
    await joiner.goto(url);
    await expect(joiner.locator(PAGE)).toBeVisible();
    await typeInEditor(joiner, "early");
    await joiner.keyboard.press("Enter");
    await joiner.keyboard.type("bird", { delay: 15 });

    // Now let the session settle and check nothing forked: whatever of the
    // early typing survived must be on BOTH sides, byte-identically.
    await waitHook(joiner);
    await converge([owner, joiner], "type-before-ready settle");

    // And the room still works normally afterwards.
    await typeInEditor(owner, "later");
    await converge([owner, joiner], "post-race sanity");
    expect(await docText(joiner)).toContain("later");

    await ctxA.close();
    await ctxB.close();
  });
});
